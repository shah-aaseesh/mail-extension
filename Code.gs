/**
 * Gmail Domain Labeler — labels every email by its sender's domain.
 *
 * 1. scan()  → lists every sender domain in the "Domains" sheet with a proposed label
 * 2. Edit the Label column if you like (clear a cell to skip that domain)
 * 3. apply() → labels past emails and creates a Gmail filter per domain for future mail
 *
 * Labels are only ever added — nothing is archived, moved or deleted.
 */

const SHEET_NAME = 'Domains';
const HEADERS = ['Domain', 'Emails', 'Label (edit, or clear to skip)', 'Status', 'Example senders'];
const SCAN_QUERY = '-in:sent -in:drafts -in:chats';
const MAX_RUNTIME_MS = 5 * 60 * 1000; // Apps Script kills runs at 6 min
const FETCH_BATCH = 40;               // parallel requests, kept under Gmail's per-user rate limit

// Suffixes where the registrable domain has three parts (e.g. amazon.co.in, not co.in).
const MULTI_PART_SUFFIXES = new Set([
  'co.in', 'net.in', 'org.in', 'firm.in', 'gen.in', 'ind.in', 'ac.in', 'edu.in', 'gov.in', 'res.in',
  'co.uk', 'org.uk', 'ac.uk', 'gov.uk', 'com.au', 'net.au', 'org.au', 'co.nz', 'co.za', 'co.jp',
  'com.sg', 'com.my', 'com.br', 'com.mx', 'com.cn', 'com.hk', 'com.tw', 'co.kr', 'com.tr', 'co.id'
]);

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Domain Labeler')
    .addItem('1. Scan mailbox', 'scan')
    .addItem('2. Apply labels + filters', 'apply')
    .addToUi();
}

// ---------- Step 1: scan ----------

function scan() {
  const ids = listMessageIds_(SCAN_QUERY);
  const domains = {}; // registrable domain → { count, senders }
  fetchSenders_(ids).forEach(addr => {
    const domain = registrableDomain_(addr.split('@')[1]);
    if (!domain) return;
    const d = domains[domain] || (domains[domain] = { count: 0, senders: new Set() });
    d.count++;
    d.senders.add(addr);
  });

  const proposed = proposeLabels_(Object.keys(domains));
  const existing = readSheet_(); // keep your edits and progress across re-scans
  const rows = Object.keys(domains)
    .sort((a, b) => domains[b].count - domains[a].count)
    .map(domain => {
      const prev = existing[domain];
      return [
        domain,
        domains[domain].count,
        prev ? prev.label : proposed[domain],
        prev ? prev.status : '',
        [...domains[domain].senders].slice(0, 3).join(', ')
      ];
    });

  writeSheet_(rows);
  notify_(`Found ${rows.length} domains across ${ids.length} emails. Review the Label column, then run Apply.`);
}

// One label per domain; when an organisation sends from several domains, nest them.
//   swiggy.in only            → "swiggy.in"
//   amazon.in + amazon.com    → "amazon/amazon.in", "amazon/amazon.com"
function proposeLabels_(domainList) {
  const byOrg = {};
  domainList.forEach(d => {
    const org = d.split('.')[0];
    (byOrg[org] = byOrg[org] || []).push(d);
  });
  const labels = {};
  Object.keys(byOrg).forEach(org => {
    const group = byOrg[org];
    group.forEach(d => { labels[d] = group.length > 1 ? `${org}/${d}` : d; });
  });
  return labels;
}

// ---------- Step 2: apply ----------

function apply() {
  const start = Date.now();
  const sheet = getSheet_();
  const rows = sheet.getDataRange().getValues();
  const labelIds = existingLabels_();
  const filters = existingFilters_();
  let done = 0;
  let remaining = 0;

  for (let i = 1; i < rows.length; i++) {
    const domain = String(rows[i][0]).trim().toLowerCase();
    const label = String(rows[i][2]).trim();
    if (!domain || !label || rows[i][3] === 'Done') continue;
    if (Date.now() - start > MAX_RUNTIME_MS) { remaining++; continue; }

    const labelId = ensureLabel_(label, labelIds);
    const ids = listMessageIds_(`from:${domain} ${SCAN_QUERY}`);
    for (let j = 0; j < ids.length; j += 1000) {
      Gmail.Users.Messages.batchModify({ ids: ids.slice(j, j + 1000), addLabelIds: [labelId] }, 'me');
    }

    if (!filters.has(`${domain}|${labelId}`)) {
      try {
        Gmail.Users.Settings.Filters.create(
          { criteria: { from: domain }, action: { addLabelIds: [labelId] } }, 'me');
      } catch (e) {
        if (!/exists/i.test(e.message)) throw e;
      }
    }

    sheet.getRange(i + 1, 4).setValue('Done');
    done++;
  }

  notify_(remaining
    ? `Labeled ${done} domains. ${remaining} left — run Apply again to continue.`
    : `Labeled ${done} domains. All done; new mail will be labeled by Gmail filters.`);
}

// Creates "a", then "a/b", etc. Gmail label names are case-insensitive.
function ensureLabel_(name, labelIds) {
  const parts = name.split('/');
  for (let k = 1; k <= parts.length; k++) {
    const path = parts.slice(0, k).join('/');
    if (!labelIds[path.toLowerCase()]) {
      labelIds[path.toLowerCase()] = Gmail.Users.Labels.create(
        { name: path, labelListVisibility: 'labelShow', messageListVisibility: 'show' }, 'me').id;
    }
  }
  return labelIds[name.toLowerCase()];
}

function existingLabels_() {
  const map = {};
  (Gmail.Users.Labels.list('me').labels || []).forEach(l => { map[l.name.toLowerCase()] = l.id; });
  return map;
}

function existingFilters_() {
  const set = new Set();
  (Gmail.Users.Settings.Filters.list('me').filter || []).forEach(f => {
    const from = f.criteria && f.criteria.from;
    if (!from) return;
    ((f.action && f.action.addLabelIds) || []).forEach(id => set.add(`${from.toLowerCase()}|${id}`));
  });
  return set;
}

// ---------- Gmail helpers ----------

function listMessageIds_(query) {
  const ids = [];
  let pageToken;
  do {
    const res = Gmail.Users.Messages.list('me', { q: query, maxResults: 500, pageToken });
    (res.messages || []).forEach(m => ids.push(m.id));
    pageToken = res.nextPageToken;
  } while (pageToken);
  return ids;
}

// Reads only the From header, many messages in parallel, retrying rate-limited ones.
function fetchSenders_(ids) {
  const token = ScriptApp.getOAuthToken();
  const base = 'https://gmail.googleapis.com/gmail/v1/users/me/messages/';
  const senders = [];

  for (let i = 0; i < ids.length; i += FETCH_BATCH) {
    let batch = ids.slice(i, i + FETCH_BATCH);
    for (let attempt = 0; batch.length && attempt < 5; attempt++) {
      if (attempt) Utilities.sleep(2000 * attempt);
      const responses = UrlFetchApp.fetchAll(batch.map(id => ({
        url: `${base}${id}?format=metadata&metadataHeaders=From`,
        headers: { Authorization: `Bearer ${token}` },
        muteHttpExceptions: true
      })));
      const retry = [];
      responses.forEach((r, j) => {
        const code = r.getResponseCode();
        if (code === 200) {
          const payload = JSON.parse(r.getContentText()).payload || {};
          const from = (payload.headers || []).find(h => h.name.toLowerCase() === 'from');
          const addr = from && parseAddress_(from.value);
          if (addr) senders.push(addr);
        } else if (code === 429 || code >= 500) {
          retry.push(batch[j]);
        }
      });
      batch = retry;
    }
    Utilities.sleep(800);
  }
  return senders;
}

// '"Amazon.in" <order-update@amazon.in>' → 'order-update@amazon.in'
function parseAddress_(from) {
  const m = from.match(/<([^<>\s]+@[^<>\s]+)>/) || from.match(/([^\s<>"]+@[^\s<>"]+)/);
  return m ? m[1].toLowerCase() : null;
}

// 'alerts.hdfcbank.net' → 'hdfcbank.net', 'mail.amazon.co.in' → 'amazon.co.in'
function registrableDomain_(host) {
  if (!host) return null;
  const parts = host.toLowerCase().replace(/\.$/, '').split('.');
  if (parts.length < 2) return null;
  const n = MULTI_PART_SUFFIXES.has(parts.slice(-2).join('.')) ? 3 : 2;
  return parts.slice(-n).join('.');
}

// ---------- Sheet helpers ----------

function getSheet_() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function readSheet_() {
  const map = {};
  getSheet_().getDataRange().getValues().slice(1).forEach(r => {
    if (r[0]) map[String(r[0]).toLowerCase()] = { label: String(r[2]).trim(), status: r[3] };
  });
  return map;
}

function writeSheet_(rows) {
  const sheet = getSheet_();
  sheet.clear();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  if (rows.length) sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, HEADERS.length);
}

function notify_(msg) {
  Logger.log(msg);
  SpreadsheetApp.getActive().toast(msg, 'Domain Labeler', 10);
}
