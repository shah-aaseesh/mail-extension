# Gmail Domain Labeler

Personal Google Apps Script that labels every email by its sender's domain and creates
Gmail filters so future mail gets labeled automatically. It only adds labels; nothing is archived.

## Label rules
- One label per domain, with subdomains collapsed: `alerts.hdfcbank.net` → `hdfcbank.net`
- When one organisation uses several domains, they're nested: `amazon/amazon.in`, `amazon/amazon.com`
- Sent mail, drafts and chats are ignored

## Setup (one time)
1. Create a new Google Sheet (any name).
2. **Extensions → Apps Script**.
3. Replace the contents of `Code.gs` with this folder's `Code.gs`.
4. **Project Settings (gear) → tick "Show appsscript.json manifest file"**, then go back to the editor
   and replace `appsscript.json` with this folder's version. Save.
5. Reload the Sheet. A **Domain Labeler** menu appears.

## Use
1. **Domain Labeler → 1. Scan mailbox.** The first time, Google asks you to authorize the script. It will say
   "Google hasn't verified this app"; that's expected for your own script. Click **Advanced → Go to … (unsafe)**.
   A ~4,000-email mailbox takes about 3 minutes.
2. Review the **Domains** sheet. Edit labels, or clear a Label cell to skip that domain
   (e.g. `gmail.com`, which would otherwise collect all your personal mail).
3. **Domain Labeler → 2. Apply labels + filters.** If it says some domains are left, run it again;
   finished rows are marked `Done` and skipped.

Re-run **Scan** later to pick up new senders. Your edits and `Done` statuses are kept.

## Undo
- Filters: Gmail → Settings → Filters and Blocked Addresses
- Labels: deleting a label in Gmail removes it from all emails (the emails themselves stay)
