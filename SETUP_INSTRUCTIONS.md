# Cadence — Google Sheets + Apps Script setup

## What this gives you
- Your tasks live as rows in a Google Sheet (tabs: **Daily**, **Weekly**, **Monthly**, **Settings**), so you can edit them directly in the Sheet or through the app.
- Every task has an **Urgency**: High / Medium / Low.
- A **daily email digest** lists everything still open, sorted by urgency, sent at whatever hour you choose.
- A small web app (styled like your original Cadence design) to add/check off/delete tasks from your phone or browser.

## Setup (about 5 minutes)

1. **Create a new Google Sheet** (sheets.new).
2. Go to **Extensions → Apps Script**. Delete the default `Code.gs` content.
3. Copy the contents of **Code.gs** (from this download) into that editor.
4. In the Apps Script editor, click the **+** next to Files → **HTML** → name it exactly `Index` (it will save as `Index.html`). Paste in the contents of **Index.html** from this download.
5. Click **Save** (the disk icon), then go back to your Sheet and **reload the page**. You should see a new **Cadence** menu appear next to Extensions.
6. Click **Cadence → Send reminder email now** once — this will trigger Google's permission prompt. Approve access (it's your own script, running on your own account — the warning screen is normal for personal Apps Script projects). This also auto-creates the Daily/Weekly/Monthly/Settings sheets with headers.

## Turn on the web app UI

1. In the Apps Script editor: **Deploy → New deployment**.
2. Click the gear icon next to "Select type" → choose **Web app**.
3. Set:
   - Execute as: **Me**
   - Who has access: **Only myself** (or "Anyone with the link" if you want to open it from your phone without being logged in as the script owner — not recommended unless you understand the tradeoff)
4. Click **Deploy**, authorize again if asked, then copy the **Web app URL**. Open it — that's your Cadence app, now backed by the Sheet.
5. Bookmark that URL on your phone's home screen for quick access.

## Turn on daily email reminders

Either:
- In the web app, go to the **Reminders** tab, set your email + hour, click **Save & schedule**, or
- In the Sheet, use **Cadence → Set up daily reminder trigger…** from the menu.

This creates a time-based trigger that runs `sendDailyDigest()` once a day. The email is skipped automatically on any day with zero open tasks, so you won't get spammed when you're all caught up.

## Editing tasks directly in the Sheet
You can also just type rows into the Daily/Weekly/Monthly tabs by hand — columns are `ID | Task | Urgency | Completed | CreatedAt | ReminderSentAt`. Leave ID blank and the app will still read the row, but toggling/deleting from the web app requires an ID, so it's easiest to add tasks through the app or fill in a unique ID yourself (e.g. `t1`, `t2`...).

## Notes / things you might want to extend later
- This version drops the streaks/rings dashboard and week/month auto-rollover from your original HTML, to keep the Sheet-based version simple and reliable. Both could be added back (e.g. a "History" sheet + a nightly trigger that archives and clears Weekly/Monthly rows) — happy to build that next if you want it.
- `MailApp.sendEmail` uses your own Gmail quota (100/day on a free account, much higher on Workspace), so a once-a-day digest is well within limits.
