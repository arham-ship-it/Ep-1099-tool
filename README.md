# EP 1099 Tool

Upload a general ledger → get a color-coded 1099 tracker for **each company** → read W-9s (upload or straight from a Google Drive folder) → send vendors that still need a W-9 to a **Google Sheet email tracker**, where one click creates W-9 request drafts and follow-ups in the sender's Gmail.

```
ep-1099-tool/
├─ index.html              app shell
├─ src/
│  ├─ logic.js             IRS rules + 1099 engine (edit RULES when the law changes)
│  ├─ main.js              screens: GL import, trackers, W-9s, settings
│  ├─ integrations.js      Google Drive, Gemini (optional), Google Sheet
│  ├─ store.js             saving (browser storage; swap for Firebase to share)
│  └─ styles.css           colors and layout
├─ apps-script/
│  ├─ EmailTracker.gs      Google Sheet: drafts, follow-ups, summary, IRS check
│  └─ appsscript.json
├─ sample/sample_GL_2026.csv
├─ tests/logic.test.mjs    7 rule tests (npm test)
└─ ANTIGRAVITY_PROMPT.md   prompts to paste into Antigravity's agent
```

## 1. Run it (Antigravity or any editor)

1. Unzip, then in Antigravity: **File → Open Folder → ep-1099-tool**.
2. Open the terminal (needs Node.js 18+):
   ```
   npm install
   npm test        # should say: pass 7, fail 0
   npm run dev     # opens http://localhost:5173
   ```
3. Click **Try it with a sample GL** to see it work.

## 2. Set up the Google Sheet email tracker (once, ~10 min)

1. Create a Google Sheet: `EP 1099 – Email Tracker`.
2. **Extensions → Apps Script**. Paste `apps-script/EmailTracker.gs` into `Code.gs`.
   Project Settings (gear) → tick **Show "appsscript.json"** → paste `apps-script/appsscript.json`. Save.
3. Reload the Sheet → menu **1099 Emails → Setup → 1. Set up this sheet** → Allow. Note the **App key**.
4. Fill the Settings tab: W-9 Due Date, Sender Title, W-9 Save Folder URL.
5. Apps Script → **Deploy → New deployment → Web app**. Execute as **Me**, access **Anyone** (needed for the app's background send; the App key protects it) → Deploy → copy the URL ending in `/exec`.
6. In the app: **Settings** → paste the URL and App key → Save → **Test the connection**.
7. Share the Sheet (Editor) with every sender.

## 3. Read W-9s from a client's Drive folder (once, ~10 min)

1. Go to <https://console.cloud.google.com> → create a project → **APIs & Services → Library** → enable **Google Drive API**.
2. **OAuth consent screen** → External (or Internal for Workspace) → add your team as test users.
3. **Credentials → Create credentials → OAuth client ID → Web application**. Authorized JavaScript origins: `http://localhost:5173` (and your live site URL later).
4. Copy the Client ID into the app's **Settings → Google OAuth Client ID**.
5. On a company tracker, paste the client's Drive folder link (view-only access is enough) → **Read W-9s from Drive** → sign in with the Google account that the folder is shared with.

Optional: add a **Gemini API key** (from <https://aistudio.google.com>) in Settings to read scanned or flattened W-9s. Fillable W-9 PDFs are read from their form fields without any AI.

## 4. Each season, per client

1. **Upload general ledger** → check columns → check GL account categories (most important review) → **Run 1099 check**. One tracker per company.
2. On each tracker: Drive folder link, sender name and email.
3. **Read W-9s from Drive** or **Upload W-9s** (PDF, photo or zip) → confirm matches → Save.
4. Add vendor contact names/emails.
5. **Send N to email tracker** → vendors land in that company's tab of the Sheet.
6. Sender opens the Sheet → **1099 Emails → Create my W-9 request drafts** → reviews and sends from Gmail Drafts.
7. Follow-ups: select rows → **Follow up on selected row(s)**, or **Create all my due follow-ups**.
8. Returned W-9s → read them into the app. **Client view** shows the client only the counts and statuses.

## Colors

| Color | App | Sheet |
|---|---|---|
| Green | Eligible · W-9 received | W-9 received · Verified |
| Yellow | Needs W-9 (entity unknown) | Draft waiting |
| Orange | Needs review (unmapped account) | Needs review |
| Blue | Requested | Sent / following up |
| Red | To request | Not sent · follow-up overdue |
| Grey | Not eligible | — |

## When the law changes

The Sheet's weekly IRS check (Setup → Turn on weekly IRS rule check) emails the admin when IRS 1099 pages change. Then add a new year block to `RULES` in `src/logic.js`, add a test, run `npm test`, and redeploy.

## Good to know

- Trackers are saved in **this browser only** until you move storage to Firebase (prompt 1 in ANTIGRAVITY_PROMPT.md). Use **Back up data** often.
- Only the last 4 TIN digits are stored. W-9 files themselves stay in Drive.
- Not tax advice — have your CPA sign off on the rules and GL account mapping.
