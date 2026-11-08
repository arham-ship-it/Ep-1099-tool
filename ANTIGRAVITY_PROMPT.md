# Prompts for Antigravity's agent

Open the `ep-1099-tool` folder in Antigravity, then paste one prompt at a time into the Agent panel. Review each change and run `npm test` before moving on.

---

## 0. Get oriented (paste first)

```
This repo is the EP 1099 Tool: a Vite + vanilla JS web app that reads a general ledger,
applies IRS 1099-NEC/MISC rules (src/logic.js), builds a color-coded tracker per company,
reads W-9s (src/integrations.js), and sends vendors to a Google Sheet email tracker
(apps-script/EmailTracker.gs) that creates Gmail drafts and follow-ups.
Read README.md and every file in src/ and apps-script/. Run npm install, npm test, npm run dev.
Summarize the architecture back to me and list anything that looks broken. Do not change code yet.
```

## 1. Shared team storage (Firebase) — do this before real client data

```
Replace src/store.js (browser localStorage) with Firebase Firestore + Firebase Auth (Google sign-in),
keeping the exact same Store method names so main.js barely changes.
- Collections: clients/{cid}, vendors/{cid} (one doc per company; vendor fields keyed v_<key>),
  settings/main.
- Only signed-in users from our Google Workspace domain can read/write (Firestore security rules).
- Live updates with onSnapshot so two teammates see the same tracker.
- Keep Back up / Restore working.
- Put Firebase config in .env (VITE_FIREBASE_*), add .env.example, never commit secrets.
Show me the security rules file before deploying.
```

## 2. Deploy it as a website

```
Deploy this app to Firebase Hosting (or Vercel). Add the production URL to the Google OAuth
client's Authorized JavaScript origins (tell me exactly what to add). Add an npm script "deploy".
```

## 3. Client-facing share link (read-only)

```
Add a read-only client view at /client/<cid>?t=<token>: it shows only the KPI tiles, progress bar
and vendor table (no contacts, no edit controls). Generate a random token per client stored on
clients/{cid}, a "Copy client link" button on the tracker, and a "Revoke link" button.
Firestore rules must allow reading only that client's docs with a matching token.
```

## 4. Two-way sync with the email tracker Sheet

```
Today the app sends vendors to the Sheet but doesn't hear back. Add a doGet action
?k=<key>&action=status&company=<name> in apps-script/EmailTracker.gs that returns JSON
(vendor, Request Status, Sent Count, Last Contact, Vendor Replied, W-9 Received).
In the app, add "Refresh email status" on each tracker that calls it (JSONP or a CORS-safe
approach that works with Apps Script) and shows Draft / Sent / Follow-up due / Replied per vendor.
```

## 5. Store W-9 PDFs

```
When a W-9 is saved, upload the file to Firebase Storage at w9/<cid>/<vendorKey>/<filename>
(or copy it into our Drive folder via the Drive API) and show a "View W-9" link in the table.
Restrict access to signed-in domain users.
```

## 6. Year-end export for filing

```
Add "Export for e-filing" per company: a CSV with payer info (company name, EIN, address — add
these fields to the client settings), recipient legal name, TIN last 4 (full TIN entered by the
user at export time, never stored), address, form, box and amount, in the column layout the
IRS IRIS portal's CSV template expects. Add tests.
```

## Rules for the agent (paste with any prompt)

```
- Never store full TINs or SSNs anywhere — last 4 digits only.
- Keep IRS thresholds in src/logic.js RULES; add tests in tests/logic.test.mjs for any rule change.
- Every vendor result must keep its plain-English "why" reason.
- Emails are only ever created as Gmail DRAFTS; nothing sends automatically.
- Run npm test and npm run build after every change and show me the output.
```
