# MyBudjet Dashboard (GitHub Pages)

A static page (`index.html`) hosted on GitHub Pages. A page like this cannot read or write a Google Sheet
by itself, so a small Apps Script web app (`Code.gs`) acts as the backend. The page calls it with `fetch`.

```
GitHub Pages (index.html)  --fetch + login session-->  Apps Script web app (Code.gs)  -->  Google Sheet
```

## Files
- `index.html` - the whole UI: Records, Accounts & verification, Categorize, Add entry (SMS paste autofill)
- `Code.gs` - backend, pasted into Apps Script (does not need to be in the repo, but can be)

## 1. Backend (Apps Script)
1. Open the spreadsheet: https://docs.google.com/spreadsheets/d/1AWH16EV239lPmdIAXTdz46d38uF0s9dJdvf-okukz8M/edit
2. Extensions > Apps Script. Replace the contents of `Code.gs` with the provided `Code.gs`.
3. Create your login: in `Code.gs` open the function `addUserNow`, type a username and a password (at least 8 characters) between the quotes, choose `addUserNow` in the function dropdown and press **Run** (accept the permissions). Then **erase the password from the code again**. Run it again with another username to add more people, or with the same username to reset a password.
4. Deploy > New deployment > Web app. Execute as: **Me**. Who has access: **Anyone**. Authorize when asked.
5. Copy the web app URL (ends with `/exec`).
6. After any later change to `Code.gs`: Deploy > Manage deployments > edit > New version.
7. The old `API_TOKEN` script property is no longer used. You can delete it.

Password rules: a password expires 90 days after it was set. During the last 7 days the page shows a reminder banner every day (**Remind me tomorrow** hides it until the next day). After expiry, signing in forces a new password. **Change password** in the header changes it at any time, starts a fresh 90-day period, resets the reminder and signs out the other devices. Running `addUserNow` for a user also starts a fresh 90-day period.

Other helpers (run from the editor): `listUsersNow` shows who can sign in, `removeUserNow` removes a user (type the name inside it first).

## 2. Frontend (GitHub)
1. Open `index.html` and set `API_URL` near the top of the script to the URL from step 5.
2. Create a repo, add `index.html` (and `README.md`, `Code.gs`), push.
3. Settings > Pages > Deploy from branch > `main` / root.
4. Open `https://<user>.github.io/<repo>/`. On first load it shows a sign-in screen. After signing in, the session is kept in that browser for 30 days. Use **Sign out** in the header to end it.

## Data
- The dashboard reads the `Entries` sheet. On first load it is created automatically: copied from `Responses` if that sheet exists, otherwise created with the standard headers.
- Category / Sub-category columns are added automatically. Opening balances and category lists live in Script Properties (`dashCfg`).

## Security notes
- Anyone can open the public page, but data only loads after a valid login. Passwords are never stored: only salted hashes in the `AUTH_USERS` script property.
- After 5 wrong passwords a username is locked for 15 minutes. Use passwords of 12+ characters.
- Never leave a password typed inside `addUserNow` or commit it to the repo.
- Removing a user (or deleting the `AUTH_SECRET` script property) signs everyone out immediately.
- `parseSms` in `index.html` has bank account last-digit patterns and `PEOPLE` has names. Use a private repo (Pages on private repos needs a paid GitHub plan) or replace them with placeholders.
- The page has `noindex`, which only asks search engines to skip it.
