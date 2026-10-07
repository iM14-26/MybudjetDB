# MyBudjet Dashboard (GitHub Pages)

A static page (`index.html`) hosted on GitHub Pages. A page like this cannot read or write a Google Sheet
by itself, so a small Apps Script web app (`Code.gs`) acts as the backend. The page calls it with `fetch`.

```
GitHub Pages (index.html)  --fetch + token-->  Apps Script web app (Code.gs)  -->  Google Sheet
```

## Files
- `index.html` - the whole UI: Records, Accounts & verification, Categorize, Add entry (SMS paste autofill)
- `Code.gs` - backend, pasted into Apps Script (does not need to be in the repo, but can be)

## 1. Backend (Apps Script)
1. Open the spreadsheet: https://docs.google.com/spreadsheets/d/1AWH16EV239lPmdIAXTdz46d38uF0s9dJdvf-okukz8M/edit
2. Extensions > Apps Script. Replace the contents of `Code.gs` with the provided `Code.gs`.
3. Project Settings > Script properties > Add property: `API_TOKEN` = a long random secret (you will type it once in the browser).
4. Deploy > New deployment > Web app. Execute as: **Me**. Who has access: **Anyone**. Authorize when asked.
5. Copy the web app URL (ends with `/exec`).
6. After any later change to `Code.gs`: Deploy > Manage deployments > edit > New version.

## 2. Frontend (GitHub)
1. Open `index.html` and set `API_URL` near the top of the script to the URL from step 5.
2. Create a repo, add `index.html` (and `README.md`, `Code.gs`), push.
3. Settings > Pages > Deploy from branch > `main` / root.
4. Open `https://<user>.github.io/<repo>/`. On first load it asks for the token; it is kept in that browser's localStorage.

## Data
- The dashboard reads the `Entries` sheet. On first load it is created automatically: copied from `Responses` if that sheet exists, otherwise created with the standard headers.
- Category / Sub-category columns are added automatically. Opening balances and category lists live in Script Properties (`dashCfg`).

## Security notes
- Anyone can open the public page, but the data only loads with the correct `API_TOKEN`. Keep the token out of the repo.
- Anyone who finds the web app URL can try tokens, so use a long random one. Rotate it by changing the script property.
- `parseSms` in `index.html` has bank account last-digit patterns and `PEOPLE` has names. Use a private repo (Pages on private repos needs a paid GitHub plan) or replace them with placeholders.
- The page has `noindex`, which only asks search engines to skip it.
