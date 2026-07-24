# pGC Coaching Effectivity Bank

Netlify-hosted **pGC Looker bank** that monitors coaching dates and measures coaching effectivity from Looker pGC:

| Window | Definition |
|--------|------------|
| **L7** | Mean pGC over the 7 calendar days **before** coaching day (day excluded) |
| **N7** | Mean pGC over the 7 calendar days **after** coaching day (day excluded) |
| **Delta** | `N7 − L7` |
| **Effective** | `delta > 0` |
| **Poor / neutral** | `delta ≤ 0` |
| **Ready** | Scored only after `coaching_date + 7` days have fully elapsed |

## Architecture

```
Google Sheet (coaching events)
        │
        ▼
Netlify scheduled sync (@daily) ──► Looker API (daily pGC)
        │
        ▼
   Score L7 / N7 / delta / verdict
        │
        ├──► Bank store (Netlify Blobs or data/results.json)
        ├──► Optional sheet writeback columns
        └──► Dashboard + /api/results
```

Coaching sheet (configured tab, default `Coaching`):

| Coachee | Coachee Email | Coach | Coaching Date | pGC L7 | pGC N7 | pGC Delta | Verdict | Scored At |
|---------|---------------|-------|---------------|--------|--------|-----------|---------|-----------|

Header names are configurable via env (`SHEET_COL_*`). Writeback columns are created automatically when missing.

## Quick start (demo, no credentials)

```bash
npm install
PGC_MODE=demo npm run sync
npm test
```

Then open the dashboard with `npx netlify dev` (or deploy to Netlify) and hit **Run sync**.

Demo fixtures score:

- **Alex / Jordan** → effective (pGC lift)
- **Sam / Jordan** → poor (pGC drop)
- **Casey / Morgan** → pending (N7 not complete as of 2026-07-24)

## Live credentials

Copy [`.env.example`](.env.example) into Netlify site env vars (or `.env` for local):

1. **Google Sheets** — share the coaching spreadsheet with a service account; set `GOOGLE_SERVICE_ACCOUNT_EMAIL` and `GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY` (use `\n` for newlines in env). Spreadsheet ID defaults to the coaching sheet in the plan.
2. **Looker** — API3 `LOOKER_CLIENT_ID` / `LOOKER_CLIENT_SECRET` / `LOOKER_BASE_URL`, plus either:
   - `LOOKER_MODEL` + `LOOKER_EXPLORE` + field names (`LOOKER_PERSON_FIELD`, `LOOKER_DATE_FIELD`, `LOOKER_PGC_FIELD`), or
   - `LOOKER_LOOK_ID` for an existing Look that returns person, date, pGC.
3. Set `PGC_MODE=live` and `SHEET_WRITEBACK=true` (default) to patch scores back onto the sheet.

Join key: sheet **Coachee Email** (fallback: Coachee name) must match Looker `LOOKER_PERSON_FIELD`.

## API

| Endpoint | Method | Purpose |
|----------|--------|---------|
| `/api/health` | GET | Liveness + mode |
| `/api/sync` | GET/POST | Run scoring sync (`?mode=demo` forces fixtures) |
| `/api/results` | GET | Bank JSON (`?coach=` `?verdict=`) |

Scheduled function: `sync` runs `@daily` via [`netlify.toml`](netlify.toml).

## Deploy

1. Connect this repo to Netlify.
2. Set env vars from `.env.example`.
3. Publish directory: `public` (already in `netlify.toml`).
4. Confirm scheduled functions are enabled on the site plan.

## Local scripts

```bash
PGC_MODE=demo node scripts/run-sync.js
# or
npm run sync   # uses env PGC_MODE
```
