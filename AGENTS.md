# automation

This repository is a collection of independent automation projects. Each project lives on its own `cursor/*` branch; the `main` branch is currently just a placeholder README. This branch contains the **pGC Coaching Effectivity Bank** project.

## Cursor Cloud specific instructions

### Project: pGC Coaching Effectivity Bank (this branch)

A Netlify-hosted Node.js app (ESM, `"type": "module"`, Node `>=20`) that scores coaching effectivity from Looker pGC data. See `README.md` for the full architecture and env var reference.

- Standard commands are defined in `package.json` scripts: `npm test` (node's built-in test runner), `npm run sync` (one-off scoring sync), `npm run dev` (`netlify dev`). There is no separate lint step.
- **Demo mode needs no credentials.** Set `PGC_MODE=demo` to use bundled fixtures instead of live Google Sheets / Looker. Use this for local runs and tests:
  - `PGC_MODE=demo npm run sync` scores fixtures and writes to `data/results.json`.
  - `PGC_MODE=demo npx netlify dev` serves the dashboard; without demo mode, live sync requires Google service-account + Looker API credentials (see `.env.example`).
- `netlify dev` serves the dashboard on `http://localhost:8888` and mounts the functions under `/api/*` (`/api/health`, `/api/sync`, `/api/results`) via the redirect in `netlify.toml`. First startup is slow ("Setting up the Edge Functions environment", ~20-30s). It cannot auto-open a browser in this container — open the printed URL yourself.
- **Bank storage fallback:** when Netlify Blobs isn't configured, results persist to the local `data/results.json` file (`"backend": "file"`). The dashboard's `/api/results` reads from this store, so a `sync` must run at least once before results appear.
