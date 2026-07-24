# AutoNudge (Slack)

Dedicated Slack App for **Year-Round Workshop Scheduling** automation messaging (separate from Ops Bot).

## Create via Playwright (preferred)

Workspace app creation automatically queues an admin/IT approval request — no separate IT ticket.

```bash
# Optional: provide Google Workspace password for aaron.bunch@varsitytutors.com
export VT_SSO_PASSWORD='...'   # or SLACK_PASSWORD / GOOGLE_PASSWORD

npm install
npx playwright install chromium   # only needed if system Chrome unavailable
npm run create-autonudge
```

Script: [`create-autonudge-playwright.js`](create-autonudge-playwright.js)

Flow automated:

1. Open `varsitytutors.slack.com`
2. Click **Google** SSO
3. Enter `aaron.bunch@varsitytutors.com`
4. Complete Google password (requires `VT_SSO_PASSWORD` in the environment)
5. Open `api.slack.com/apps` → **Create New App** → **From a manifest**
6. Paste [`autonudge.manifest.yaml`](autonudge.manifest.yaml) and create
7. Slack notifies workspace admins/IT for approval when required

### Current cloud-agent run status

Playwright reached Google’s password challenge for `aaron.bunch@varsitytutors.com`. This environment has **no** `VT_SSO_PASSWORD` / `SLACK_PASSWORD` / `GOOGLE_PASSWORD` secret, so app creation could not be submitted yet.

Artifacts from the attempt: `/opt/cursor/artifacts/autonudge-slack-create/` (screenshots + `result.json`).

Once the SSO password secret is available to the agent, re-run `npm run create-autonudge` to finish create + auto IT request with no further manual IT ticket.

## Artifacts

| File | Purpose |
| --- | --- |
| [`autonudge.manifest.yaml`](autonudge.manifest.yaml) | Slack App Manifest (DM scopes) |
| [`create-autonudge-playwright.js`](create-autonudge-playwright.js) | Browser automation to create the app |
| [`SETUP_CHECKLIST.md`](SETUP_CHECKLIST.md) | Post-approval token storage / smoke test |
| [`IT_PERMISSION_REQUEST.md`](IT_PERMISSION_REQUEST.md) | Reference scope justification (not required if create triggers IT) |

## Scopes (v1)

- `chat:write`
- `im:write`
- `users:read`
- `users:read.email`
