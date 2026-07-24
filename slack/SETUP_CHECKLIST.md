# AutoNudge — Post-IT Setup Checklist

Use this after IT creates/installs AutoNudge and returns the Bot User OAuth Token.

## 1. Confirm the app in Slack

- [ ] Search Slack for **AutoNudge** (Apps / Direct Messages)
- [ ] Confirm display name is **AutoNudge** (not Ops Bot)
- [ ] Do **not** remove or rotate Ops Bot until cutover is intentional

## 2. Store the bot token securely

Preferred property name for a clean cutover later:

- [ ] Add `AUTONUDGE_SLACK_BOT_TOKEN` = `xoxb-…` in Year-Round Workshop Apps Script **Script Properties** (and/or TS Config if that is where tokens are managed)

Until Apps Script is updated to read the AutoNudge-specific property, the existing scheduler still uses:

- `SLACK_BOT_TOKEN` (today: Ops Bot)

**Cutover options (choose one when ready):**

| Approach | Action |
| --- | --- |
| A — Swap in place | Replace the value of `SLACK_BOT_TOKEN` with the AutoNudge `xoxb-…` token (DMs will then come from AutoNudge; Ops Bot token kept elsewhere for rollback) |
| B — Named property | Update scheduler code to prefer `AUTONUDGE_SLACK_BOT_TOKEN`, then fall back to `SLACK_BOT_TOKEN` |

Keep a secure copy of the previous Ops Bot token until you verify live DMs.

**Never commit the token to git.**

## 3. Smoke-test DMs

With the AutoNudge token active (Approach A or temporary test):

- [ ] In the Year-Round Workshop workbook, run **Test Slack DM to me** (or equivalent Training Scheduler menu item)
- [ ] Open Slack → Direct Messages → **AutoNudge**
- [ ] Confirm the test message arrives from AutoNudge
- [ ] Check **TS Audit** for `SLACK` rows if the send fails (`missing_scope`, lookup errors, etc.)

## 4. Before live consultant messaging

- [ ] Confirm `TEST_MODE` behavior in TS Config (test DMs should still go to the configured test recipient while testing)
- [ ] Send one controlled resend/test offer if needed
- [ ] Only then set `TEST_MODE = FALSE` for live consultant DMs

## 5. Rollback

If AutoNudge DMs fail after cutover:

- [ ] Restore Ops Bot token into `SLACK_BOT_TOKEN`
- [ ] Re-run **Test Slack DM to me** and confirm Ops Bot delivery
- [ ] File follow-up with IT if scopes were incomplete (`users:read.email`, `im:write`, etc.)

## Related artifacts

- [`autonudge.manifest.yaml`](autonudge.manifest.yaml) — what IT installed
- [`IT_PERMISSION_REQUEST.md`](IT_PERMISSION_REQUEST.md) — original request text
