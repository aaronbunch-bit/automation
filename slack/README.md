# AutoNudge (Slack)

IT-provisioned Slack App for **Year-Round Workshop Scheduling** automation messaging.

AutoNudge is a dedicated bot (separate from Ops Bot) that will send outbound Slack DMs to consultants—and related manager/coach nudges—driven by the Training Scheduler automations.

## Artifacts

| File | Purpose |
| --- | --- |
| [`autonudge.manifest.yaml`](autonudge.manifest.yaml) | Slack App Manifest for IT to create/install from |
| [`IT_PERMISSION_REQUEST.md`](IT_PERMISSION_REQUEST.md) | Paste-ready IT ticket/email |
| [`SETUP_CHECKLIST.md`](SETUP_CHECKLIST.md) | Post-install token storage and smoke-test steps |

## How to use

1. Submit [`IT_PERMISSION_REQUEST.md`](IT_PERMISSION_REQUEST.md) to IT with the manifest attached or linked.
2. When IT returns the `xoxb-…` Bot User OAuth Token, follow [`SETUP_CHECKLIST.md`](SETUP_CHECKLIST.md).
3. Point the Year-Round Workshop / Training Scheduler Slack sender at AutoNudge (token swap or named property) when ready to cut over from Ops Bot.

## Scopes (v1)

Outbound DMs only:

- `chat:write`
- `im:write`
- `users:read`
- `users:read.email`

No channel posting, slash commands, interactivity, or event subscriptions in v1.

## Intended automation consumers

Once the token is installed, these Training Scheduler flows are the expected consumers:

- Consultant training offers / reminders / reoffers / escalations
- Manager + sales coach nudges for unbooked offers

Apps Script wiring and Ops Bot cutover are a follow-up after IT delivers the token.
