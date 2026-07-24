# IT Permission Request — Slack App: AutoNudge

**Copy/paste the section below into an IT ticket or email.** Attach or link [`autonudge.manifest.yaml`](autonudge.manifest.yaml).

---

## Subject

Request: Create and install Slack App **AutoNudge** (Year-Round Workshop Scheduling messaging)

## Requester

- **Name:** Aaron Bunch
- **Email:** aaron.bunch@varsitytutors.com
- **Team / use case:** Year-Round Workshop Scheduling automation (consultant training offers and related nudges)

## Business purpose

We need a dedicated Slack bot, **AutoNudge**, to send automated direct messages about Year-Round Workshop Scheduling updates. Recipients are consultants (and, where configured, their managers and sales coaches). Message types include training slot offers, reminders, reoffers when capacity changes, escalations, and manager/coach nudges for unbooked offers.

This app is **separate from Ops Bot**. Please do **not** reuse or replace Ops Bot credentials. AutoNudge should have its own app identity and Bot User OAuth Token.

## Requested actions

1. Create a new Slack App named **AutoNudge** in the Varsity Tutors workspace using the attached App Manifest (`autonudge.manifest.yaml` — Create New App → From a manifest).
2. Install the app to the Varsity Tutors workspace.
3. Return the **Bot User OAuth Token** (`xoxb-…`) to Aaron Bunch via an approved secret channel (do not put the token in email/chat history if policy requires a vault or 1:1 secret share).
4. Confirm the bot display name in Slack is **AutoNudge**.

## Bot token scopes (least privilege)

| Scope | Justification |
| --- | --- |
| `chat:write` | Send DM message content to consultants / managers / coaches |
| `im:write` | Open 1:1 DM conversations with users |
| `users:read` | Resolve Slack users by name / alias (`firstname.lastname`) |
| `users:read.email` | Resolve consultants by work email address |

**Not requested:** channel post/history, file upload, admin, commands, interactivity, or event subscriptions. v1 is **outbound DM messaging only**.

## Security / handling

- Token will be stored only in Google Apps Script **Script Properties** / TS Config for the Year-Round Workshop workbook — **not** committed to git.
- No user (xoxp) tokens. Bot token only.
- App does not require public distribution or App Directory listing; workspace install is sufficient.
- Ops Bot remains unchanged; cutover of scheduler messaging to AutoNudge is a separate, intentional step after this token is received.

## Manifest location

Repository path: `slack/autonudge.manifest.yaml` (this package).

## Acceptance criteria

- [ ] AutoNudge app exists and is installed on the Varsity Tutors Slack workspace
- [ ] Bot appears as **AutoNudge** in DMs
- [ ] Bot User OAuth Token (`xoxb-…`) delivered securely to Aaron Bunch
- [ ] Granted scopes match the table above (no extra admin scopes)
