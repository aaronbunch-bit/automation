# automation

Internal automation for Varsity Tutors workflows.

## Slack — AutoNudge

Create the dedicated AutoNudge Slack App via Playwright (workspace create auto-requests IT/admin approval):

```bash
export VT_SSO_PASSWORD='...'   # Google Workspace password for aaron.bunch@varsitytutors.com
npm install && npm run create-autonudge
```

Details: [`slack/README.md`](slack/README.md)
