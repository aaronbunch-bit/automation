# ReflexAI weekly simulation exception reporting

This repository contains a Google Apps Script automation that pulls ReflexAI
simulation results into a Google Sheet and sends weekly exception-only emails to
Managers and Team Leads.

The report highlights:

1. Simulations that have not been attempted.
2. Simulations completed below the required passing threshold, defaulting to 80%.

Each exception row identifies the representative, journey, simulation,
completion status, score, and required follow-up action.

## What is included

- `Code.gs` - Apps Script source code.
- `appsscript.json` - Apps Script manifest and OAuth scopes.

## Required ReflexAI detail

ReflexAI does not publish public API documentation for simulation journey score
exports. Because of that, this project does **not** hard-code guessed ReflexAI
URLs or login behavior.

Use the credentials/API information you have from ReflexAI to set Script
Properties. If your credentials are only for browser login and do not include an
API or report-export endpoint, ask ReflexAI for these exact details:

```text
We need to automate weekly Google Sheets reporting for Prepare/Studio monthly
workshop journeys. Please provide the supported API or report-export endpoint
that returns one row per representative/journey/simulation attempt, including:

- representative name and email
- manager name and email
- team lead name and email, if available
- journey ID and journey name
- simulation ID and simulation name
- completion/attempt status
- score percentage
- completed-at or attempted-at timestamp, if available

Please also provide the authentication method for server-to-server access
(bearer token, API key header, or username/password token exchange), the base
tenant URL, and the JSON path or CSV format returned by the report.
```

## Google Sheet setup

1. Create or choose the Google Sheet that should receive the report.
2. Copy the Sheet ID from the URL:

   `https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit`

3. Create a new Apps Script project attached to that Sheet.
4. Add the files from this repository to the Apps Script project:
   - paste `Code.gs` into the script editor
   - paste `appsscript.json` into the manifest file
5. In Apps Script, open **Project Settings > Script Properties** and add the
   settings below.

## Script Properties

### Required

| Property | Description |
| --- | --- |
| `SPREADSHEET_ID` | Google Sheet ID that receives the exception report. |
| `REFLEXAI_BASE_URL` | Base ReflexAI tenant URL. |
| `REFLEXAI_REPORT_PATH` | ReflexAI API/report-export path that returns simulation results. |

### Authentication: choose one

| Auth type | Properties |
| --- | --- |
| Bearer token | `REFLEXAI_BEARER_TOKEN` |
| API key | `REFLEXAI_API_KEY`, optional `REFLEXAI_API_KEY_HEADER` (defaults to `x-api-key`) |
| Login token exchange | `REFLEXAI_USERNAME`, `REFLEXAI_PASSWORD`, `REFLEXAI_LOGIN_PATH`, optional `REFLEXAI_LOGIN_TOKEN_JSON_PATH` |

### Optional

| Property | Default | Description |
| --- | --- | --- |
| `REFLEXAI_REPORT_ROWS_JSON_PATH` | none | Dot path to the array of report rows if the JSON response is nested, for example `data.rows`. Not needed for CSV. |
| `FIELD_MAP_JSON` | built-in common names | JSON object mapping the script's normalized fields to ReflexAI field names. |
| `MONTHLY_JOURNEY_IDS` | all journeys | Comma-separated list of journey IDs to include. |
| `MONTHLY_JOURNEY_NAME_PATTERN` | all journey names | Case-insensitive regular expression for monthly workshop journeys, for example `monthly workshop`. |
| `PASSING_SCORE_PERCENT` | `80` | Passing score threshold. |
| `SEND_EMAILS` | `false` | Set to `true` after validation to send email notifications. |
| `DRY_RUN` | `false` | Set to `true` to log email output without sending. |
| `EMAIL_SUBJECT_PREFIX` | `ReflexAI weekly simulation follow-up` | Email subject prefix. |
| `EMAIL_SENDER_NAME` | `ReflexAI Simulation Reporting` | Display name for outgoing emails. |

## Expected ReflexAI report fields

The script accepts JSON or CSV. By default it looks for common field names such
as:

- `repName`, `representativeName`, `learnerName`, `agentName`, `userName`
- `repEmail`, `representativeEmail`, `learnerEmail`, `agentEmail`, `userEmail`
- `managerName`, `managerEmail`
- `teamLeadName`, `teamLeadEmail`
- `journeyId`, `journeyName`
- `simulationId`, `simulationName`
- `status`, `completionStatus`, `attemptStatus`
- `score`, `scorePercent`, `overallScore`

If ReflexAI returns different names, set `FIELD_MAP_JSON`. Example:

```json
{
  "repName": "employee.full_name",
  "repEmail": "employee.email",
  "managerName": "manager.full_name",
  "managerEmail": "manager.email",
  "teamLeadName": "teamLead.full_name",
  "teamLeadEmail": "teamLead.email",
  "journeyId": "journey.id",
  "journeyName": "journey.title",
  "simulationId": "simulation.id",
  "simulationName": "simulation.title",
  "status": "attempt.status",
  "score": "attempt.score_percent",
  "completedAt": "attempt.completed_at"
}
```

## Running the automation

1. In Apps Script, run `createSetupSheets`.
2. Run `validateConfiguration`.
3. Run `runWeeklySimulationExceptionReport` with `SEND_EMAILS=false` first.
4. Review the `Simulation Exceptions` tab.
5. Set `SEND_EMAILS=true`.
6. Run `installWeeklyTrigger`.

The weekly trigger runs Monday at 8 AM in the script timezone configured in
`appsscript.json`.

## Output tabs

- `Simulation Exceptions` - replaced on every run with the current exception
  list.
- `Run Log` - appends a run summary.
- `Setup Checklist` - generated setup reference inside the Sheet.

## Security notes

- Do not commit ReflexAI credentials to this repository.
- Store credentials only in Apps Script Properties.
- Prefer a ReflexAI API key or bearer token scoped to reporting/read-only access.
- Keep `SEND_EMAILS=false` until the first successful validation run is reviewed.
