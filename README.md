# ReflexAI weekly simulation exception reporting

This is the current band-aid workflow while Bobby gets the ReflexAI API key from
Derek.

Instead of automatically pulling from ReflexAI, you export the ReflexAI CSV once
per week, import it into a Google Sheet, and run the script. The script then
creates the manager follow-up list.

## What the report catches

The report only shows problems:

1. A simulation is `Not Started` or `In Progress`.
2. A simulation is `Completed` but the score is below 80%.

The report does not list people who are already passing.

## Tabs in the Google Sheet

The script creates these tabs:

| Tab | What it is for |
| --- | --- |
| `ReflexAI CSV Dump` | Import the weekly ReflexAI CSV here. Replace the old data each week. |
| `Manager Roster` | Paste representative-to-manager/team lead emails here. Keep this updated as people move teams. |
| `Simulation Exceptions` | The script writes the weekly follow-up list here. |
| `Run Log` | History of when the report ran. |
| `Setup Checklist` | Built-in reminder sheet. |

## Weekly process

1. Export the CSV from ReflexAI.
2. Open the Google Sheet.
3. Go to the `ReflexAI CSV Dump` tab.
4. Replace the old CSV data with the new CSV data.
5. Run `ReflexAI Reporting > Run report from CSV dump now`.
6. Review `Simulation Exceptions`.
7. Run `ReflexAI Reporting > Test email batches to Aaron/Robert` to send test
   copies to Aaron and Robert first.
8. Run `ReflexAI Reporting > Email managers current exceptions` when you are
   ready to email managers.

The test email step sends the same manager batches to
`aaron.bunch@varsitytutors.com` and `robert.sorrell@varsitytutors.com`, then
marks those rows with `Test Sent = Y`.

The real email step sends one batched email per manager, not one email per row.
Inside each email, exceptions are grouped to one entry per representative and
journey. Completed simulations below 80% appear first with a red-to-yellow score
gradient, followed by not-started and in-progress simulations.
After each manager email sends successfully, the script marks the related
exception rows with `Manager Email Sent = Yes` and a `Manager Email Sent At`
timestamp. Rows already marked `Yes` are skipped to prevent duplicate manager
emails.

## Manager Roster format

The ReflexAI CSV you provided does not include manager or team lead emails, so
the Sheet needs one simple roster tab.

Use these columns:

```text
Representative Email | Manager Name | Manager Email | Team Lead Name | Team Lead Email
```

Example:

```text
jane.rep@example.com | Sam Manager | sam.manager@example.com | Taylor Lead | taylor.lead@example.com
```

This roster does not need to be replaced every week unless assignments change.

## ReflexAI CSV columns currently supported

The uploaded CSV has these headers, and the script now supports them directly:

```text
User Name
User Email
Simulation Name
Status
Best Score (%)
Passing Score (%)
Passed
Attempts Used
Max Attempts
Completed At
Journey Completion
```

The script uses the mission requirement of 80% as the passing threshold. It does
not rely on the `Passing Score (%)` column because the sample CSV had `2` in
that column, which does not match the stated 80% business rule.

## Simple setup

1. Create a Google Sheet.
2. Open **Extensions > Apps Script**.
3. Paste `Code.gs` into the script editor.
4. Add `appsscript.json` as the manifest.
5. Run `createSetupSheets` once.
6. Import the ReflexAI CSV into `ReflexAI CSV Dump`.
7. Add manager/team lead emails in `Manager Roster`.
8. Run `runWeeklySimulationExceptionReport`.

Start with `SEND_EMAILS=false` so you can review the output first. Set
`SEND_EMAILS=true` only after the report looks right.

## Later API upgrade

When Bobby/Derek provide the ReflexAI API key, set `DATA_SOURCE=api` and add the
ReflexAI API settings. The same reporting and email logic will keep working; the
only thing that changes is where the raw data comes from.
