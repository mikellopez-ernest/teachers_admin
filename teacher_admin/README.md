# Professors DB Web App

Google Apps Script web app for reading professor data from a spreadsheet-based DB and managing status filters and teacher records.

## Data Model

The script property `Tables` must contain the ID of a registry spreadsheet. That registry spreadsheet must have a sheet named `tables` where:

- Column A contains logical table names.
- Column B contains spreadsheet IDs.

The app finds the row where column A is exactly `Dades de professors`; column B in that row is the DB spreadsheet ID. The registry lookup is exact and case-sensitive.

The DB spreadsheet must have a sheet named `Llista`. Row 1 is treated as the header row, and data starts at row 2.

The DB spreadsheet also has a `leave_absence` sheet with `row_id`, `teacher_code`, `substitute_code`, `start_date`, `end_date`, and `comments`. Both `teacher_code` and `substitute_code` store `REDUIT` from `Llista` column F.

## Displayed Columns

| UI column | DB source |
| --- | --- |
| `ESP` | column A |
| `DEPT.` | column B |
| `NOM SENCER` | columns C, D, E joined with spaces |
| `SITUACIO` | column G |
| `JORNADA` | column H |
| `DNI` | column I |
| `TELF` | column J |
| `CORREU` | column L |
| `XTEC` | column K |

## Status Fields

- `NOUS` reads DB column M as a boolean or string `TRUE`.
- `ACTIU` reads DB column N as active-state data; writes a real boolean.
- `BAIXA?` reads DB column O as a boolean or string `TRUE`; writes a real boolean.
- `SUBST?` reads DB column P as a boolean or string `TRUE`.

## Features

- Responsive Bootstrap web endpoint served by `doGet`.
- Bootstrap filter layout with a two-row status matrix.
- Bottom action bar stays fixed at the bottom of the page.
- Department filter populated from DB column B.
- Text filter over the full professor name.
- `RESET` clears only the department and name filters.
- Clickable table headers toggle A-Z / Z-A sorting.
- `NOM SENCER` displays name plus surnames, but sorts by first surname.
- Row checkboxes select rows for bulk updates.
- Header checkbox selects or deselects all currently filtered rows.
- Clicking a teacher-name hyperlink opens a Bootstrap-organized detail window.
- Detail windows can be changed to edit forms with the same block structure using `Edit`, then saved or cancelled.
- No delete option is provided.
- `Donar de baixa` / `Donar d'alta` manages one selected teacher at a time, updates absent-teacher column O, records dates/substitute/comments in `leave_absence`, marks the substitute with `SUBST?` and `ACTIU`, and clears those substitute flags when the leave ends.
- Starting or ending a leave calls the Horaris schedule cache rebuild endpoint with JSON POST `action=rebuildScheduleCache`.
- `Desactivar` / `Activar` updates column N.
- `Exportar` opens export options for selected rows: spreadsheet-compatible CSV, a browser-generated PDF signature list, or PDF teacher stickers.
- The signature list includes only selected teachers, sorted by first surname, with columns for teacher name, signature, and observations.
- `Etiquetes` creates a selected-teacher PDF with two columns and fifteen rows per page, with each teacher name centered in its cell.
- Rows with `SUBST?` checked are shown in slight green; leave-of-absence rows are shown in slight red.
- Action buttons warn if no row is selected.
- `BAIXA?` is read-only in the teacher edit form and is only changed through the leave workflow.

## Local Development

This project is connected to Apps Script with `clasp`.

The PDF exports are generated in the browser with jsPDF and jsPDF AutoTable loaded from CDN by `Index.html`.

Detailed behavior is specified in [`docs/SPECS.md`](docs/SPECS.md).

## Script Properties

- `Tables`: required registry spreadsheet ID.
- `cache_rebuild_token`: required token sent in the JSON POST payload to the Horaris cache rebuild endpoint after starting or ending a leave.
- `cache_rebuild_url`: optional override for the Horaris cache rebuild endpoint. If omitted, the app uses `https://script.google.com/macros/s/AKfycbyhSqCTkS27bDxsfILI64rlSMUTN5A7VbHGgpSf_G6efxrWfOuUKJULnN2rlMtHuWqwmA/exec`.

## Cache Rebuild Endpoint

The Horaris rebuild endpoint must be deployed so this app can reach it with `UrlFetchApp`.
In practice, that means the Horaris web app deployment must allow access by link/anyone, and authorization should be enforced by the shared `cache_rebuild_token`.

The request sent after leave changes is JSON POST:

```json
{
  "action": "rebuildScheduleCache",
  "token": "VALUE_FROM_cache_rebuild_token"
}
```

The expected response is JSON with `ok: true`. Non-JSON responses, non-2xx status codes, or `ok: false` are treated as errors.

## Cache Rebuild Logs

Leave changes write structured Apps Script logs with event names beginning `scheduleCacheRebuild.notify`.
Useful events are `scheduleCacheRebuild.notify.start`, `scheduleCacheRebuild.notify.response`, `scheduleCacheRebuild.notify.success`, and error variants for config, fetch, parse, or rebuild failures.
Logs include endpoint URL, token presence, HTTP status, affected teacher/substitute codes, and a truncated response body, but never the token value.
The same information is also written to DB sheet `cache_rebuild_log`, which is created automatically if missing. This is the easiest place to inspect failed cache notifications when Apps Script execution logs are empty.

Common failures:

- Missing `cache_rebuild_token`: create the script property in this project.
- HTTP 401 or Google Sign-In HTML response: redeploy the Horaris web app so `UrlFetchApp` can reach it without an interactive login.
- HTTP 200 with non-JSON response: the Horaris endpoint reached an HTML/default route instead of returning JSON for `rebuildScheduleCache`.

```sh
clasp push -f
clasp deploy --description "Deploy professors DB web app"
```

The script ID is stored in `.clasp.json`.
