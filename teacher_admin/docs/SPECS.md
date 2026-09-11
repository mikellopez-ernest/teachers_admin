# Professors Endpoint Specs

## Vocabulary

- **Tables registry**: spreadsheet whose ID is stored in the script property `Tables`.
- **DB**: spreadsheet resolved from the Tables registry row where column A is `Dades de professors`; column B of that row contains the DB spreadsheet ID.
- **Endpoint**: Apps Script web app endpoint that reads from DB and renders a responsive Bootstrap HTML page.

## Access Control

The endpoint has two access layers:

1. Apps Script web app access is restricted to the `iernestlluch.cat` domain.
2. Server-side authorization checks the signed-in user's institutional email against roles configured in script properties.

Required script property:

- `access_granted`: comma-separated list of allowed càrrecs. Example: `Coord. 3ESO,COCOBE`.

Authorization flow:

1. Read the active user's email with `Session.getActiveUser().getEmail()`.
2. Read script property `access_granted` and split it by commas.
3. Resolve the `Càrrega lectiva` spreadsheet through the `Tables` registry sheet `tables`.
4. Open `Càrrega lectiva -> carrecs`.
5. For each configured càrrec, find a matching row in column A (`carrec`).
6. Read the assigned person names from column D (`asignado?`). Multiple people are comma-separated.
7. Open `Càrrega lectiva -> professors`.
8. Match each assigned person name against column Q.
9. Read the matching institutional email from column L (`CORREU INSTIT`).
10. Allow access only when the active user's email matches one resolved email.

If access is denied, `doGet()` renders a simple no-access page. Each server-side data, edit, status, leave, and export function also calls the same authorization helper before returning data or writing DB.

Reusable helper responsibilities:

- `getAccessDecision_()`: resolves the active user and returns an authorization decision.
- `assertUserAccess_()`: throws when the current user is not authorized.
- `getAccessGrantedRoles_()`: reads and parses script property `access_granted`.
- `getPeopleByAccessRole_()`: maps configured càrrecs from `carrecs` column A to assigned people in column D.
- `getEmailsForPeople_()`: maps assigned people to institutional emails using `professors` columns Q and L.
- `createAccessDeniedOutput_()`: returns the no-access HTML page.

## DB Resolution

1. Read script property `Tables`.
2. Open that spreadsheet by ID.
3. Open sheet `tables`.
4. Find the row where column A exactly equals `Dades de professors`.
5. Read column B from that row.
6. Open that spreadsheet by ID. This spreadsheet is **DB**.

## Source Data

The endpoint reads professor rows from DB sheet `Llista`.

DB has a header row, so professor data starts at row 2.

DB also contains a `leave_absence` sheet used to record leave-of-absence periods. Its columns are:

| Column | Meaning |
| --- | --- |
| A | `row_id`: original row number in `Llista` |
| B | `teacher_code`: absent teacher code from `Llista` column F (`REDUIT`) |
| C | `substitute_code`: substitute teacher code from `Llista` column F (`REDUIT`) |
| D | `start_date` |
| E | `end_date` |
| F | `comments` |

Displayed table columns:

| UI column | DB source |
| --- | --- |
| `ESP` | column A |
| `DEPT.` | column B |
| `NOM SENCER` | concat of columns C, D, and E using a single space |
| `SITUACIO` | column G |
| `JORNADA` | column H |
| `DNI` | column I |
| `TELF` | column J |
| `CORREU` | column L |
| `XTEC` | column K |

Additional DB columns used by workflows:

| Meaning | DB source |
| --- | --- |
| `REDUIT` | column F |
| `SITUACIO` allowed values | `FUNC. DEF`, `FUNC. PERFIL`, `FUNC. SNS PLAÇA`, `INT`, `INT. PERF`, `LABORAL` |
| `JORNADA` allowed values | `SENCERA`, `MITJA`, `REDUCCIÓ UN TERÇ` |

Control/status columns:

| Meaning | DB source |
| --- | --- |
| `NOUS` | column M |
| `ACTIU` | column N |
| `BAIXA?` | column O |
| `SUBST?` | column P |

## Page Layout

The HTML page contains:

1. Filters above the table.
2. Bootstrap status matrix above the table.
3. Data table with selectable rows.
4. Action buttons fixed at the bottom of the page.

Rows with `SUBST?` true are shown with a slight green background. Rows with `BAIXA?` true are shown with a slight red background. If both conditions apply, the leave-of-absence red background takes precedence.

## Filters

### Department Filter

- A combo/select labeled `DEPT.`.
- Values are read from DB column B when the page loads.
- Values should be unique.
- Selecting a department filters visible rows to that department.

### Name Filter

- A text input.
- It filters by columns C, D, and E from DB.
- In the UI those values are shown as `NOM SENCER`.
- Filtering should match against the concatenated full name.

### Reset

- A button with text `RESET`.
- It clears the department filter and the name filter.
- It does not reset the status matrix.

## Status Matrix

The status filters are displayed as two rows: row 1 has `Actius`, `No actius`, and `Només nous`; row 2 has `No baixa` and `Baixa`. They are additive except `Només nous`, which is exclusive.

| Row | Value 1 | Default | Value 2 | Default | Value 3 | Default |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | `Actius` | checked | `No actius` | unchecked | `Només nous` | unchecked |
| 2 | `No baixa` | checked | `Baixa` | unchecked | | |

- `Actius` checked includes rows where DB column N is true.
- `No actius` checked includes rows where DB column N is not true.
- `No baixa` checked includes rows where DB column O is not true.
- `Baixa` checked includes rows where DB column O is true.
- `Només nous` checked restricts the result to rows where DB column M is true.
- If none of the active-state boxes are checked, no active-state rows are shown.
- If none of the baixa-state boxes are checked, no baixa-state rows are shown.
- True can be a real boolean `true` or the string `TRUE`, matched case-insensitively.

## Sorting

- Every displayed table header is clickable.
- Clicking a header sorts the table by that displayed column.
- Clicking `NOM SENCER` sorts by first surname, then second surname, then name.
- Row data sent to the browser includes raw `NOM`, `COGNOM1`, and `COGNOM2` fields so exports can compose names differently from the visible `NOM SENCER`.
- First click sorts A-Z.
- Clicking the same header again toggles to Z-A.
- Sorting should apply to the currently filtered rows.

## Row Selection

- Every row has a checkbox at the far left.
- Selected rows are the targets for action buttons.
- Row identity must preserve the original DB row number so updates write back to the correct DB rows after filtering or sorting.
- The table header contains a checkbox that selects or deselects all currently filtered rows.

## Teacher Detail And Edit

- Clicking the `NOM SENCER` hyperlink opens a small modal window with teacher information organized with Bootstrap.
- Detail block 1 has two columns:
  - Column 1: `NOM`, `COGNOM1`, `COGNOM2`.
  - Column 2: `ESP`, `REDUIT`, `DEP`, `SITUACIO`, `JORNADA`.
- Detail block 2 has two columns:
  - Column 1: `DNI`, `TELF`, `CORREU`, `XTEC`.
  - Column 2: `ACTIU`, `BAIXA`, `NOUS`, `SUBST?`, presented as boolean values.
- The detail window has an `Edit` button.
- Clicking `Edit` converts the detail window into a form with the same block and column structure.
- In edit mode, `SITUACIO` and `JORNADA` are edited with fixed option controls.
- In edit mode, `ACTIU`, `NOUS`, and `SUBST?` are editable checkboxes.
- In edit mode, `BAIXA` is displayed as a disabled/read-only checkbox. It can only be changed by the `Donar de baixa` / `Donar d'alta` workflow.
- The edit form has `Save` and `Cancel` buttons.
- `Save` writes the edited values back to the same DB row.
- `Cancel` returns to the read-only detail view without writing changes.
- There is no delete option.

## Action Buttons

Buttons are fixed at the bottom of the page and operate on selected rows.

If no rows are selected, clicking an action button shows a warning and does not update DB.

### `Donar de baixa` / `Donar d'alta`

- The button is disabled when no teacher is selected.
- The workflow only accepts one selected teacher. If more than one teacher is selected, show a warning and do not update DB.
- If the selected teacher is not currently `BAIXA`, the button text is `Donar de baixa`.
- If the selected teacher is currently `BAIXA`, the button text is `Donar d'alta`.

#### Start Leave

- Clicking `Donar de baixa` opens a modal asking for:
  - start date, defaulting to today, using a Spanish datepicker with Monday as the first day of the week;
  - substitute teacher, selected from teachers in `Llista` with a `REDUIT` code, where `BAIXA?` column O is not true and `SUBST?` column P is not true; the selected absent teacher is excluded, but teachers with the same `ESP` are allowed; the combo displays name and surnames, but the stored substitute code is `REDUIT` from column F;
  - comments.
- Confirming the modal sets `Llista` column O to boolean `true` for the selected absent teacher.
- Confirming the modal sets `Llista` column P `SUBST?` to boolean `true` for the selected substitute teacher.
- Confirming the modal sets `Llista` column N `ACTIU` to boolean `true` for the selected substitute teacher.
- Confirming the modal appends a row to `leave_absence` with `row_id`, `teacher_code` from the absent teacher's `REDUIT` column F, `substitute_code` from the substitute teacher's `REDUIT` column F, `start_date`, blank `end_date`, and `comments`.
- After DB writes are flushed, the app calls the schedule cache rebuild endpoint with POST action `rebuildScheduleCache`.
- After updating, refresh the visible data.

#### End Leave

- Clicking `Donar d'alta` opens a confirmation modal asking for end date, defaulting to today, using a Spanish datepicker with Monday as the first day of the week.
- Confirming the modal fills `end_date` in the latest open `leave_absence` row for that teacher.
- Confirming the modal sets `Llista` column O `BAIXA?` to boolean `false` for the selected absent teacher. The open `leave_absence` record is the authority for ending a leave, so the action can recover if column O is not currently truthy.
- Confirming the modal finds the recorded substitute by `leave_absence.substitute_code` and sets that substitute's `SUBST?` column P and `ACTIU` column N to boolean `false`.
- After DB writes are flushed, the app calls the schedule cache rebuild endpoint with POST action `rebuildScheduleCache`.
- After updating, refresh the visible data.

#### Schedule Cache Rebuild Notification

- The leave workflow must notify the Horaris cache rebuild web app after starting or ending a leave.
- The default endpoint URL is `https://script.google.com/macros/s/AKfycbyhSqCTkS27bDxsfILI64rlSMUTN5A7VbHGgpSf_G6efxrWfOuUKJULnN2rlMtHuWqwmA/exec`.
- The endpoint URL can be overridden with script property `cache_rebuild_url`.
- The authorization token is read from script property `cache_rebuild_token`.
- If `cache_rebuild_token` is not set in this project, the leave workflow reports an error.
- The Apps Script manifest must include `https://www.googleapis.com/auth/script.external_request` so `UrlFetchApp` can call the Horaris endpoint.
- The Horaris web app deployment must be reachable by `UrlFetchApp`; it should enforce authorization with `cache_rebuild_token` rather than requiring an interactive Google Sign-In page.
- The app calls the endpoint with JSON POST:

```js
{
  action: 'rebuildScheduleCache',
  token: 'VALUE_FROM_cache_rebuild_token'
}
```

- A response is accepted only when the HTTP status is 2xx and the parsed JSON response has `ok: true`.
- Rebuild errors are surfaced to the user so stale schedule-cache problems are visible.
- The app writes structured Apps Script logs for the notification flow:
  - `scheduleCacheRebuild.notify.start`
  - `scheduleCacheRebuild.notify.response`
  - `scheduleCacheRebuild.notify.success`
  - `scheduleCacheRebuild.notify.configError`
  - `scheduleCacheRebuild.notify.fetchError`
  - `scheduleCacheRebuild.notify.parseError`
  - `scheduleCacheRebuild.notify.rebuildError`
- Logs include the leave event, affected row numbers/codes, endpoint URL, whether a token is configured, HTTP status, and a truncated response body. Logs never include the token value.
- The app also writes a persistent audit row in DB sheet `cache_rebuild_log` for each notification step.
- If `cache_rebuild_log` does not exist, the app creates it.
- `cache_rebuild_log` columns are:
  - `timestamp`
  - `event`
  - `leave_event`
  - `row_number`
  - `teacher_code`
  - `substitute_code`
  - `leave_absence_row`
  - `substitute_row_number`
  - `endpoint_url`
  - `has_token`
  - `status_code`
  - `ok`
  - `response_text`
  - `error`
- The persistent log never stores the token value.

### `Desactivar` / `Activar`

- Button text depends on the active-state matrix:
  - If only `No actius` is checked: `Activar`.
  - Otherwise: `Desactivar`.
- On click, update column N in DB for every selected row:
  - `Desactivar` sets column N to boolean `false`.
  - `Activar` sets column N to boolean `true`.
- After updating, refresh the visible data.

### `Exportar`

- Button appears next to `Desactivar` / `Activar`.
- If no rows are selected, clicking it shows a warning and does not continue.
- If rows are selected, clicking it opens a small modal window with export options.
- Export option `Full de càlcul` downloads a spreadsheet-compatible CSV file containing selected rows.
- Export option `Llistat signatures` downloads a PDF containing selected rows.
- Export option `Etiquetes` downloads a PDF sticker sheet containing selected rows.
- CSV exported data includes DB columns A through P and the header row.
- The app does not create any export file in Google Drive.

#### `Llistat signatures` PDF

- The PDF contains only selected teachers.
- Teacher names are sorted alphabetically by first surname, using the same sort key as `NOM SENCER`.
- The PDF contains a table with three columns:
  - column 1: teacher full name from DB columns C, D, and E;
  - column 2: blank signing space;
  - column 3: `Observacions`.
- Column 1 is sized from the widest selected teacher name, with a cap so column 3 has more usable width.
- Rows are tall enough to fit 17 teachers per page.
- Cell text is vertically centered and horizontally left-aligned.
- The PDF is generated in the browser and downloaded to the user's computer.

#### `Etiquetes` PDF

- The PDF contains only selected teachers.
- Teacher names are sorted alphabetically by first surname, using the same sort key as `NOM SENCER`.
- The PDF contains a table with 2 columns and 15 rows per page.
- The PDF has no header row and no title text.
- Each cell contains one teacher name from DB columns D, E, and C, composed as `COGNOM1 COGNOM2, NOM`.
- Teacher names are rendered in uppercase.
- Teacher names are horizontally and vertically centered in each cell.
- Blank cells are allowed on the final page when the selected-teacher count is odd.
- The PDF is generated in the browser and downloaded to the user's computer.

## Apps Script Functions

Expected server-side functions:

- `doGet(e)`: returns the HTML page.
- `getDbSpreadsheet_()`: resolves and opens DB.
- `getProfessorsData()`: reads DB rows and returns data for the page.
- `startLeaveAbsence(rowNumber, leaveData)`: starts a leave workflow, writes absent-teacher DB column O, writes substitute columns N and P, and appends a `leave_absence` row.
- `endLeaveAbsence(rowNumber, leaveData)`: ends a leave workflow, writes absent-teacher DB column O, clears substitute columns N and P, and fills the open `leave_absence.end_date`.
- `notifyScheduleCacheRebuild_()`: POSTs `action=rebuildScheduleCache` to the Horaris cache rebuild endpoint using script property `cache_rebuild_token`.
- `getLeaveAbsenceSheet_()`: opens DB sheet `leave_absence` and ensures expected headers.
- `updateActiu(rowNumbers, value)`: writes boolean values to DB column N.
- `getTeacherDetails(rowNumber)`: reads all editable teacher fields from DB columns A through P.
- `saveTeacherDetails(rowNumber, fields)`: writes edited teacher fields back to DB columns A through P.
- `exportTeachers(rowNumbers)`: returns CSV data for selected teacher rows so the browser can download it.

Expected client-side export functions:

- `downloadSignatureListPdf()`: generates the selected-teacher signature PDF in the browser.
- `downloadLabelsPdf()`: generates the selected-teacher label PDF in the browser.

## Decisions

- Source sheet inside DB: `Llista`.
- DB has a header row; data starts at row 2.
- Status reads accept real booleans and string `TRUE`; status writes use real booleans.
- DB column M is interpreted as `NOUS`.
- DB column N is interpreted as active-state data; `No actius` shows the inverse of column N.
- DB column O is interpreted as `BAIXA?`.
- DB column P is interpreted as `SUBST?`.
- `RESET` clears only the department and name filters.
- Action buttons warn when no row is selected.
- Phase two edit scope is DB columns A through P, except `BAIXA?` which is read-only in the edit form.
- Export formats are selected from the `Exportar` modal: `Full de càlcul` CSV, `Llistat signatures` PDF, and `Etiquetes` PDF.

## Repository Security

- `.clasp.json` must not be committed to public GitHub repositories.
- `.clasp.json` contains the Apps Script project ID. It is not an OAuth secret, but exposing it unnecessarily reveals project metadata.
- Local clasp credentials, such as `.clasprc.json` or `~/.clasprc.json`, must never be committed.
- The repository `.gitignore` excludes `.clasp.json`, `.clasprc.json`, and `.DS_Store`.
- If `.clasp.json` has already been committed, remove it from Git tracking with `git rm --cached .clasp.json` before the next commit.
