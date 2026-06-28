# Professors Endpoint Specs

## Vocabulary

- **Tables registry**: spreadsheet whose ID is stored in the script property `Tables`.
- **DB**: spreadsheet resolved from the Tables registry row where column A is `Dades de professors`; column B of that row contains the DB spreadsheet ID.
- **Endpoint**: Apps Script web app endpoint that reads from DB and renders an HTML page.

## DB Resolution

1. Read script property `Tables`.
2. Open that spreadsheet by ID.
3. Open sheet `tables`.
4. Find the row where column A, trimmed, equals `Dades de professors`.
5. Read column B from that row.
6. Open that spreadsheet by ID. This spreadsheet is **DB**.

## Source Data

The endpoint reads professor rows from DB sheet `Llista`.

DB has a header row, so professor data starts at row 2.

Displayed table columns:

| UI column | DB source |
| --- | --- |
| `ESP` | column A |
| `DEPT.` | column B |
| `NOM SENCER` | concat of columns C, D, and E using a single space |
| `SITUACIO` | column G |
| `DNI` | column H |
| `TELF` | column I |
| `CORREU` | column K |
| `XTEC` | column J |

Control/status columns:

| Meaning | DB source |
| --- | --- |
| `BAIXA?` | column L |
| `NO ACTIUS` | column N |

## Page Layout

The HTML page contains:

1. Filters above the table.
2. Status checkboxes above the table.
3. Data table with selectable rows.
4. Action buttons below the table.

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
- It does not reset `BAIXA?` or `NO ACTIUS`.

## Status Checkboxes

### `BAIXA?`

- Default: deactivated.
- When deactivated, show rows where DB column L is not true.
- When activated, show only rows where DB column L is true.
- True can be a real boolean `true` or the string `TRUE`, matched case-insensitively.

### `NO ACTIUS`

- Default: deactivated.
- DB column N is treated as active-state data.
- When deactivated, show rows where DB column N is true.
- When activated, show only rows where DB column N is not true.
- True can be a real boolean `true` or the string `TRUE`, matched case-insensitively.

## Sorting

- Every displayed table header is clickable.
- Clicking a header sorts the table by that displayed column.
- Clicking `NOM SENCER` sorts by first surname, then second surname, then name.
- First click sorts A-Z.
- Clicking the same header again toggles to Z-A.
- Sorting should apply to the currently filtered rows.

## Row Selection

- Every row has a checkbox at the far left.
- Selected rows are the targets for action buttons.
- Row identity must preserve the original DB row number so updates write back to the correct DB rows after filtering or sorting.
- The table header contains a checkbox that selects or deselects all currently filtered rows.

## Teacher Detail And Edit

- Clicking `NOM SENCER` opens a small modal window with all teacher information from DB columns A through N.
- The detail window has an `Edit` button.
- Clicking `Edit` converts the detail window into a form.
- The edit form has `Save` and `Cancel` buttons.
- `Save` writes the edited values back to the same DB row.
- `Cancel` returns to the read-only detail view without writing changes.
- There is no delete option.

## Action Buttons

Buttons appear below the table and operate on selected rows.

If no rows are selected, clicking an action button shows a warning and does not update DB.

### `Donar de baixa` / `Donar de alta`

- Button text depends on the current `BAIXA?` checkbox:
  - If `BAIXA?` is activated: `Donar de alta`.
  - If `BAIXA?` is deactivated: `Donar de baixa`.
- On click, update column L in DB for every selected row:
  - `Donar de baixa` sets column L to boolean `true`.
  - `Donar de alta` sets column L to boolean `false`.
- After updating, refresh the visible data.

### `Desactivar` / `Activar`

- Button text depends on the current `NO ACTIUS` checkbox:
  - If `NO ACTIUS` is activated: `Activar`.
  - If `NO ACTIUS` is deactivated: `Desactivar`.
- On click, update column N in DB for every selected row:
  - `Desactivar` sets column N to boolean `false`.
  - `Activar` sets column N to boolean `true`.
- After updating, refresh the visible data.

### `Exportar`

- Button appears next to `Desactivar` / `Activar`.
- If no rows are selected, clicking it shows a warning and does not continue.
- If rows are selected, clicking it opens a small modal window with export options.
- Phase two includes one export option: `Full de càlcul`, which downloads a spreadsheet-compatible CSV file.
- Confirming the export downloads a CSV file containing the selected rows to the user's computer.
- Exported data includes DB columns A through N and the header row.
- The app does not create any export file in Google Drive.

## Apps Script Functions

Expected server-side functions:

- `doGet(e)`: returns the HTML page.
- `getDbSpreadsheet_()`: resolves and opens DB.
- `getProfessorsData()`: reads DB rows and returns data for the page.
- `updateBaixa(rowNumbers, value)`: writes boolean values to DB column L.
- `updateActiu(rowNumbers, value)`: writes boolean values to DB column N.
- `getTeacherDetails(rowNumber)`: reads all editable teacher fields from DB columns A through N.
- `saveTeacherDetails(rowNumber, fields)`: writes edited teacher fields back to DB columns A through N.
- `exportTeachers(rowNumbers)`: returns CSV data for selected teacher rows so the browser can download it.

## Decisions

- Source sheet inside DB: `Llista`.
- DB has a header row; data starts at row 2.
- Status reads accept real booleans and string `TRUE`; status writes use real booleans.
- DB column N is interpreted as active-state data, so `NO ACTIUS` is the inverse of column N.
- `RESET` clears only the department and name filters.
- Action buttons warn when no row is selected.
- Phase two edit scope is DB columns A through N.
- Phase two export format is CSV; the visible option is `Full de càlcul`.
