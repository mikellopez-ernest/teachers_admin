# Teacher List

Google Apps Script web app for read-only teacher list endpoints.

## Views

- `Llistat`: active-teacher list with name and department filters.
- `Per classe`: selected-class teacher/subject list from `distribucio`.
- `Estructura`: placeholder section for later behavior.

## Data Access

This script resolves DB the same way as the other scripts in this workspace:

1. Read script property `db`.
2. Open that registry spreadsheet.
3. Read sheet `tables`.
4. Find the row where column A is exactly `Dades de professors`.
5. Use column B from that row as the DB spreadsheet ID.
6. Open DB sheet `Llista`.

## List View

The `Llistat` view lists active teachers from `Dades de professors -> Llista`.

Displayed columns:

- `DEPT.`: column B.
- `NOM`: column C.
- `COGNOM1`: column D.
- `COGNOM2`: column E.
- `CORREU INSTIT`: column L.

Only rows where `ACTIU` column N is true are shown.

## Per Classe View

The `Per classe` view reads script property `db` as the registry spreadsheet ID, finds `Càrrega lectiva` in `tables`, and opens its `distribucio` sheet.

It builds a class combo from `Curs` and `Grup`, expands comma-separated groups and `TOTS`, and displays unique teacher/subject pairs for the selected class. Teacher order follows the order in `distribucio`. It enriches teachers with `CORREU INSTIT` by matching `distribucio` row 2 teacher names against `Càrrega lectiva -> professors` column Q and reading email from column L.

The view is strictly read-only.

## XLSX Export

The bottom-right floating export button uses Bootstrap Icon `bi-file-earmark-spreadsheet`. It exports the current filtered teacher list on `Llistat`, and the selected class teacher/subject/email table on `Per classe`.

The backend creates a temporary Google Spreadsheet, writes the list into it, exports it through the Drive XLSX export endpoint, returns the file as base64 to the browser, and trashes the temporary spreadsheet in a `finally` block.

## Permissions

Run `grantRequiredPermissions()` manually from the Apps Script editor to grant all required scopes before using the web app.

Required script property:

- `db`: registry spreadsheet ID.

## Local Development

This folder is connected to Apps Script with local clasp configuration.

The local `.clasp.json` file is intentionally ignored by git.

Detailed behavior is specified in [`docs/SPECS.md`](docs/SPECS.md).
