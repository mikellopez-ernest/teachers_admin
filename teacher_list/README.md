# Teacher List

Google Apps Script web app for read-only teacher list endpoints.

## Access Control

The web app is deployed for `iernestlluch.cat` domain users only, executes as the accessing user, and then applies a second server-side authorization check.

Required script property:

- `access_granted`: comma-separated list of allowed càrrecs, for example `Coord. 3ESO,COCOBE`.

How authorization is resolved:

1. The app reads the signed-in user's email.
2. It reads `access_granted`.
3. It opens `Càrrega lectiva -> carrecs` and matches each configured càrrec against column A.
4. It reads assigned people from column D (`asignado?`), split by commas.
5. It opens `Càrrega lectiva -> professors`, matches each person against column Q, and reads `CORREU INSTIT` from column L.
6. Access is granted only if the signed-in user's email is one of those resolved emails.

The reusable authorization helper block is isolated in `Code.js` with `getAccessDecision_()`, `assertUserAccess_()`, `getAccessGrantedRoles_()`, `getPeopleByAccessRole_()`, `getEmailsForPeople_()`, and `createAccessDeniedOutput_()`.

## Views

- `Llistat`: active-teacher list with name and department filters.
- `Per classe`: selected-class teacher/subject list from `distribucio`.
- `Estructura`: read-only list of càrrecs from `Càrrega lectiva -> carrecs`.

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

It builds a two-level selectable class combo from `Curs` and `Grup`, with level options such as `1ESO` and indented group options such as `--- 1ESO A`. Selecting a level shows all teacher/subject pairs involved in that full level; selecting an indented group shows only that class. It expands comma-separated groups and `TOTS`, keeps teacher order from `distribucio`, and enriches teachers with `CORREU INSTIT` by matching `distribucio` row 2 teacher names against `Càrrega lectiva -> professors` column Q and reading email from column L.

The view is strictly read-only.

## Estructura View

The `Estructura` view reads `Càrrega lectiva -> carrecs` and shows rows where `is_carrec` column F is true.

Displayed columns:

- `carrec`: column A.
- `asignado?`: column D.

A text filter searches both displayed columns, and `Reinicia` clears it.

## XLSX Export

The bottom-right floating export button uses Bootstrap Icon `bi-file-earmark-spreadsheet`. It exports the current filtered teacher list on `Llistat`, the selected class teacher/subject/email table on `Per classe`, and the current filtered càrrecs table on `Estructura`.

The backend creates a temporary Google Spreadsheet, writes the list into it, exports it through the Drive XLSX export endpoint, returns the file as base64 to the browser, and trashes the temporary spreadsheet in a `finally` block.

## Permissions

Run `grantRequiredPermissions()` manually from the Apps Script editor to grant all required scopes before using the web app.

Required script property:

- `db`: registry spreadsheet ID.
- `access_granted`: comma-separated list of allowed càrrecs.

## Local Development

This folder is connected to Apps Script with local clasp configuration.

The local `.clasp.json` file is intentionally ignored by git.

Detailed behavior is specified in [`docs/SPECS.md`](docs/SPECS.md).
