# Teacher List Specs

## Apps Script Project

- Script folder: `teacher_list`.
- Apps Script project ID: `1xJW2r9kBZHib0bKVPZuNo-rXUNofZKy5-F6T6Zy4OLrb5CIHu1QcOTw5`.
- The project is deployed as a web app endpoint.
- Web app execution: owner/deployer, `admindomini@iernestlluch.cat`.
- Web app access: anyone can access.

## File Structure

- `Code.js`: Apps Script server code, DB-resolution helpers, list data API, XLSX export, and permission grant helper.
- `Index.html`: static HTML structure.
- `Styles.html`: page CSS and Bootstrap Icons include.
- `Client.html`: browser-side JavaScript.
- `appsscript.json`: Apps Script manifest.
- `README.md`: script-specific documentation.
- `docs/SPECS.md`: script-specific specs.
- `.clasp.json`: local clasp binding, ignored by git.

## DB Resolution

The script works with DB using the same registry pattern as the other scripts in this repository.

1. Read script property `db`.
2. Treat that property as the ID of a registry spreadsheet.
3. Open sheet `tables` in the registry spreadsheet.
4. Find the row where column A is exactly `Dades de professors`.
5. Read column B from that row.
6. Treat that value as the DB spreadsheet ID.
7. Open DB sheet `Llista`.

## Endpoint Layout

The page has a reusable Apps Script HTML structure based on:

- `Index.html` for static structure;
- `Styles.html` included with `<?!= include_('Styles'); ?>`;
- `Client.html` included with `<?!= include_('Client'); ?>`.

Main layout:

- full-page loading overlay;
- sidebar navigation;
- main content area with view sections;
- bottom-right floating export button.

The sidebar has two sections:

- `Llistat`;
- `Estructura`.

## `Llistat` View

- Data source: DB spreadsheet `Dades de professors`, sheet `Llista`.
- Only rows where `ACTIU` is true are shown.
- `ACTIU` is DB column N.
- Boolean reads accept real boolean `true` and string `TRUE`.
- Filters:
  - text box filtering by teacher name and surnames;
  - department combo populated from active-row `DEPT.` values;
  - reset button.
- Displayed table columns:
  - `DEPT.` from column B;
  - `NOM` from column C;
  - `COGNOM1` from column D;
  - `COGNOM2` from column E;
  - `CORREU INSTIT` from column L.
- `CORREU INSTIT` is the actual DB header for column L.
- Rows are sorted by `COGNOM1`, then `COGNOM2`, then `NOM`, then `DEPT.`.

## `Estructura` View

- Section exists in the left menu and page shell.
- Behavior is reserved for a later phase.

## XLSX Export

The floating bottom-right export button:

- is visible on `Llistat`;
- uses Bootstrap Icons class `bi-file-earmark-spreadsheet`;
- exports the current filtered list to XLSX format.

Export process:

1. The browser calls backend function `createTeacherListXlsx(filters)` through `google.script.run`.
2. The backend reads DB and applies the current text and department filters.
3. The backend builds an in-memory XLSX model with rows, headers, title, subtitle, column widths, and formatting metadata.
4. The backend creates a temporary Google Spreadsheet in Drive.
5. The backend writes values and formatting into the temporary spreadsheet.
6. `SpreadsheetApp.flush()` forces pending spreadsheet writes.
7. The backend exports the temporary spreadsheet through the Drive export endpoint as XLSX.
8. The backend verifies that the export response status is HTTP 200.
9. The backend returns the XLSX file as base64 with `fileName` and MIME type.
10. The backend moves the temporary spreadsheet to Drive trash in a `finally` block.
11. The browser creates a temporary download link with the returned base64 content and downloads the file.

The temporary Google Spreadsheet is not kept in Drive after export.

## Permission Grant Function

Run `grantRequiredPermissions()` manually from Apps Script after changing scopes or installing the project.

The function touches:

- script properties;
- spreadsheet registry and DB reads;
- temporary Spreadsheet creation;
- Drive file access/trashing;
- Drive export through `UrlFetchApp`;
- `ScriptApp.getOAuthToken()`.

This is intended to trigger all required OAuth permission prompts before users access the endpoint.

## Apps Script Functions

Expected server-side functions:

- `doGet()`: returns the HTML page.
- `include_(filename)`: includes template partials.
- `grantRequiredPermissions()`: helper to trigger all needed authorization scopes.
- `getTeacherListData()`: returns active teachers and department options for the browser.
- `createTeacherListXlsx(filters)`: exports the current filtered list as XLSX.
- `getTeacherDbSpreadsheet_()`: resolves and opens DB.
- `getTeacherDbSheet_()`: opens DB sheet `Llista`.

## Manifest Scopes

Required scopes:

- `https://www.googleapis.com/auth/script.storage`
- `https://www.googleapis.com/auth/spreadsheets`
- `https://www.googleapis.com/auth/drive`
- `https://www.googleapis.com/auth/script.external_request`
