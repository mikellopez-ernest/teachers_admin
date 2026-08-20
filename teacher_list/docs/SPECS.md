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

The sidebar has three sections:

- `Llistat`;
- `Per classe`;
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

## `Per classe` View

Goal: show which teachers teach a selected class/group and which subject each teacher teaches.

Data source for this section:

- Read script property `db`.
- Treat `db` as the registry spreadsheet ID.
- In registry sheet `tables`, find row where column A is exactly `Càrrega lectiva`.
- Use column B from that row as the Càrrega lectiva spreadsheet ID.
- Open sheet `distribucio` in that spreadsheet.
- Open sheet `professors` in the same `Càrrega lectiva` spreadsheet for teacher email enrichment.
- Do not write anything.

Expected `distribucio` structure:

- teacher names: row 2, from column M onward;
- lesson header row: row where column A exactly equals `Curs` and column C exactly equals `Assignatura`;
- lesson data rows: rows below the lesson header row;
- course: column A;
- group: column B;
- subject: column C;
- teacher allocation cells: columns M onward.

Expected `professors` structure in `Càrrega lectiva`:

- `CORREU INSTIT`: column L;
- teacher lookup key/name: column Q.

The teacher names from `distribucio` row 2 are matched against `professors` column Q to add `CORREU INSTIT` from column L.

Class combo generation:

1. Find the lesson header row.
2. Read rows below that header.
3. Read `Curs` from column A and `Grup` from column B.
4. Skip rows where `Curs` or `Grup` is empty.
5. Build class labels as `Curs + " " + Grup`.
6. Expand comma-separated groups into individual labels.
7. Expand `TOTS` using known groups.
8. Remove duplicates.
9. Sort naturally: ESO first, BAT next, other groups last; within each stage, by course number and group.

Known `TOTS` expansion:

- `1ESO`: `A`, `B`, `C`, `D`, `E`;
- `2ESO`: `A`, `B`, `C`, `D`, `E`;
- `3ESO`: `A`, `B`, `C`, `D`, `E`;
- `4ESO`: `A`, `B`, `C`, `D`, `E`;
- `1BAT`: `A`, `B`, `C`;
- `2BAT`: `A`, `B`.

When a class is selected:

1. Parse the selected label into course and group.
2. Keep rows where column A matches course and column B includes the selected group.
3. A row matches when `Grup` exactly equals the group, is comma-separated and includes it, or is `TOTS` and the selected group belongs to the course's known groups.
4. For each matching row, read subject from column C.
5. Scan columns M onward.
6. If an allocation cell is numeric and greater than 0, row 2 of that column is the teacher for that subject.
7. Collapse duplicate teacher/subject pairs using key `teacherName + "::" + subject`.
8. Keep result order from `distribucio`: lesson rows are read top-to-bottom, and teacher allocations are scanned from column M onward.

Display:

- combo box with all classes/groups;
- results table with columns `Teacher`, `Subject`, and `CORREU INSTIT`;
- if no class is selected, show no results;
- if a class has no rows, show `No teachers found for this class.`.

Read-only rule:

- allowed: read script property `db`, open registry spreadsheet, open `Càrrega lectiva`, read `distribucio`, read `professors`, render UI;
- forbidden: writing cell values, clearing cells, modifying sheets, creating sheets, editing properties.

## `Estructura` View

- Section exists in the left menu and page shell.
- Behavior is reserved for a later phase.

## XLSX Export

The floating bottom-right export button:

- is visible on `Llistat`;
- uses Bootstrap Icons class `bi-file-earmark-spreadsheet`;
- exports the current filtered list to XLSX format on `Llistat`;
- exports the selected class teacher/subject/email table to XLSX format on `Per classe`.

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
- `getClassOptions()`: returns class labels from `distribucio`.
- `getTeachersForClass(classLabel)`: returns teacher/subject/email rows for the selected class.
- `createTeacherListXlsx(filters)`: exports the current filtered list as XLSX.
- `createClassTeachersXlsx(classLabel)`: exports the selected class teacher/subject/email table as XLSX.
- `getTeacherDbSpreadsheet_()`: resolves and opens DB.
- `getTeacherDbSheet_()`: opens DB sheet `Llista`.

## Manifest Scopes

Required scopes:

- `https://www.googleapis.com/auth/script.storage`
- `https://www.googleapis.com/auth/spreadsheets`
- `https://www.googleapis.com/auth/drive`
- `https://www.googleapis.com/auth/script.external_request`
