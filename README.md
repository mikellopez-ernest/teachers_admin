# Professors DB Web App

Google Apps Script web app for reading professor data from a spreadsheet-based DB and managing two boolean status fields.

## Data Model

The script property `Tables` must contain the ID of a registry spreadsheet. That registry spreadsheet must have a sheet named `tables` where:

- Column A contains logical table names.
- Column B contains spreadsheet IDs.

The app finds the row where column A is `Dades de professors`; column B in that row is the DB spreadsheet ID.

The DB spreadsheet must have a sheet named `Llista`. Row 1 is treated as the header row, and data starts at row 2.

## Displayed Columns

| UI column | DB source |
| --- | --- |
| `ESP` | column A |
| `DEPT.` | column B |
| `NOM SENCER` | columns C, D, E joined with spaces |
| `SITUACIO` | column G |
| `DNI` | column H |
| `TELF` | column I |
| `CORREU` | column K |
| `XTEC` | column J |

## Status Fields

- `BAIXA?` reads DB column L as a boolean or string `TRUE`; writes a real boolean.
- `NO ACTIUS` treats DB column N as active-state data: boolean/string `TRUE` means active, so unchecked rows are shown by default; the checked filter shows rows where column N is not true.

## Features

- Web endpoint served by `doGet`.
- Department filter populated from DB column B.
- Text filter over the full professor name.
- `RESET` clears only the department and name filters.
- Clickable table headers toggle A-Z / Z-A sorting.
- `NOM SENCER` displays name plus surnames, but sorts by first surname.
- Row checkboxes select rows for bulk updates.
- Header checkbox selects or deselects all currently filtered rows.
- Clicking a teacher name opens a detail window.
- Detail windows can be changed to edit forms with `Edit`, then saved or cancelled.
- No delete option is provided.
- `Donar de baixa` / `Donar de alta` updates column L.
- `Desactivar` / `Activar` updates column N.
- `Exportar` downloads a spreadsheet-compatible CSV file from selected rows.
- Action buttons warn if no row is selected.

## Local Development

This project is connected to Apps Script with `clasp`.

```sh
clasp push -f
clasp deploy --description "Deploy professors DB web app"
```

The script ID is stored in `.clasp.json`.
