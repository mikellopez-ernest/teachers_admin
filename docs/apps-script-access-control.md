# Apps Script Role-Based Access Control

This document describes the reusable access-control system used in `teacher_admin`.

Use it when another Google Apps Script web app must allow access only to selected people from the `iernestlluch.cat` workspace, based on roles/càrrecs stored in `Càrrega lectiva`.

## Goal

The web app should:

1. Be reachable only by users from the `iernestlluch.cat` domain.
2. Read the signed-in user's email.
3. Read a script property named `access_granted`.
4. Resolve the allowed roles/càrrecs into people.
5. Resolve those people into institutional email addresses.
6. Allow access only if the signed-in user's email is in that resolved list.
7. Show a no-access page otherwise.
8. Protect server-side methods too, not only `doGet()`.

## Script Property

Create this Apps Script property in the script that needs protection:

```text
access_granted
```

Its value is a comma-separated list of allowed càrrecs.

Example:

```text
Coord. 3ESO,COCOBE
```

The names must match `Càrrega lectiva -> carrecs` column A.

## Data Sources

The system expects the script to be able to resolve `Càrrega lectiva`.

In this repo, scripts use a registry spreadsheet:

- script property `Tables` contains the spreadsheet ID of the registry;
- registry sheet name is `tables`;
- registry column A contains logical spreadsheet names;
- registry column B contains spreadsheet IDs;
- the row where column A is `Càrrega lectiva` contains the spreadsheet ID needed here.

If another script uses a different registry property name, adapt `CONFIG.tablesPropertyName`.

### `Càrrega lectiva -> carrecs`

Used to map allowed role names to assigned people.

Required columns:

| Column | Header / meaning |
| --- | --- |
| A | `carrec` |
| D | `asignado?` |

For each role in `access_granted`, find the row where column A matches exactly. Then read column D.

Column D can contain one person or multiple people separated by commas.

Example:

```text
Xavi Guitart Naya,Rubén Guim Domenech
```

### `Càrrega lectiva -> professors`

Used to map assigned people to institutional email addresses.

Required columns:

| Column | Header / meaning |
| --- | --- |
| L | `CORREU INSTIT` |
| Q | full teacher name / lookup key |

For each assigned person read from `carrecs` column D, find the row where `professors` column Q matches the person name. Then read email from column L.

## Deployment Settings

In `appsscript.json`, restrict the web app to the domain:

```json
"webapp": {
  "executeAs": "USER_DEPLOYING",
  "access": "DOMAIN"
}
```

Notes:

- `access: "DOMAIN"` blocks users outside `iernestlluch.cat`.
- `executeAs: "USER_DEPLOYING"` lets the script keep using the owner's permissions for spreadsheet reads/writes.
- The server-side check still uses `Session.getActiveUser().getEmail()` to identify the visitor.
- If a script needs the visitor's own Drive/Spreadsheet permissions instead, use `executeAs: "USER_ACCESSING"`, but then each allowed user may need file access and may need to authorize scopes.

## Required Config Constants

Add or adapt these constants:

```js
const CONFIG = {
  tablesPropertyName: 'Tables',
  accessGrantedPropertyName: 'access_granted',
  tablesSheetName: 'tables',
  workloadRegistryName: 'Càrrega lectiva',
  workloadProfessorsSheetName: 'professors',
  workloadCarrecsSheetName: 'carrecs',
};

const WORKLOAD_PROFESSORS_COLUMNS = {
  correuInstit: 12,
  teacherKey: 17,
};

const CARRECS_COLUMNS = {
  carrec: 1,
  asignado: 4,
};
```

If the target script already has a `CONFIG` object, merge these fields into it.

## Required Helper Methods

Copy these helper methods into the target script.

```js
function assertUserAccess_() {
  const access = getAccessDecision_();
  if (!access.allowed) {
    throw new Error(access.message);
  }
  return access;
}

function getAccessDecision_() {
  try {
    const userEmail = normalizeEmail_(Session.getActiveUser().getEmail());
    if (!userEmail) {
      return {
        allowed: false,
        email: '',
        message: 'No s\'ha pogut identificar el correu de l\'usuari actiu.',
      };
    }

    const roles = getAccessGrantedRoles_();
    if (roles.length === 0) {
      return {
        allowed: false,
        email: userEmail,
        message: `Falta configurar la propietat de script "${CONFIG.accessGrantedPropertyName}".`,
      };
    }

    const peopleByRole = getPeopleByAccessRole_();
    const people = [];
    roles.forEach((role) => {
      const assignedPeople = peopleByRole.get(normalizeText_(role)) || [];
      assignedPeople.forEach((person) => people.push(person));
    });

    const authorizedEmails = getEmailsForPeople_(people);
    const allowed = authorizedEmails.has(userEmail);

    return {
      allowed,
      email: userEmail,
      roles,
      people,
      message: allowed
        ? 'Acces autoritzat.'
        : 'No tens permisos per accedir a aquesta aplicacio.',
    };
  } catch (error) {
    return {
      allowed: false,
      email: normalizeEmail_(Session.getActiveUser().getEmail()),
      message: error && error.message ? error.message : String(error),
    };
  }
}

function getAccessGrantedRoles_() {
  return splitCommaList_(
    PropertiesService.getScriptProperties().getProperty(CONFIG.accessGrantedPropertyName)
  );
}

function getPeopleByAccessRole_() {
  const sheet = getWorkloadCarrecsSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Map();

  const values = sheet.getRange(2, 1, lastRow - 1, CARRECS_COLUMNS.asignado).getValues();
  const peopleByRole = new Map();

  values.forEach((row) => {
    const roleName = toDisplayString_(row[CARRECS_COLUMNS.carrec - 1]);
    if (!roleName) return;

    peopleByRole.set(
      normalizeText_(roleName),
      splitCommaList_(row[CARRECS_COLUMNS.asignado - 1])
    );
  });

  return peopleByRole;
}

function getEmailsForPeople_(people) {
  const sheet = getWorkloadProfessorsSheet_();
  const lastRow = sheet.getLastRow();
  const emails = new Set();
  if (lastRow < 2 || people.length === 0) return emails;

  const peopleSet = new Set(people.map((person) => normalizeText_(person)));
  const values = sheet.getRange(2, 1, lastRow - 1, WORKLOAD_PROFESSORS_COLUMNS.teacherKey).getValues();

  values.forEach((row) => {
    const teacherKey = normalizeText_(row[WORKLOAD_PROFESSORS_COLUMNS.teacherKey - 1]);
    if (!peopleSet.has(teacherKey)) return;

    const email = normalizeEmail_(row[WORKLOAD_PROFESSORS_COLUMNS.correuInstit - 1]);
    if (email) emails.add(email);
  });

  people.forEach((person) => {
    const directEmail = normalizeEmail_(person);
    if (directEmail.indexOf('@') !== -1) emails.add(directEmail);
  });

  return emails;
}

function createAccessDeniedOutput_(access) {
  const email = access && access.email ? access.email : 'usuari no identificat';
  const message = access && access.message
    ? access.message
    : 'No tens permisos per accedir a aquesta aplicacio.';

  return HtmlService
    .createHtmlOutput(`
      <!doctype html>
      <html lang="ca">
        <head>
          <base target="_top">
          <meta charset="utf-8">
          <meta name="viewport" content="width=device-width, initial-scale=1">
          <title>Acces no autoritzat</title>
          <style>
            body {
              margin: 0;
              min-height: 100vh;
              display: grid;
              place-items: center;
              font-family: Arial, sans-serif;
              background: #f5f7fb;
              color: #17202a;
            }
            main {
              width: min(520px, calc(100vw - 32px));
              border: 1px solid #d8dee9;
              background: #fff;
              padding: 28px;
              box-shadow: 0 18px 45px rgba(20, 31, 47, 0.12);
            }
            h1 {
              margin: 0 0 12px;
              font-size: 24px;
            }
            p {
              margin: 8px 0;
              line-height: 1.5;
            }
            .email {
              font-family: monospace;
              color: #465466;
            }
          </style>
        </head>
        <body>
          <main>
            <h1>Acces no autoritzat</h1>
            <p>${escapeHtml_(message)}</p>
            <p class="email">${escapeHtml_(email)}</p>
          </main>
        </body>
      </html>
    `)
    .setTitle('Acces no autoritzat')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
```

## Required Spreadsheet Helpers

If the target script already has registry helpers, adapt these to its naming.

```js
function getWorkloadSpreadsheet_() {
  return getRegisteredSpreadsheet_(CONFIG.workloadRegistryName);
}

function getWorkloadProfessorsSheet_() {
  const spreadsheet = getWorkloadSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(CONFIG.workloadProfessorsSheetName);

  if (!sheet) {
    throw new Error(`No s'ha trobat el full "${CONFIG.workloadProfessorsSheetName}" a Càrrega lectiva.`);
  }

  return sheet;
}

function getWorkloadCarrecsSheet_() {
  const spreadsheet = getWorkloadSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(CONFIG.workloadCarrecsSheetName);

  if (!sheet) {
    throw new Error(`No s'ha trobat el full "${CONFIG.workloadCarrecsSheetName}" a Càrrega lectiva.`);
  }

  return sheet;
}

function getRegisteredSpreadsheet_(registryName) {
  const properties = PropertiesService.getScriptProperties();
  const tablesSpreadsheetId = String(
    properties.getProperty(CONFIG.tablesPropertyName) || ''
  ).trim();

  if (!tablesSpreadsheetId) {
    throw new Error(`Falta la propietat de script "${CONFIG.tablesPropertyName}".`);
  }

  const tablesSpreadsheet = SpreadsheetApp.openById(tablesSpreadsheetId);
  const tablesSheet = tablesSpreadsheet.getSheetByName(CONFIG.tablesSheetName);

  if (!tablesSheet) {
    throw new Error(`No s'ha trobat el full "${CONFIG.tablesSheetName}".`);
  }

  const lastRow = tablesSheet.getLastRow();
  if (lastRow < 1) {
    throw new Error(`El full "${CONFIG.tablesSheetName}" esta buit.`);
  }

  const values = tablesSheet.getRange(1, 1, lastRow, 2).getValues();
  const match = values.find((row) => row[0] === registryName);

  if (!match) {
    throw new Error(
      `No s'ha trobat "${registryName}" al full "${CONFIG.tablesSheetName}".`
    );
  }

  const spreadsheetId = String(match[1] || '').trim();
  if (!spreadsheetId) {
    throw new Error(`La fila "${registryName}" no te ID a la columna B.`);
  }

  return SpreadsheetApp.openById(spreadsheetId);
}
```

If the target script already has `getDbSpreadsheet_()`, update it to call `getRegisteredSpreadsheet_()`:

```js
function getDbSpreadsheet_() {
  return getRegisteredSpreadsheet_(CONFIG.dbRegistryName);
}
```

## Required Utility Helpers

Copy these if the target script does not already have equivalent helpers.

```js
function toDisplayString_(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value).trim();
}

function normalizeText_(value) {
  return toDisplayString_(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLocaleLowerCase('ca');
}

function normalizeEmail_(value) {
  return toDisplayString_(value).toLocaleLowerCase('ca');
}

function splitCommaList_(value) {
  return toDisplayString_(value)
    .split(',')
    .map((item) => toDisplayString_(item))
    .filter(Boolean);
}

function escapeHtml_(value) {
  return toDisplayString_(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
```

## How To Apply The Check

### Protect `doGet()`

At the top of `doGet()`, add:

```js
function doGet() {
  const access = getAccessDecision_();
  if (!access.allowed) return createAccessDeniedOutput_(access);

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Your app title')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
```

### Protect server-side methods

Every function callable from the browser with `google.script.run` should start with:

```js
assertUserAccess_();
```

Examples:

```js
function getData() {
  assertUserAccess_();
  // read and return data
}

function saveData(payload) {
  assertUserAccess_();
  // write data
}

function exportData(filters) {
  assertUserAccess_();
  // export data
}
```

This matters because a user could try to call server functions directly even if the page did not render.

## Permission Grant Function

If the script has a manual permission helper, make it touch:

- script property `access_granted`;
- active user email;
- `Càrrega lectiva -> carrecs`;
- `Càrrega lectiva -> professors`.

Example:

```js
function grantRequiredPermissions() {
  const properties = PropertiesService.getScriptProperties();
  properties.getProperty(CONFIG.tablesPropertyName);
  properties.getProperty(CONFIG.accessGrantedPropertyName);
  Session.getActiveUser().getEmail();

  getWorkloadCarrecsSheet_().getRange(1, 1).getValue();
  getWorkloadProfessorsSheet_().getRange(1, 1).getValue();

  return {
    ok: true,
    message: 'Permisos concedits correctament.',
  };
}
```

## Checklist For A New Script

1. Add `access_granted` script property.
2. Ensure the script can resolve `Càrrega lectiva` from the registry.
3. Add config constants for:
   - `accessGrantedPropertyName`;
   - `workloadRegistryName`;
   - `workloadProfessorsSheetName`;
   - `workloadCarrecsSheetName`.
4. Add column constants for:
   - `professors` column L;
   - `professors` column Q;
   - `carrecs` column A;
   - `carrecs` column D.
5. Copy the authorization helpers.
6. Copy or adapt the registry/sheet helpers.
7. Copy utility helpers if missing.
8. Add the `doGet()` gate.
9. Add `assertUserAccess_()` to all browser-callable server methods.
10. Change `appsscript.json` web app access to `DOMAIN`.
11. Push with `clasp push -f`.
12. Deploy over the existing deployment ID if the URL must stay the same.
13. Test with:
    - one allowed user;
    - one same-domain but not allowed user;
    - optionally one non-domain user.

## Troubleshooting

### A user outside `access_granted` can still access

Check:

- the script was deployed after the code change;
- the user is testing the deployed `/exec` URL, not an old `/dev` URL;
- `doGet()` calls `getAccessDecision_()`;
- all `google.script.run` methods call `assertUserAccess_()`;
- the user is not accidentally included through another role in `access_granted`.

### Active user email is blank

Check:

- web app access is restricted to domain users;
- the user is signed in with an `iernestlluch.cat` account;
- the app is not published in a mode where Apps Script cannot reveal the user identity.

### A valid user is denied

Check:

- the càrrec text in `access_granted` exactly matches `carrecs` column A, ignoring accents/case only after normalization;
- the assigned person in `carrecs` column D matches `professors` column Q after normalization;
- `professors` column L has the correct institutional email;
- column D names are separated by commas, not semicolons.

### The app can read authorization sheets but cannot access its own data

If the app executes as `USER_ACCESSING`, allowed users need file access and may need to authorize scopes. If owner-level access is intended, use `executeAs: "USER_DEPLOYING"` and keep the explicit authorization check.
