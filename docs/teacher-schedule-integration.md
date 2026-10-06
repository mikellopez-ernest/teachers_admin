# Teacher Schedule Integration Manual

This manual describes how another Google Apps Script project can find a teacher's current schedule from their full name and include that schedule in a document.

It describes the database and rendering behavior implemented in the schedule viewer. The recommended data source is `Horaris -> schedule_cache`: a generated timetable containing resolved teacher names, active substitutes, and subject full names.

## 1. Connect To The Database

### Script property `db`

In the consuming Apps Script project, create a script property named `db`. Its value must be the ID of the **registry spreadsheet**, not the ID of the Horaris spreadsheet.

Set it under **Project Settings -> Script Properties**.

```text
Property: db
Value: <database registry spreadsheet ID>
```

Each Apps Script project has its own script properties. A property configured in the schedule viewer is not automatically available to another project.

The account executing the consuming script must have read access to the registry spreadsheet and the database spreadsheets it opens. Reading these sheets directly does not require the schedule viewer's web app URL or its rebuild token.

### Registry spreadsheet

The registry has a sheet named `tables`:

| Column | Content |
| --- | --- |
| A | Logical database name, for example `Horaris` |
| B | ID of the spreadsheet containing that database |

Example registry entries:

```text
Horaris             <Horaris spreadsheet ID>
Dades de professors <teacher spreadsheet ID>
Càrrega lectiva     <subject catalog spreadsheet ID>
```

The IDs in column B identify **spreadsheets**, not sheet tabs. First open the spreadsheet by ID, then select its sheet by name.

### Configure spreadsheet -> sheet relationships

Keep the logical spreadsheet name and the sheet name together in the consuming program's configuration:

```javascript
const TEACHER_SCHEDULE_DB = {
  registryProperty: 'db',
  registrySheet: 'tables',
  schedule: { database: 'Horaris', sheet: 'schedule_cache' },
  teachers: { database: 'Dades de professors', sheet: 'Llista' },
};
```

Do not hardcode database spreadsheet IDs. Resolve them through `db -> tables` on each execution, or through a deliberately managed application cache.

## 2. Tables To Query

### Required: Horaris -> schedule_cache

This is the main source for current timetable display. Read it once per document generation and filter the rows in memory.

The cache already supplies:

- The original/source teacher and the effective teacher.
- Substitution results for active leaves.
- Subject short codes and full names.
- Groups, classrooms, weekdays, and slot numbers.

The schedule viewer uses the **effective teacher** for teacher selection and schedule display.

### Recommended for full-name resolution: Dades de professors -> Llista

Use this table to turn an input full name into a stable teacher code.

| Column | Header | Use |
| --- | --- | --- |
| A | `ESP` | Original teacher code used by leave mapping |
| C | `NOM` | First name |
| D | `COGNOM1` | First surname |
| E | `COGNOM2` | Second surname |
| F | `REDUIT` | Teacher code used in the timetable and cache |
| N | `ACTIU` | Active flag |
| O | `BAIXA?` | Leave flag |
| P | `SUBST?` | Substitute flag |

The actual sheet name is **`Llista`**.

Build the full name by joining nonblank `NOM`, `COGNOM1`, and `COGNOM2` with a single space:

```text
Mikel + López + Villarroya -> Mikel López Villarroya
```

Map columns by their headers rather than assuming they always remain in the same physical positions.

Do not exclude a teacher from schedule lookup solely because of `ACTIU`, `BAIXA?`, or `SUBST?`. The effective rows in the cache determine whose current timetable is available.

### Upstream tables: usually no need to query them

| Spreadsheet -> sheet | Role |
| --- | --- |
| `Horaris -> GPU001` | Raw, headerless timetable used to build the cache |
| `Dades de professors -> leave_absence` | Original-teacher/substitute mapping used during rebuilding |
| `Càrrega lectiva -> assignatures` | Subject catalog used to populate `subject_full_name` |

The consuming script should normally read the cache rather than repeat its rebuilding logic.

For context, `GPU001` has seven columns: row identifier, group, teacher `REDUIT`, subject code, classroom, weekday number, slot number. Its first row is data, not a header.

In `assignatures`, `short_name` matches the timetable subject code and `full_name` supplies the display name. The other catalog headers are `ETAPA`, `untis_name`, and `true_subject`. Normal timetable rendering includes all cached entries; it does not filter out guards, meetings, or tutoring using `true_subject`.

## 3. schedule_cache Columns

Row 1 is the header. Data starts in row 2.

```csv
row_id,group,source_teacher_code,source_teacher_name,source_teacher_original_code,effective_teacher_code,effective_teacher_name,teacher_was_substituted,subject_code,subject_full_name,classroom,day,slot
```

These are the 13 columns generated by the current cache builder. Resolve them by header name and tolerate additional columns from other applications. The cache builder replaces the generated contents, so application-specific configuration should live in a separate sheet.

| Position | Header | Type and meaning |
| --- | --- | --- |
| A / 1 | `row_id` | String or number copied from GPU001 column 1. Identifies a source schedule unit. It is **not a unique cache-row key**: several groups or occurrences can share it. |
| B / 2 | `group` | School group, such as `1A`, `2B`, or `4C`. May be blank for activities without a student group. |
| C / 3 | `source_teacher_code` | Original timetable teacher's `REDUIT`, copied from GPU001 column 3. For example `ESCANG`. |
| D / 4 | `source_teacher_name` | Source teacher's full name, assembled from `Llista.NOM`, `COGNOM1`, and `COGNOM2`. If the teacher cannot be resolved, a code fallback can appear. |
| E / 5 | `source_teacher_original_code` | Source teacher's `ESP` from Llista column A. This is the expected identifier for `leave_absence.teacher_code`. It is a different identifier from `REDUIT`. |
| F / 6 | `effective_teacher_code` | `REDUIT` of the teacher currently responsible for this row after active substitution resolution. Use this field for a current teacher's timetable. |
| G / 7 | `effective_teacher_name` | Full name of that effective teacher. Use it when displaying who teaches the lesson. |
| H / 8 | `teacher_was_substituted` | Real boolean indicating whether the effective teacher differs from the source teacher because of a resolved active leave. An exported CSV may represent it as `TRUE` or `FALSE`. |
| I / 9 | `subject_code` | Original short subject/activity code, for example `ANG`, `DIG`, `GUARDIA`, or `TUT`. Use it for colour classification and code-based logic. |
| J / 10 | `subject_full_name` | Full display name from `assignatures.full_name`. The cache builder falls back to `subject_code` when no usable translation exists. |
| K / 11 | `classroom` | Room code or name, copied from GPU001 column 5. May be blank. |
| L / 12 | `day` | Weekday number: `1` Monday, `2` Tuesday, `3` Wednesday, `4` Thursday, `5` Friday. |
| M / 13 | `slot` | Daily timetable slot number, from `1` to `12`. This is an ordinal, not a clock time. |

The cache does not contain teacher email addresses, lesson start/end clock times, or leave dates. Query the appropriate source separately if your document needs those fields. Do not invent clock times from the slot numbers.

## 4. Find A Schedule From A Full Name

### Recommended lookup sequence

1. Read `Llista` and `schedule_cache` once.
2. Build each teacher's full name from the three name fields.
3. Normalize the input and candidate names for comparison: trim, collapse repeated spaces, ignore case, and remove accents. Keep the original spelling for display.
4. Find an exact normalized full-name match in Llista.
5. Obtain that teacher's `REDUIT`.
6. Filter schedule_cache by normalized `effective_teacher_code == REDUIT`.
7. Validate `day` and `slot`, then place the matching rows in the timetable grid.

Avoid substring matching. A first name or one surname alone is not a reliable teacher identifier.

If multiple teachers share the same normalized full name, require a `REDUIT` or another identifying value. Do not silently pick the first record.

If the name cannot be found in Llista, a fallback lookup against distinct `effective_teacher_name` values in the cache can be useful. Accept it only if all matches resolve to one effective teacher code. The example below implements that fallback.

### Active substitutes

The cache reflects leaves active on the rebuild date, evaluated using `Europe/Madrid`. A leave starts on `start_date`, ends on `end_date` inclusively, and remains open if `end_date` is blank.

The builder maps timetable `REDUIT -> Llista.ESP -> leave_absence.teacher_code`, then resolves the substitute using `leave_absence.substitute_code`, which is a **REDUIT**. For compatibility, the current builder also accepts a leave teacher code matching the source REDUIT if the ESP lookup fails. New integrations should preserve the ESP/REDUIT distinction.

Example:

```text
source_teacher_code:          ESCANG
source_teacher_name:          Gemma Escudé Pont
source_teacher_original_code: Ang
effective_teacher_code:       MAREXP
effective_teacher_name:       Alba Martínez López
teacher_was_substituted:      TRUE
```

For Alba's current schedule, filter by `effective_teacher_code = MAREXP`. Gemma's name remaining in source fields is expected.

For Gemma's current schedule, filtering by `ESCANG` may return no rows while Alba covers her timetable. That is a valid result. Report that there are no current effective timetable entries for the selected teacher.

If a document explicitly needs the original teacher's underlying timetable even while it is covered, filter by `source_teacher_code` instead and label that result clearly as a source timetable. Do not silently switch lookup modes when the effective result is empty.

When a substitute is selected, their effective cache rows already contain the covered timetable. There is no need to find the original teacher again in the consuming script.

### Cache freshness

The cache is a snapshot, not a live join. It is rebuilt after schedule uploads, by the daily rebuild trigger when installed, and by authorized manual/API rebuilds.

Changes to teacher names, subject names, or leave rows are reflected only after rebuilding. A read request does not refresh substitution dates automatically.

There is no `updated_at` column in schedule_cache. Do not use the document generation timestamp as proof of cache freshness. If freshness is required, arrange an authorized rebuild before reading or consult the producer's rebuild diagnostics.

Treat this sheet as read-only. Do not delete rows or write document-specific data into it.

## 5. Paint The Teacher Schedule

### Grid layout

- Document title: the selected teacher's full name.
- Five timetable columns: Monday to Friday (`Dilluns`, `Dimarts`, `Dimecres`, `Dijous`, `Divendres` if using Catalan).
- Twelve timetable rows: slots `1` through `12`.
- An additional narrow first column can show the slot number.
- Each entry goes into the cell at its numeric `slot` and `day`.
- An empty cell stays blank.

For a screen-style schedule, keep all 12 rows. For a printed document matching the viewer's print behavior, omit a slot row only if **all five cells in that row are empty**. Keep the five weekday columns and retain the original slot labels; never renumber the remaining rows.

The viewer's print layout uses **A4 landscape**, with 10 mm margins.

### Group entries within each cell

After selecting the teacher and partitioning rows by day and slot, group entries by:

```text
(display subject name, classroom)
```

The teacher is already fixed by the lookup, and day/slot are already fixed by the cell.

For each such group, collect distinct nonblank school-group names and join them with `, `. This preserves a single bubble for a shared lesson:

```text
DIGITALITZACIÓ
Grup: 2A, 2B, 2C, 2D, 2E
Classe: 2E
```

Different classrooms remain separate bubbles even when the subject is the same. Different subjects also remain separate. Several legitimate items can therefore coexist in one cell.

The current viewer groups by the display subject name, not row_id. For exact visual parity, use the same rule. Keep subject_code on the grouped item for colour selection; the current viewer retains the first contributing row's colour when display names collide.

Do not globally deduplicate teacher/subject/group combinations: that would erase occurrences on other weekdays or slots. Deduplication for display happens **inside a particular cell**. Do not remove lower-hour teachers or infer other assignment rules when reading this cache.

### Text inside a bubble

Show these lines vertically:

1. **Subject full name**, falling back to subject_code.
2. `Grup: ` followed by the distinct group names, only when at least one exists.
3. `Classe: ` followed by the classroom, only when nonblank.

Do not show the teacher's name inside a teacher-schedule bubble; it is already the document title. Omit an entire absent line, including its prefix. Never display `Grup: -` or `Classe: -`.

Centre the stack of bubbles vertically within the table cell. Wrap long subject names and allow rows to grow when needed. Keep items in the same cell visually separate and avoid clipping or overlapping their text.

### Colours

Classify using **subject_code**, not its translated full name. Trim and compare codes in uppercase.

| Codes | Colour | Background | Accent border |
| --- | --- | --- | --- |
| `GUARDIA`, `GUARDIA_PATI` | Orange | `#fff4e5` | `#f59f00` |
| `RC_ESO`, `RC_FP_BAT`, `RDEP`, `RDIM1`, `RDIM2`, `RDIR`, `REAP`, `REC`, `CARREC`, `FCT`, `TREC` | Blue | `#eef6ff` | `#4dabf7` |
| `TUT`, `TUT_FAMILIES` | Pink | `#fff0f6` | `#f06595` |
| `3R` | Yellow | `#fff9db` | `#f2c94c` |
| Every other code | Green | `#edf7f6` | `#7fc7bd` |

Use a bold subject title, smaller supporting text, and a coloured left border. The viewer uses title text `#134e4a` and supporting text `#344054`.

For a Google Docs table, apply the category background to the individual entry when the document API/layout allows it. If several differently coloured items share a cell, preserve their individual colour distinction with separate blocks or a colour accent per item. Avoid assigning the whole cell the colour of just one of its items.

For HTML/PDF rendering, these styles approximate the viewer:

```css
@page { size: A4 landscape; margin: 10mm; }
table { width: 100%; border-collapse: collapse; table-layout: fixed; }
th, td { border: 1px solid #d8dee8; padding: 4px; }
td { vertical-align: middle; }
.cell-items { display: grid; gap: 4px; align-content: center; }
.schedule-item {
  padding: 4px;
  border-left: 3px solid var(--item-accent);
  border-radius: 4px;
  background: var(--item-bg);
  overflow-wrap: anywhere;
}
.subject { font-size: 9px; font-weight: bold; color: #134e4a; }
.meta { font-size: 9px; color: #344054; }
* { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
```

HTML/CSS support varies across PDF converters. Verify the output in the document generator you actually use. Escape database text before inserting it into HTML; using DOM `textContent` is safe in a browser.

## 6. Copyable Apps Script Lookup And Grid Builder

This example reads the registry, resolves a teacher name, and returns a renderer-independent grid. It does not create or modify spreadsheets, rebuild the cache, or create a document.

Call:

```javascript
const timetable = getTeacherDocumentSchedule('Alba Martínez López');
// timetable.teacherName: document title
// timetable.rows: all 12 slots, with five cells each
// timetable.printRows: only slots containing at least one item
// timetable.scheduleRows: original matching cache records
```

Paste the following functions into the consuming project. They use a `ts` prefix for private helpers to reduce collisions with other application code.

```javascript
function getTeacherDocumentSchedule(fullName) {
  const config = {
    registryProperty: 'db',
    registrySheet: 'tables',
    schedule: { database: 'Horaris', sheet: 'schedule_cache' },
    teachers: { database: 'Dades de professors', sheet: 'Llista' },
  };
  if (!tsText_(fullName)) throw new Error('A teacher full name is required.');

  const dbId = tsText_(PropertiesService.getScriptProperties()
    .getProperty(config.registryProperty));
  if (!dbId) throw new Error('Missing script property "db".');

  const registrySheet = SpreadsheetApp.openById(dbId)
    .getSheetByName(config.registrySheet);
  if (!registrySheet) throw new Error('Registry sheet "tables" is missing.');

  const registry = new Map();
  registrySheet.getDataRange().getDisplayValues().forEach(function(row) {
    const name = tsText_(row[0]);
    const id = tsText_(row[1]);
    if (!name || !id) return;
    if (registry.has(name)) {
      throw new Error('Duplicate database registry entry: ' + name);
    }
    registry.set(name, id);
  });

  const teachers = tsReadTable_(registry, config.teachers,
    ['NOM', 'COGNOM1', 'COGNOM2', 'REDUIT']);
  const cache = tsReadTable_(registry, config.schedule, [
    'row_id', 'group', 'source_teacher_code', 'source_teacher_name',
    'source_teacher_original_code', 'effective_teacher_code',
    'effective_teacher_name', 'teacher_was_substituted',
    'subject_code', 'subject_full_name', 'classroom', 'day', 'slot',
  ]);

  const nameKey = tsKey_(fullName);
  const matches = teachers.filter(function(teacher) {
    return tsKey_(tsFullName_(teacher)) === nameKey;
  });
  let teacherCode;
  let teacherName;

  if (matches.length > 1) {
    throw new Error('Ambiguous teacher name. Select the teacher by REDUIT.');
  }
  if (matches.length === 1) {
    teacherCode = tsText_(matches[0].REDUIT);
    teacherName = tsFullName_(matches[0]);
    if (!teacherCode) throw new Error('Matching teacher has no REDUIT.');
    const codeMatches = teachers.filter(function(teacher) {
      return tsKey_(teacher.REDUIT) === tsKey_(teacherCode);
    });
    if (codeMatches.length > 1) {
      throw new Error('Duplicate teacher REDUIT in Llista: ' + teacherCode);
    }
  } else {
    const fallback = cache.filter(function(row) {
      return tsKey_(row.effective_teacher_name) === nameKey &&
        tsText_(row.effective_teacher_code);
    });
    const codes = new Set(fallback.map(function(row) {
      return tsKey_(row.effective_teacher_code);
    }));
    if (codes.size !== 1) {
      throw new Error(codes.size ? 'Ambiguous teacher name in cache.' :
        'Teacher not found: ' + fullName);
    }
    teacherCode = tsText_(fallback[0].effective_teacher_code);
    teacherName = tsText_(fallback[0].effective_teacher_name);
  }

  const invalidRows = [];
  const scheduleRows = cache.filter(function(row) {
    if (tsKey_(row.effective_teacher_code) !== tsKey_(teacherCode)) return false;
    const day = Number(row.day);
    const slot = Number(row.slot);
    const valid = Number.isInteger(day) && day >= 1 && day <= 5 &&
      Number.isInteger(slot) && slot >= 1 && slot <= 12;
    if (!valid) invalidRows.push(row);
    return valid;
  });

  const rows = Array.from({ length: 12 }, function(_, index) {
    return { slot: index + 1, cells: Array.from({ length: 5 }, function() {
      return { items: [], index: new Map() };
    }) };
  });

  scheduleRows.forEach(function(row) {
    const cell = rows[Number(row.slot) - 1].cells[Number(row.day) - 1];
    const title = tsText_(row.subject_full_name) ||
      tsText_(row.subject_code) || 'No subject';
    const classroom = tsText_(row.classroom);
    const key = JSON.stringify([title, classroom]);
    let item = cell.index.get(key);
    if (!item) {
      item = {
        subjectName: title,
        subjectCode: tsText_(row.subject_code),
        groups: [],
        classroom: classroom,
        colour: tsSubjectColour_(row.subject_code),
      };
      cell.index.set(key, item);
      cell.items.push(item);
    }
    const group = tsText_(row.group);
    if (group && item.groups.indexOf(group) === -1) item.groups.push(group);
  });

  rows.forEach(function(row) {
    row.cells.forEach(function(cell) { delete cell.index; });
  });
  return {
    teacherCode: teacherCode,
    teacherName: teacherName,
    days: ['Dilluns', 'Dimarts', 'Dimecres', 'Dijous', 'Divendres'],
    scheduleRows: scheduleRows,
    invalidRows: invalidRows,
    rows: rows,
    printRows: rows.filter(function(row) {
      return row.cells.some(function(cell) { return cell.items.length > 0; });
    }),
  };
}

function tsReadTable_(registry, source, requiredHeaders) {
  const id = registry.get(source.database);
  if (!id) throw new Error('Database not found in registry: ' + source.database);
  const sheet = SpreadsheetApp.openById(id).getSheetByName(source.sheet);
  if (!sheet) throw new Error('Missing sheet: ' + source.database +
    ' -> ' + source.sheet);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(tsText_);
  const seen = new Set();
  headers.forEach(function(header) {
    if (!header) return;
    if (seen.has(header)) throw new Error('Duplicate header: ' + header);
    seen.add(header);
  });
  requiredHeaders.forEach(function(header) {
    if (!seen.has(header)) throw new Error(source.sheet +
      ' is missing required header "' + header + '".');
  });
  return values.slice(1).filter(function(row) {
    return row.some(function(value) { return tsText_(value) !== ''; });
  }).map(function(row) {
    const record = {};
    headers.forEach(function(header, index) {
      if (header) record[header] = row[index];
    });
    return record;
  });
}

function tsFullName_(teacher) {
  return [teacher.NOM, teacher.COGNOM1, teacher.COGNOM2]
    .map(tsText_).filter(Boolean).join(' ');
}

function tsText_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function tsKey_(value) {
  return tsText_(value).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ').toUpperCase();
}

function tsSubjectColour_(subjectCode) {
  const code = tsText_(subjectCode).toUpperCase();
  if (['GUARDIA', 'GUARDIA_PATI'].indexOf(code) !== -1) {
    return { category: 'orange', background: '#fff4e5', accent: '#f59f00' };
  }
  if (['TUT', 'TUT_FAMILIES'].indexOf(code) !== -1) {
    return { category: 'pink', background: '#fff0f6', accent: '#f06595' };
  }
  if (code === '3R') {
    return { category: 'yellow', background: '#fff9db', accent: '#f2c94c' };
  }
  if (['RC_ESO', 'RC_FP_BAT', 'RDEP', 'RDIM1', 'RDIM2', 'RDIR',
    'REAP', 'REC', 'CARREC', 'FCT', 'TREC'].indexOf(code) !== -1) {
    return { category: 'blue', background: '#eef6ff', accent: '#4dabf7' };
  }
  return { category: 'green', background: '#edf7f6', accent: '#7fc7bd' };
}
```

The returned cell structure is:

```javascript
{
  items: [
    {
      subjectName: 'DIGITALITZACIÓ',
      subjectCode: 'DIG',
      groups: ['2A', '2B', '2C', '2D', '2E'],
      classroom: '2E',
      colour: {
        category: 'green',
        background: '#edf7f6',
        accent: '#7fc7bd'
      }
    }
  ]
}
```

For each item, render subjectName, then `Grup: ` plus `groups.join(', ')` if nonempty, then `Classe: ` plus classroom if nonblank. Apply its colour and centre the collection of items vertically in the cell.

## 7. Expected Error Handling

| Situation | Handling |
| --- | --- |
| Missing `db` property | Stop with a clear configuration error. |
| Missing registry entry or sheet | Report the exact database/sheet required. |
| Spreadsheet access denied | Report the access error; grant the executing account the required permission. |
| Missing required header | Report that header; do not guess a different schema. |
| Teacher name not found | Ask the caller to verify the full name or supply REDUIT. |
| Duplicate normalized name | Require an explicit teacher selection/code. |
| Teacher exists but has no effective rows | Return an empty timetable or a clear no-current-schedule message. |
| Invalid day or slot | Exclude that row from placement and expose it for diagnostics. |
| Blank group/classroom | Omit that display line. |
| Blank translated subject name | Use subject_code as the title. |
| Changed leave not reflected | Verify that the producer rebuilt the cache after the leave change. |

## 8. Integration Checks

Before using the document generator, check these cases:

1. A normal teacher resolves by full name to REDUIT and displays all expected slots.
2. A name with accents resolves when the input differs only in case, spacing, or accents.
3. A covered teacher's source fields remain present, but their substitute receives the effective timetable.
4. A shared lesson for several groups becomes one bubble with comma-separated groups.
5. Two lessons in different classrooms at the same day/slot remain separate bubbles.
6. Blank groups or rooms do not produce empty labels.
7. GUARDIA_PATI is orange, CARREC/FCT/TREC are blue, TUT_FAMILIES is pink, and 3R is yellow.
8. Printed output retains five weekdays, skips wholly empty slots, and keeps original slot numbers.
9. Duplicate teacher names produce an ambiguity error rather than a schedule for the wrong person.

## 9. Instructions For The Consuming Project

Use the `db` script property to open the registry spreadsheet. In `tables`, column A identifies a logical database and column B gives its spreadsheet ID. Configure database/sheet pairs explicitly.

Resolve a full teacher name through `Dades de professors -> Llista`, using NOM + COGNOM1 + COGNOM2, and obtain REDUIT. Read `Horaris -> schedule_cache` and filter by effective_teacher_code for the current teacher timetable. Preserve the distinction between source teachers and effective teachers. Do not repeat leave resolution or subject translation when the cache already supplies the result.

Render a Monday-Friday grid by numeric day and slot. Within each teacher cell, group by displayed subject name and classroom, joining distinct school groups with commas. Show the full subject name, optional `Grup: ...`, and optional `Classe: ...` vertically, using the documented colours based on subject_code. Centre the bubbles vertically. For printed output, use A4 landscape and include only slots with at least one nonempty weekday cell.
