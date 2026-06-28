const CONFIG = {
  tablesPropertyName: 'Tables',
  tablesSheetName: 'tables',
  dbRegistryName: 'Dades de professors',
  dbSheetName: 'Llista',
  firstDataRow: 2,
  editableColumnCount: 14,
  baixaColumn: 12,
  noActiusColumn: 14,
};

const DB_COLUMNS = [
  { index: 1, key: 'esp', label: 'ESP', type: 'text' },
  { index: 2, key: 'dept', label: 'DEPT.', type: 'text' },
  { index: 3, key: 'nom', label: 'NOM', type: 'text' },
  { index: 4, key: 'cognom1', label: 'COGNOM 1', type: 'text' },
  { index: 5, key: 'cognom2', label: 'COGNOM 2', type: 'text' },
  { index: 6, key: 'colF', label: 'COLUMNA F', type: 'text' },
  { index: 7, key: 'situacio', label: 'SITUACIO', type: 'text' },
  { index: 8, key: 'dni', label: 'DNI', type: 'text' },
  { index: 9, key: 'telf', label: 'TELF', type: 'text' },
  { index: 10, key: 'xtec', label: 'XTEC', type: 'text' },
  { index: 11, key: 'correu', label: 'CORREU', type: 'text' },
  { index: 12, key: 'baixa', label: 'BAIXA?', type: 'boolean' },
  { index: 13, key: 'colM', label: 'COLUMNA M', type: 'text' },
  { index: 14, key: 'actiu', label: 'ACTIU', type: 'boolean' },
];

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dades de professors')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getProfessorsData() {
  const sheet = getDbSheet_();
  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), CONFIG.noActiusColumn);

  if (lastRow < CONFIG.firstDataRow) {
    return {
      rows: [],
      departments: [],
    };
  }

  const rowCount = lastRow - CONFIG.firstDataRow + 1;
  const values = sheet
    .getRange(CONFIG.firstDataRow, 1, rowCount, lastColumn)
    .getValues();

  const rows = values
    .map((row, index) => mapProfessorRow_(row, CONFIG.firstDataRow + index))
    .filter((row) => row.hasData);

  const departments = [...new Set(rows.map((row) => row.dept).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b, 'ca'));

  return {
    rows,
    departments,
  };
}

function updateBaixa(rowNumbers, value) {
  return updateBooleanColumn_(rowNumbers, CONFIG.baixaColumn, value);
}

function updateActiu(rowNumbers, value) {
  return updateBooleanColumn_(rowNumbers, CONFIG.noActiusColumn, value);
}

function getTeacherDetails(rowNumber) {
  const normalizedRowNumber = normalizeRowNumber_(rowNumber);
  const sheet = getDbSheet_();
  const values = sheet
    .getRange(normalizedRowNumber, 1, 1, CONFIG.editableColumnCount)
    .getValues()[0];
  const displayValues = sheet
    .getRange(normalizedRowNumber, 1, 1, CONFIG.editableColumnCount)
    .getDisplayValues()[0];

  return {
    rowNumber: normalizedRowNumber,
    title: buildFullName_(values),
    fields: DB_COLUMNS.map((column) => {
      const index = column.index - 1;
      const rawValue = values[index];
      const value = column.type === 'boolean'
        ? isTruthyValue_(rawValue)
        : toDisplayString_(displayValues[index]);

      return {
        key: column.key,
        label: column.label,
        type: column.type,
        value,
      };
    }),
  };
}

function saveTeacherDetails(rowNumber, fields) {
  const normalizedRowNumber = normalizeRowNumber_(rowNumber);

  if (!fields || typeof fields !== 'object') {
    throw new Error('Les dades del formulari no son valides.');
  }

  const values = DB_COLUMNS.map((column) => {
    const value = fields[column.key];
    return column.type === 'boolean' ? value === true : toDisplayString_(value);
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = getDbSheet_();
    sheet.getRange(normalizedRowNumber, 1, 1, CONFIG.editableColumnCount).setValues([values]);
    SpreadsheetApp.flush();

    return getTeacherDetails(normalizedRowNumber);
  } finally {
    lock.releaseLock();
  }
}

function exportTeachers(rowNumbers) {
  const uniqueRows = normalizeRowNumbers_(rowNumbers);
  const sheet = getDbSheet_();
  const header = sheet
    .getRange(1, 1, 1, CONFIG.editableColumnCount)
    .getDisplayValues()[0];
  const rows = uniqueRows.map((rowNumber) => {
    return sheet
      .getRange(rowNumber, 1, 1, CONFIG.editableColumnCount)
      .getDisplayValues()[0];
  });

  const fileName = `professors-export-${Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyyMMdd-HHmmss'
  )}.csv`;

  return {
    fileName,
    mimeType: 'text/csv;charset=utf-8',
    content: toCsv_([header, ...rows]),
    rowCount: rows.length,
  };
}

function getDbSheet_() {
  const spreadsheet = getDbSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(CONFIG.dbSheetName);

  if (!sheet) {
    throw new Error(`No s'ha trobat el full "${CONFIG.dbSheetName}" a DB.`);
  }

  return sheet;
}

function getDbSpreadsheet_() {
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
  const match = values.find((row) => String(row[0]).trim() === CONFIG.dbRegistryName);

  if (!match) {
    throw new Error(
      `No s'ha trobat "${CONFIG.dbRegistryName}" al full "${CONFIG.tablesSheetName}".`
    );
  }

  const dbSpreadsheetId = String(match[1] || '').trim();

  if (!dbSpreadsheetId) {
    throw new Error(`La fila "${CONFIG.dbRegistryName}" no te ID a la columna B.`);
  }

  return SpreadsheetApp.openById(dbSpreadsheetId);
}

function mapProfessorRow_(row, rowNumber) {
  const esp = toDisplayString_(row[0]);
  const dept = toDisplayString_(row[1]);
  const name = toDisplayString_(row[2]);
  const firstSurname = toDisplayString_(row[3]);
  const secondSurname = toDisplayString_(row[4]);
  const nameParts = [name, firstSurname, secondSurname].filter(Boolean);
  const fullName = nameParts.join(' ');
  const fullNameSort = [firstSurname, secondSurname, name].filter(Boolean).join(' ');
  const situacio = toDisplayString_(row[6]);
  const dni = toDisplayString_(row[7]);
  const telf = toDisplayString_(row[8]);
  const xtec = toDisplayString_(row[9]);
  const correu = toDisplayString_(row[10]);
  const baixa = isTruthyValue_(row[11]);
  const noActiu = !isTruthyValue_(row[13]);

  return {
    rowNumber,
    esp,
    dept,
    fullName,
    situacio,
    dni,
    telf,
    correu,
    xtec,
    baixa,
    noActiu,
    fullNameSort,
    hasData: [esp, dept, fullName, situacio, dni, telf, correu, xtec].some(Boolean),
  };
}

function updateBooleanColumn_(rowNumbers, column, value) {
  const booleanValue = value === true;
  const uniqueRows = normalizeRowNumbers_(rowNumbers);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = getDbSheet_();
    uniqueRows.forEach((rowNumber) => {
      sheet.getRange(rowNumber, column).setValue(booleanValue);
    });

    SpreadsheetApp.flush();

    return {
      updatedRows: uniqueRows.length,
    };
  } finally {
    lock.releaseLock();
  }
}

function normalizeRowNumbers_(rowNumbers) {
  if (!Array.isArray(rowNumbers) || rowNumbers.length === 0) {
    throw new Error('Cal seleccionar almenys una fila.');
  }

  const uniqueRows = [...new Set(rowNumbers.map(Number))]
    .filter((rowNumber) => Number.isInteger(rowNumber) && rowNumber >= CONFIG.firstDataRow)
    .sort((a, b) => a - b);

  if (uniqueRows.length === 0) {
    throw new Error('Les files seleccionades no son valides.');
  }

  return uniqueRows;
}

function normalizeRowNumber_(rowNumber) {
  const normalizedRowNumber = Number(rowNumber);

  if (!Number.isInteger(normalizedRowNumber) || normalizedRowNumber < CONFIG.firstDataRow) {
    throw new Error('La fila seleccionada no es valida.');
  }

  return normalizedRowNumber;
}

function toDisplayString_(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value).trim();
}

function isTruthyValue_(value) {
  if (value === true) {
    return true;
  }

  if (typeof value === 'string') {
    return value.trim().toLowerCase() === 'true';
  }

  return false;
}

function buildFullName_(row) {
  return [row[2], row[3], row[4]]
    .map(toDisplayString_)
    .filter(Boolean)
    .join(' ');
}

function toCsv_(rows) {
  return rows.map((row) => row.map(toCsvCell_).join(',')).join('\r\n');
}

function toCsvCell_(value) {
  const text = toDisplayString_(value);
  if (/[",\r\n]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }

  return text;
}
