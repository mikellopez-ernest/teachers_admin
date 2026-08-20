const CONFIG = {
  registryPropertyName: 'Tables',
  registrySheetName: 'tables',
  teacherDbName: 'Dades de professors',
  teacherDbSheetName: 'Llista',
  xlsxMimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

const TEACHER_COLUMNS = {
  dept: 2,
  nom: 3,
  cognom1: 4,
  cognom2: 5,
  correuInstit: 12,
  actiu: 14,
};

function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Llistat professorat')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function grantRequiredPermissions() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.getProperty(CONFIG.registryPropertyName);

  const db = getTeacherDbSpreadsheet_();
  const sheet = getTeacherDbSheet_();
  sheet.getRange(1, 1).getValue();

  const tempSpreadsheet = SpreadsheetApp.create('teacher_list_permission_check');
  const tempFile = DriveApp.getFileById(tempSpreadsheet.getId());
  try {
    const response = UrlFetchApp.fetch(getXlsxExportUrl_(tempSpreadsheet.getId()), {
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
      },
      muteHttpExceptions: true,
    });

    return {
      ok: true,
      dbName: db.getName(),
      teacherSheetName: sheet.getName(),
      exportStatus: response.getResponseCode(),
      message: 'Permisos concedits correctament.',
    };
  } finally {
    tempFile.setTrashed(true);
  }
}

function getTeacherListData() {
  const rows = getActiveTeacherRows_({});
  const departments = [...new Set(rows.map((row) => row.dept).filter(Boolean))]
    .sort((a, b) => compareText_(a, b));

  return {
    rows,
    departments,
  };
}

function createTeacherListXlsx(filters) {
  const normalizedFilters = normalizeFilters_(filters || {});
  const rows = getActiveTeacherRows_(normalizedFilters);
  const model = buildTeacherListXlsxModel_(rows, normalizedFilters);
  const fileName = buildTeacherListFileName_();
  const tempSpreadsheet = SpreadsheetApp.create(fileName.replace(/\.xlsx$/i, ''));
  const tempFile = DriveApp.getFileById(tempSpreadsheet.getId());

  try {
    const sheet = tempSpreadsheet.getSheets()[0];
    sheet.setName('Llistat');
    applyTeacherListSheet_(sheet, model);
    SpreadsheetApp.flush();

    const response = UrlFetchApp.fetch(getXlsxExportUrl_(tempSpreadsheet.getId()), {
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
      },
      muteHttpExceptions: true,
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`No s'ha pogut exportar el llistat XLSX (${response.getResponseCode()}).`);
    }

    return {
      fileName,
      mimeType: CONFIG.xlsxMimeType,
      base64: Utilities.base64Encode(response.getBlob().getBytes()),
      message: `S'ha exportat el llistat amb ${rows.length} professors.`,
    };
  } finally {
    tempFile.setTrashed(true);
  }
}

function getTeacherDbSpreadsheet_() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const registrySpreadsheetId = scriptProperties.getProperty(CONFIG.registryPropertyName);

  if (!registrySpreadsheetId) {
    throw new Error(`Falta la propietat de script "${CONFIG.registryPropertyName}".`);
  }

  const registrySpreadsheet = SpreadsheetApp.openById(registrySpreadsheetId);
  const registrySheet = registrySpreadsheet.getSheetByName(CONFIG.registrySheetName);

  if (!registrySheet) {
    throw new Error(`No s'ha trobat el full "${CONFIG.registrySheetName}" al registre.`);
  }

  const values = registrySheet.getDataRange().getValues();
  const match = values.find((row) => row[0] === CONFIG.teacherDbName);

  if (!match || !match[1]) {
    throw new Error(`No s'ha trobat "${CONFIG.teacherDbName}" al registre de taules.`);
  }

  return SpreadsheetApp.openById(String(match[1]).trim());
}

function getTeacherDbSheet_() {
  const spreadsheet = getTeacherDbSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(CONFIG.teacherDbSheetName);

  if (!sheet) {
    throw new Error(`No s'ha trobat el full "${CONFIG.teacherDbSheetName}" a DB.`);
  }

  return sheet;
}

function getActiveTeacherRows_(filters) {
  const normalizedFilters = normalizeFilters_(filters || {});
  const sheet = getTeacherDbSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, TEACHER_COLUMNS.actiu).getValues();
  const rows = values
    .map((row, index) => ({
      rowNumber: index + 2,
      dept: valueToString_(row[TEACHER_COLUMNS.dept - 1]),
      nom: valueToString_(row[TEACHER_COLUMNS.nom - 1]),
      cognom1: valueToString_(row[TEACHER_COLUMNS.cognom1 - 1]),
      cognom2: valueToString_(row[TEACHER_COLUMNS.cognom2 - 1]),
      correuInstit: valueToString_(row[TEACHER_COLUMNS.correuInstit - 1]),
      actiu: isTrue_(row[TEACHER_COLUMNS.actiu - 1]),
    }))
    .filter((row) => row.actiu)
    .filter((row) => !normalizedFilters.department || row.dept === normalizedFilters.department)
    .filter((row) => {
      if (!normalizedFilters.search) return true;
      return normalizeText_([row.nom, row.cognom1, row.cognom2].join(' ')).includes(normalizedFilters.search);
    })
    .sort((a, b) => compareTeacherRows_(a, b));

  return rows;
}

function buildTeacherListXlsxModel_(rows, filters) {
  return {
    title: 'Llistat de professorat',
    subtitle: buildSubtitle_(rows, filters),
    headers: ['DEPT.', 'NOM', 'COGNOM1', 'COGNOM2', 'CORREU INSTIT'],
    rows: rows.map((row) => [row.dept, row.nom, row.cognom1, row.cognom2, row.correuInstit]),
    columnWidths: [130, 150, 170, 170, 260],
    headerBackground: '#102027',
    headerColor: '#ffffff',
  };
}

function applyTeacherListSheet_(sheet, model) {
  sheet.clear();
  sheet.getRange(1, 1, 1, model.headers.length)
    .merge()
    .setValue(model.title)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#102027')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 34);

  sheet.getRange(2, 1, 1, model.headers.length)
    .merge()
    .setValue(model.subtitle)
    .setFontColor('#64748b')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(2, 26);

  const headerRange = sheet.getRange(4, 1, 1, model.headers.length);
  headerRange
    .setValues([model.headers])
    .setBackground(model.headerBackground)
    .setFontColor(model.headerColor)
    .setFontWeight('bold')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(4, 28);

  if (model.rows.length > 0) {
    sheet.getRange(5, 1, model.rows.length, model.headers.length)
      .setValues(model.rows)
      .setVerticalAlignment('middle')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    sheet.setRowHeights(5, model.rows.length, 24);
  }

  model.columnWidths.forEach((width, index) => sheet.setColumnWidth(index + 1, width));
  sheet.setFrozenRows(4);

  const fullRange = sheet.getRange(4, 1, Math.max(model.rows.length + 1, 1), model.headers.length);
  fullRange.setBorder(true, true, true, true, true, true, '#d9e0e8', SpreadsheetApp.BorderStyle.SOLID);
  sheet.autoResizeRows(1, Math.max(model.rows.length + 4, 4));
}

function buildSubtitle_(rows, filters) {
  const parts = [`Professors actius: ${rows.length}`];
  if (filters.department) parts.push(`Departament: ${filters.department}`);
  if (filters.search) parts.push(`Filtre: ${filters.rawSearch}`);
  parts.push(`Generat: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')}`);
  return parts.join(' · ');
}

function buildTeacherListFileName_() {
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmm');
  return `llistat-professorat-${stamp}.xlsx`;
}

function getXlsxExportUrl_(fileId) {
  return `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}/export?mimeType=${encodeURIComponent(CONFIG.xlsxMimeType)}`;
}

function normalizeFilters_(filters) {
  const rawSearch = valueToString_(filters.search || filters.name || '');
  return {
    rawSearch,
    search: normalizeText_(rawSearch),
    department: valueToString_(filters.department || filters.dept || ''),
  };
}

function compareTeacherRows_(a, b) {
  return compareText_(a.cognom1, b.cognom1)
    || compareText_(a.cognom2, b.cognom2)
    || compareText_(a.nom, b.nom)
    || compareText_(a.dept, b.dept);
}

function compareText_(a, b) {
  return valueToString_(a).localeCompare(valueToString_(b), 'ca', {
    numeric: true,
    sensitivity: 'base',
  });
}

function normalizeText_(value) {
  return valueToString_(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLocaleLowerCase('ca');
}

function valueToString_(value) {
  return String(value == null ? '' : value).trim();
}

function isTrue_(value) {
  return value === true || String(value).trim().toUpperCase() === 'TRUE';
}
