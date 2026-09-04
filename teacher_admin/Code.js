const CONFIG = {
  tablesPropertyName: 'Tables',
  tablesSheetName: 'tables',
  dbRegistryName: 'Dades de professors',
  dbSheetName: 'Llista',
  leaveAbsenceSheetName: 'leave_absence',
  firstDataRow: 2,
  editableColumnCount: 16,
  nousColumn: 13,
  activeColumn: 14,
  baixaColumn: 15,
  substColumn: 16,
  cacheRebuildUrlPropertyName: 'cache_rebuild_url',
  cacheRebuildTokenPropertyName: 'cache_rebuild_token',
  defaultCacheRebuildUrl: 'https://script.google.com/macros/s/AKfycbyhSqCTkS27bDxsfILI64rlSMUTN5A7VbHGgpSf_G6efxrWfOuUKJULnN2rlMtHuWqwmA/exec',
  leaveAbsenceHeaders: ['row_id', 'teacher_code', 'substitute_code', 'start_date', 'end_date', 'comments'],
  cacheRebuildLogSheetName: 'cache_rebuild_log',
  cacheRebuildLogHeaders: [
    'timestamp',
    'event',
    'leave_event',
    'row_number',
    'teacher_code',
    'substitute_code',
    'leave_absence_row',
    'substitute_row_number',
    'endpoint_url',
    'has_token',
    'status_code',
    'ok',
    'response_text',
    'error',
  ],
};

const DB_COLUMNS = [
  { index: 1, key: 'esp', label: 'ESP', type: 'text' },
  { index: 2, key: 'dept', label: 'DEPT.', type: 'text' },
  { index: 3, key: 'nom', label: 'NOM', type: 'text' },
  { index: 4, key: 'cognom1', label: 'COGNOM 1', type: 'text' },
  { index: 5, key: 'cognom2', label: 'COGNOM 2', type: 'text' },
  { index: 6, key: 'reduit', label: 'REDUIT', type: 'text' },
  { index: 7, key: 'situacio', label: 'SITUACIO', type: 'select', options: [
    'FUNC. DEF',
    'FUNC. PERFIL',
    'FUNC. SNS PLAÇA',
    'INT',
    'INT. PERF',
    'LABORAL',
  ] },
  { index: 8, key: 'jornada', label: 'JORNADA', type: 'select', options: [
    'SENCERA',
    'MITJA',
    'REDUCCIÓ UN TERÇ',
  ] },
  { index: 9, key: 'dni', label: 'DNI', type: 'text' },
  { index: 10, key: 'telf', label: 'TELF', type: 'text' },
  { index: 11, key: 'xtec', label: 'XTEC', type: 'text' },
  { index: 12, key: 'correu', label: 'CORREU', type: 'text' },
  { index: 13, key: 'nous', label: 'NOUS', type: 'boolean' },
  { index: 14, key: 'actiu', label: 'ACTIU', type: 'boolean' },
  { index: 15, key: 'baixa', label: 'BAIXA?', type: 'boolean' },
  { index: 16, key: 'subst', label: 'SUBST?', type: 'boolean' },
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
  const lastColumn = Math.max(sheet.getLastColumn(), CONFIG.substColumn);

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
  const substitutes = rows
    .filter((row) => row.reduit && !row.baixa && !row.subst)
    .map((row) => ({ code: row.reduit, esp: row.esp, name: row.fullName }))
    .filter((row) => row.code && row.name)
    .sort((a, b) => a.name.localeCompare(b.name, 'ca'));

  return {
    rows,
    departments,
    substitutes,
  };
}

function updateBaixa() {
  throw new Error('La baixa nomes es pot gestionar amb el flux Donar de baixa / Donar d\'alta.');
}

function updateActiu(rowNumbers, value) {
  return updateBooleanColumn_(rowNumbers, CONFIG.activeColumn, value);
}

function startLeaveAbsence(rowNumber, leaveData) {
  const normalizedRowNumber = normalizeRowNumber_(rowNumber);
  const data = normalizeLeaveData_(leaveData, true);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let result;
  let notificationContext;

  try {
    const sheet = getDbSheet_();
    const teacher = sheet
      .getRange(normalizedRowNumber, 1, 1, CONFIG.editableColumnCount)
      .getValues()[0];
    const teacherReduit = toDisplayString_(teacher[5]);
    const isAlreadyOnLeave = isTruthyValue_(teacher[CONFIG.baixaColumn - 1]);

    if (!teacherReduit) {
      throw new Error('El professor seleccionat no te codi REDUIT.');
    }

    if (isAlreadyOnLeave) {
      throw new Error('Aquest professor ja esta de baixa.');
    }

    if (data.substituteCode === teacherReduit) {
      throw new Error('El substitut no pot ser el mateix professor.');
    }

    const substituteRowNumber = getAvailableSubstituteRowNumber_(sheet, data.substituteCode);
    if (!substituteRowNumber) {
      throw new Error('El substitut seleccionat no es valid.');
    }

    sheet.getRange(normalizedRowNumber, CONFIG.baixaColumn).setValue(true);
    sheet.getRange(substituteRowNumber, CONFIG.activeColumn).setValue(true);
    sheet.getRange(substituteRowNumber, CONFIG.substColumn).setValue(true);
    const leaveSheet = getLeaveAbsenceSheet_();
    leaveSheet.appendRow([
      normalizedRowNumber,
      teacherReduit,
      data.substituteCode,
      data.date,
      '',
      data.comments,
    ]);
    SpreadsheetApp.flush();

    result = { updatedRows: 1 };
    notificationContext = {
      event: 'startLeaveAbsence',
      rowNumber: normalizedRowNumber,
      teacherCode: teacherReduit,
      substituteCode: data.substituteCode,
      leaveAbsenceRow: leaveSheet.getLastRow(),
    };
  } finally {
    lock.releaseLock();
  }

  notifyScheduleCacheRebuild_(notificationContext);
  return result;
}

function endLeaveAbsence(rowNumber, leaveData) {
  const normalizedRowNumber = normalizeRowNumber_(rowNumber);
  const data = normalizeLeaveData_(leaveData, false);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let result;
  let notificationContext;

  try {
    const sheet = getDbSheet_();
    const leaveSheet = getLeaveAbsenceSheet_();
    const lastRow = leaveSheet.getLastRow();
    let updatedLeaveRow = null;
    let substituteCode = '';

    if (lastRow >= 2) {
      const values = leaveSheet.getRange(2, 1, lastRow - 1, CONFIG.leaveAbsenceHeaders.length).getValues();
      for (let index = values.length - 1; index >= 0; index -= 1) {
        const row = values[index];
        const rowId = Number(row[0]);
        const endDate = row[4];
        if (rowId === normalizedRowNumber && !endDate) {
          updatedLeaveRow = index + 2;
          substituteCode = toDisplayString_(row[2]);
          break;
        }
      }
    }

    if (!updatedLeaveRow) {
      throw new Error('No s’ha trobat cap baixa oberta per aquest professor.');
    }

    const substituteRowNumber = substituteCode
      ? getTeacherRowNumberByReduit_(sheet, substituteCode)
      : null;

    leaveSheet.getRange(updatedLeaveRow, 5).setValue(data.date);
    sheet.getRange(normalizedRowNumber, CONFIG.baixaColumn).setValue(false);
    if (substituteRowNumber) {
      sheet.getRange(substituteRowNumber, CONFIG.substColumn).setValue(false);
      sheet.getRange(substituteRowNumber, CONFIG.activeColumn).setValue(false);
    }
    SpreadsheetApp.flush();

    result = { updatedRows: 1 };
    notificationContext = {
      event: 'endLeaveAbsence',
      rowNumber: normalizedRowNumber,
      substituteCode,
      leaveAbsenceRow: updatedLeaveRow,
      substituteRowNumber,
    };
  } finally {
    lock.releaseLock();
  }

  notifyScheduleCacheRebuild_(notificationContext);
  return result;
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

  const sheet = getDbSheet_();
  const existingValues = sheet
    .getRange(normalizedRowNumber, 1, 1, CONFIG.editableColumnCount)
    .getValues()[0];
  const values = DB_COLUMNS.map((column) => {
    if (column.key === 'baixa') {
      return existingValues[column.index - 1];
    }
    const value = fields[column.key];
    if (column.type === 'boolean') {
      return value === true;
    }
    if (column.type === 'select') {
      return normalizeAllowedValue_(value, column.options, column.label);
    }
    return toDisplayString_(value);
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
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

function getLeaveAbsenceSheet_() {
  const spreadsheet = getDbSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(CONFIG.leaveAbsenceSheetName);

  if (!sheet) {
    throw new Error(`No s'ha trobat el full "${CONFIG.leaveAbsenceSheetName}" a DB.`);
  }

  ensureLeaveAbsenceHeader_(sheet);
  return sheet;
}

function ensureLeaveAbsenceHeader_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(CONFIG.leaveAbsenceHeaders);
    return;
  }

  const header = sheet.getRange(1, 1, 1, CONFIG.leaveAbsenceHeaders.length).getValues()[0];
  const hasHeader = CONFIG.leaveAbsenceHeaders.every((name, index) => header[index] === name);
  if (!hasHeader) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, CONFIG.leaveAbsenceHeaders.length).setValues([CONFIG.leaveAbsenceHeaders]);
  }
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
  const match = values.find((row) => row[0] === CONFIG.dbRegistryName);

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
  const reduit = toDisplayString_(row[5]);
  const nameParts = [name, firstSurname, secondSurname].filter(Boolean);
  const fullName = nameParts.join(' ');
  const fullNameSort = [firstSurname, secondSurname, name].filter(Boolean).join(' ');
  const situacio = toDisplayString_(row[6]);
  const jornada = toDisplayString_(row[7]);
  const dni = toDisplayString_(row[8]);
  const telf = toDisplayString_(row[9]);
  const xtec = toDisplayString_(row[10]);
  const correu = toDisplayString_(row[11]);
  const nou = isTruthyValue_(row[12]);
  const active = isTruthyValue_(row[13]);
  const baixa = isTruthyValue_(row[14]);
  const subst = isTruthyValue_(row[15]);
  const noActiu = !active;

  return {
    rowNumber,
    esp,
    dept,
    nom: name,
    cognom1: firstSurname,
    cognom2: secondSurname,
    reduit,
    fullName,
    situacio,
    jornada,
    dni,
    telf,
    correu,
    xtec,
    baixa,
    nou,
    active,
    subst,
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

function getAvailableSubstituteRowNumber_(sheet, substituteCode) {
  const values = getProfessorLookupValues_(sheet);

  for (let index = 0; index < values.length; index += 1) {
    const row = values[index];
    const reduit = toDisplayString_(row[5]);
    const baixa = isTruthyValue_(row[CONFIG.baixaColumn - 1]);
    const subst = isTruthyValue_(row[CONFIG.substColumn - 1]);
    if (reduit === substituteCode && !baixa && !subst) {
      return CONFIG.firstDataRow + index;
    }
  }

  return null;
}

function getTeacherRowNumberByReduit_(sheet, reduitCode) {
  const values = getProfessorLookupValues_(sheet);

  for (let index = 0; index < values.length; index += 1) {
    const row = values[index];
    const reduit = toDisplayString_(row[5]);
    if (reduit === reduitCode) {
      return CONFIG.firstDataRow + index;
    }
  }

  return null;
}

function getProfessorLookupValues_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.firstDataRow) {
    return [];
  }

  return sheet
    .getRange(CONFIG.firstDataRow, 1, lastRow - CONFIG.firstDataRow + 1, CONFIG.substColumn)
    .getValues();
}

function notifyScheduleCacheRebuild_(context) {
  const properties = PropertiesService.getScriptProperties();
  const endpointUrl = toDisplayString_(
    properties.getProperty(CONFIG.cacheRebuildUrlPropertyName)
      || CONFIG.defaultCacheRebuildUrl
  );
  const token = toDisplayString_(
    properties.getProperty(CONFIG.cacheRebuildTokenPropertyName)
  );
  const logContext = context || {};
  const requestLogUrl = endpointUrl;

  console.log(JSON.stringify({
    event: 'scheduleCacheRebuild.notify.start',
    context: logContext,
    endpointUrl: requestLogUrl,
    hasToken: Boolean(token),
    tokenPropertyName: CONFIG.cacheRebuildTokenPropertyName,
    urlPropertyName: CONFIG.cacheRebuildUrlPropertyName,
  }));
  appendScheduleCacheRebuildLog_('scheduleCacheRebuild.notify.start', {
    context: logContext,
    endpointUrl: requestLogUrl,
    hasToken: Boolean(token),
  });

  if (!endpointUrl) {
    console.error(JSON.stringify({
      event: 'scheduleCacheRebuild.notify.configError',
      context: logContext,
      error: 'missing endpoint URL',
    }));
    appendScheduleCacheRebuildLog_('scheduleCacheRebuild.notify.configError', {
      context: logContext,
      endpointUrl: requestLogUrl,
      hasToken: Boolean(token),
      error: 'missing endpoint URL',
    });
    throw new Error('No hi ha endpoint configurat per reconstruir schedule_cache.');
  }

  if (!token) {
    console.error(JSON.stringify({
      event: 'scheduleCacheRebuild.notify.configError',
      context: logContext,
      error: `missing ${CONFIG.cacheRebuildTokenPropertyName}`,
    }));
    appendScheduleCacheRebuildLog_('scheduleCacheRebuild.notify.configError', {
      context: logContext,
      endpointUrl: requestLogUrl,
      hasToken: false,
      error: `missing ${CONFIG.cacheRebuildTokenPropertyName}`,
    });
    throw new Error(`Falta la propietat de script "${CONFIG.cacheRebuildTokenPropertyName}".`);
  }

  let response;
  try {
    response = UrlFetchApp.fetch(endpointUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        action: 'rebuildScheduleCache',
        token,
      }),
      muteHttpExceptions: true,
    });
  } catch (error) {
    console.error(JSON.stringify({
      event: 'scheduleCacheRebuild.notify.fetchError',
      context: logContext,
      endpointUrl: requestLogUrl,
      error: error && error.message ? error.message : String(error),
    }));
    appendScheduleCacheRebuildLog_('scheduleCacheRebuild.notify.fetchError', {
      context: logContext,
      endpointUrl: requestLogUrl,
      hasToken: true,
      error: error && error.message ? error.message : String(error),
    });
    throw error;
  }

  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();
  console.log(JSON.stringify({
    event: 'scheduleCacheRebuild.notify.response',
    context: logContext,
    endpointUrl: requestLogUrl,
    statusCode,
    responseText: truncateForLog_(responseText, 1000),
  }));
  appendScheduleCacheRebuildLog_('scheduleCacheRebuild.notify.response', {
    context: logContext,
    endpointUrl: requestLogUrl,
    hasToken: true,
    statusCode,
    responseText,
  });

  let result;

  try {
    result = JSON.parse(responseText);
  } catch (error) {
    console.error(JSON.stringify({
      event: 'scheduleCacheRebuild.notify.parseError',
      context: logContext,
      statusCode,
      responseText: truncateForLog_(responseText, 1000),
      error: error && error.message ? error.message : String(error),
    }));
    appendScheduleCacheRebuildLog_('scheduleCacheRebuild.notify.parseError', {
      context: logContext,
      endpointUrl: requestLogUrl,
      hasToken: true,
      statusCode,
      responseText,
      error: error && error.message ? error.message : String(error),
    });
    throw new Error(
      `No s'ha pogut interpretar la resposta de rebuildScheduleCache (${statusCode}).`
    );
  }

  if (statusCode < 200 || statusCode >= 300 || !result.ok) {
    console.error(JSON.stringify({
      event: 'scheduleCacheRebuild.notify.rebuildError',
      context: logContext,
      statusCode,
      result,
    }));
    appendScheduleCacheRebuildLog_('scheduleCacheRebuild.notify.rebuildError', {
      context: logContext,
      endpointUrl: requestLogUrl,
      hasToken: true,
      statusCode,
      ok: result && result.ok,
      responseText,
      error: result && result.error ? result.error : '',
    });
    throw new Error(
      result && result.error
        ? `Error reconstruint schedule_cache: ${result.error}`
        : `Error reconstruint schedule_cache (${statusCode}).`
    );
  }

  console.log(JSON.stringify({
    event: 'scheduleCacheRebuild.notify.success',
    context: logContext,
    statusCode,
    result,
  }));
  appendScheduleCacheRebuildLog_('scheduleCacheRebuild.notify.success', {
    context: logContext,
    endpointUrl: requestLogUrl,
    hasToken: true,
    statusCode,
    ok: true,
    responseText,
  });

  return result;
}

function appendScheduleCacheRebuildLog_(eventName, details) {
  try {
    const detailData = details || {};
    const context = detailData.context || {};
    const sheet = getCacheRebuildLogSheet_();
    sheet.appendRow([
      new Date(),
      eventName,
      toDisplayString_(context.event),
      context.rowNumber || '',
      toDisplayString_(context.teacherCode),
      toDisplayString_(context.substituteCode),
      context.leaveAbsenceRow || '',
      context.substituteRowNumber || '',
      toDisplayString_(detailData.endpointUrl),
      detailData.hasToken === true,
      detailData.statusCode || '',
      detailData.ok === undefined ? '' : detailData.ok === true,
      truncateForLog_(detailData.responseText || '', 1000),
      toDisplayString_(detailData.error),
    ]);
  } catch (error) {
    console.error(JSON.stringify({
      event: 'scheduleCacheRebuild.log.writeError',
      originalEvent: eventName,
      error: error && error.message ? error.message : String(error),
    }));
  }
}

function getCacheRebuildLogSheet_() {
  const spreadsheet = getDbSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(CONFIG.cacheRebuildLogSheetName)
    || spreadsheet.insertSheet(CONFIG.cacheRebuildLogSheetName);

  ensureHeader_(sheet, CONFIG.cacheRebuildLogHeaders);
  return sheet;
}

function ensureHeader_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    return;
  }

  const header = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const hasHeader = headers.every((name, index) => header[index] === name);
  if (!hasHeader) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function truncateForLog_(value, maxLength) {
  const text = toDisplayString_(value);
  if (text.length <= maxLength) {
    return text;
  }

  return `${text.slice(0, maxLength)}...`;
}

function normalizeAllowedValue_(value, options, label) {
  const text = toDisplayString_(value);
  if (!text) {
    return '';
  }

  if (!options.includes(text)) {
    throw new Error(`El valor de "${label}" no es valid.`);
  }

  return text;
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

function normalizeLeaveData_(leaveData, requiresSubstitute) {
  if (!leaveData || typeof leaveData !== 'object') {
    throw new Error('Les dades de la baixa no son valides.');
  }

  const date = parseIsoDate_(leaveData.date);
  const substituteCode = toDisplayString_(leaveData.substituteCode);
  const comments = toDisplayString_(leaveData.comments);

  if (requiresSubstitute && !substituteCode) {
    throw new Error('Cal seleccionar un substitut.');
  }

  return { date, substituteCode, comments };
}

function parseIsoDate_(value) {
  const text = toDisplayString_(value);
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);

  if (!match) {
    throw new Error('La data no es valida.');
  }

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(year, month - 1, day);

  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    throw new Error('La data no es valida.');
  }

  return date;
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
