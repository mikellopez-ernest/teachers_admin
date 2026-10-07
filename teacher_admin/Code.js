const CONFIG = {
  tablesPropertyName: 'Tables',
  scheduleRegistryPropertyName: 'db',
  accessGrantedPropertyName: 'access_granted',
  tablesSheetName: 'tables',
  dbRegistryName: 'Dades de professors',
  workloadRegistryName: 'Càrrega lectiva',
  scheduleRegistryName: 'Horaris',
  dbSheetName: 'Llista',
  workloadProfessorsSheetName: 'professors',
  workloadCarrecsSheetName: 'carrecs',
  scheduleCacheSheetName: 'schedule_cache',
  leaveAbsenceSheetName: 'leave_absence',
  firstDataRow: 2,
  editableColumnCount: 16,
  nousColumn: 13,
  activeColumn: 14,
  baixaColumn: 15,
  substColumn: 16,
  leaveNotificationRecipient: 'claustre@iernestlluch.cat',
  leaveNotificationSubject: 'Nova incorporació',
  cacheRebuildUrlPropertyName: 'cache_rebuild_url',
  cacheRebuildTokenPropertyName: 'cache_rebuild_token',
  officialScheduleTeachersDocIdPropertyName: 'official_schedule_teachers_doc_id',
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

const WORKLOAD_PROFESSORS_COLUMNS = {
  correuInstit: 12,
  teacherKey: 17,
};

const CARRECS_COLUMNS = {
  carrec: 1,
  asignado: 4,
  isCarrec: 6,
};

const DOCX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const DOCX_BLANK_FIRST_FOOTER_PATH = 'word/footer-teacher-admin-blank.xml';
const DOCX_BLANK_FIRST_FOOTER_REL_ID = 'rIdTeacherAdminBlankFirstFooter';

const ANNUAL_DATA_TOTAL_TAGS = [
  'HORES_ESO', 'HORES_PFI', 'HORES_BAT', 'HORES_CICLES', 'HORES_FCT',
  'TUT_ESO', 'TUT_BAT', 'TUT_FP', 'CARREC_DIRECTIU', 'CARREC',
  'REUNIONS', 'GUARDIES',
];

const SCHEDULE_GROUPS = {
  ESO: [
    '1A', '1B', '1C', '1D', '1E', '1F',
    '2A', '2B', '2C', '2D', '2E', '2F',
    '3A', '3B', '3C', '3D', '3E', '3F',
    '4A', '4B', '4C', '4D', '4E', '4F',
  ],
  PFI: ['PFI'],
  BAT: ['1BATA', '1BATB', '1BATC', '2BATA', '2BATB'],
  FP: ['PER1A', 'PER1B', 'PER2A', 'PER2B', 'SMX1', 'SMX2', 'AACC1', 'AACC2'],
};

const SCHEDULE_DIRECTIVE_TITLES = ['C.DIR', 'RDIR2', 'RDIR3', 'REUNIÓ DIRECCIÓ'];

const SCHEDULE_EXACT_TITLE_TAGS = {
  '3r': 'CARREC',
  coordinacio: 'CARREC',
  equipconvivencia: 'CARREC',
  reduccio55: 'CARREC',
  ruec: 'CARREC',
  trec: 'CARREC',
  gpatisiei: 'GUARDIES',
  paei: 'HORES_ESO',
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
  const access = getAccessDecision_();
  if (!access.allowed) return createAccessDeniedOutput_(access);

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dades de professors')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function grantRequiredPermissions() {
  const properties = PropertiesService.getScriptProperties();
  properties.getProperty(CONFIG.tablesPropertyName);
  properties.getProperty(CONFIG.scheduleRegistryPropertyName);
  properties.getProperty(CONFIG.accessGrantedPropertyName);
  properties.getProperty(CONFIG.cacheRebuildUrlPropertyName);
  properties.getProperty(CONFIG.cacheRebuildTokenPropertyName);
  const officialScheduleTemplateId = toDisplayString_(
    properties.getProperty(CONFIG.officialScheduleTeachersDocIdPropertyName)
  );

  const activeUserEmail = Session.getActiveUser().getEmail();
  const dbSheet = getDbSheet_();
  const leaveAbsenceSheet = getLeaveAbsenceSheet_();
  const workloadProfessorsSheet = getWorkloadProfessorsSheet_();
  const workloadCarrecsSheet = getWorkloadCarrecsSheet_();
  const scheduleCacheRows = getScheduleCacheRows_();

  dbSheet.getRange(1, 1).getValue();
  leaveAbsenceSheet.getRange(1, 1).getValue();
  workloadProfessorsSheet.getRange(1, 1).getValue();
  workloadCarrecsSheet.getRange(1, 1).getValue();

  UrlFetchApp.getRequest(CONFIG.defaultCacheRebuildUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ action: 'permissionCheck' }),
  });

  const mailRemainingDailyQuota = MailApp.getRemainingDailyQuota();
  const annualDataPermissions = officialScheduleTemplateId
    ? verifyAnnualDataDrivePermissions_(officialScheduleTemplateId)
    : { templateName: '' };

  return {
    ok: true,
    message: 'Permisos concedits correctament.',
    activeUserEmail,
    dbSheetName: dbSheet.getName(),
    leaveAbsenceSheetName: leaveAbsenceSheet.getName(),
    workloadProfessorsSheetName: workloadProfessorsSheet.getName(),
    workloadCarrecsSheetName: workloadCarrecsSheet.getName(),
    scheduleCacheRows: scheduleCacheRows.length,
    officialScheduleTemplateName: annualDataPermissions.templateName,
    mailRemainingDailyQuota,
  };
}

function grantAnnualDataPermissions() {
  const templateId = toDisplayString_(
    PropertiesService.getScriptProperties()
      .getProperty(CONFIG.officialScheduleTeachersDocIdPropertyName)
  );

  if (!templateId) {
    throw new Error(
      `Falta la propietat de script "${CONFIG.officialScheduleTeachersDocIdPropertyName}".`
    );
  }

  const result = verifyAnnualDataDrivePermissions_(templateId);
  result.scheduleCacheRows = getScheduleCacheRows_().length;
  return result;
}

function verifyAnnualDataDrivePermissions_(templateId) {
  const templateFile = DriveApp.getFileById(templateId);
  if (templateFile.getMimeType() !== MimeType.GOOGLE_DOCS) {
    throw new Error('El document configurat per a Dades anuals no es un document de Google.');
  }

  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyyMMdd-HHmmss'
  );
  let temporaryFile = null;

  try {
    temporaryFile = templateFile.makeCopy(`teacher-admin-permission-check-${timestamp}`);
    const exportedFile = exportGoogleDocAsDocx_(temporaryFile.getId());
    if (!exportedFile.getBytes().length) {
      throw new Error('La comprovacio de permisos ha generat un DOCX buit.');
    }

    return {
      ok: true,
      message: 'Permisos de Dades anuals concedits i comprovats correctament.',
      templateName: templateFile.getName(),
      exportedBytes: exportedFile.getBytes().length,
      temporaryCopyTrashed: true,
    };
  } finally {
    if (temporaryFile) {
      temporaryFile.setTrashed(true);
    }
  }
}

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

    const accessEntries = getAccessGrantedRoles_();
    if (accessEntries.length === 0) {
      return {
        allowed: false,
        email: userEmail,
        message: `Falta configurar la propietat de script "${CONFIG.accessGrantedPropertyName}".`,
      };
    }

    const directEmails = accessEntries
      .map(normalizeEmail_)
      .filter((entry) => entry.indexOf('@') !== -1);
    const roles = accessEntries.filter((entry) => normalizeEmail_(entry).indexOf('@') === -1);
    const peopleByRole = getPeopleByAccessRole_();
    const people = [];
    roles.forEach((role) => {
      const assignedPeople = peopleByRole.get(normalizeText_(role)) || [];
      assignedPeople.forEach((person) => people.push(person));
    });

    const authorizedEmails = getEmailsForPeople_(people);
    directEmails.forEach((email) => authorizedEmails.add(email));
    const allowed = authorizedEmails.has(userEmail);

    return {
      allowed,
      email: userEmail,
      accessEntries,
      roles,
      directEmails,
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

function getProfessorsData() {
  assertUserAccess_();
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
  assertUserAccess_();
  throw new Error('La baixa nomes es pot gestionar amb el flux Donar de baixa / Donar d\'alta.');
}

function updateActiu(rowNumbers, value) {
  assertUserAccess_();
  return updateBooleanColumn_(rowNumbers, CONFIG.activeColumn, value);
}

function startLeaveAbsence(rowNumber, leaveData) {
  assertUserAccess_();
  const normalizedRowNumber = normalizeRowNumber_(rowNumber);
  const data = normalizeLeaveData_(leaveData, true);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let result;
  let notificationContext;
  let emailContext;

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
    const substituteTeacher = sheet
      .getRange(substituteRowNumber, 1, 1, CONFIG.editableColumnCount)
      .getValues()[0];

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
    emailContext = {
      originalTeacherName: buildFullName_(teacher),
      substituteTeacherName: buildFullName_(substituteTeacher),
    };
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

  let emailError = null;
  try {
    sendLeaveAbsenceNotificationEmail_(
      emailContext.originalTeacherName,
      emailContext.substituteTeacherName
    );
  } catch (error) {
    emailError = error;
    console.error(JSON.stringify({
      event: 'leaveAbsence.email.error',
      context: notificationContext,
      error: error && error.message ? error.message : String(error),
    }));
  }

  notifyScheduleCacheRebuild_(notificationContext);
  if (emailError) throw emailError;
  return result;
}

function endLeaveAbsence(rowNumber, leaveData) {
  assertUserAccess_();
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
  assertUserAccess_();
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
  assertUserAccess_();
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
  assertUserAccess_();
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

function createAnnualTeacherDataDocx(rowNumbers) {
  assertUserAccess_();
  const uniqueRows = normalizeRowNumbers_(rowNumbers);
  const properties = PropertiesService.getScriptProperties();
  const templateId = toDisplayString_(
    properties.getProperty(CONFIG.officialScheduleTeachersDocIdPropertyName)
  );

  if (!templateId) {
    throw new Error(
      `Falta la propietat de script "${CONFIG.officialScheduleTeachersDocIdPropertyName}".`
    );
  }

  const templateFile = DriveApp.getFileById(templateId);
  if (templateFile.getMimeType() !== MimeType.GOOGLE_DOCS) {
    throw new Error('El document configurat per a Dades anuals no es un document de Google.');
  }

  const dbSheet = getDbSheet_();
  const teachers = uniqueRows.map((rowNumber) => {
    return dbSheet
      .getRange(rowNumber, 1, 1, CONFIG.editableColumnCount)
      .getDisplayValues()[0];
  });
  const teacherCarrecs = getTeacherCarrecs_();
  const scheduleCacheRows = getScheduleCacheRows_();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyyMMdd-HHmmss'
  );
  const temporaryFile = templateFile.makeCopy(`teacher-admin-dades-anuals-${timestamp}`);

  try {
    const templateDocx = exportGoogleDocAsDocx_(temporaryFile.getId());
    const outputFileName = `dades-anuals-professorat-${timestamp}.docx`;
    const teacherDocuments = teachers.map((teacher) => {
      const schedule = buildTeacherSchedule_(teacher, scheduleCacheRows);
      return {
        tagValues: buildAnnualTeacherTagValues_(
          teacher,
          teacherCarrecs,
          today,
          calculateScheduleTotals_(schedule)
        ),
        schedule,
      };
    });
    const mergedDocx = buildAnnualTeacherDataDocx_(templateDocx, teacherDocuments);

    return {
      fileName: outputFileName,
      mimeType: DOCX_MIME_TYPE,
      base64: Utilities.base64Encode(mergedDocx.getBytes()),
      rowCount: teachers.length,
    };
  } finally {
    try {
      temporaryFile.setTrashed(true);
    } catch (error) {
      console.error(JSON.stringify({
        event: 'annualTeacherData.tempFile.trashError',
        fileId: temporaryFile.getId(),
        error: error && error.message ? error.message : String(error),
      }));
    }
  }
}

function getTeacherCarrecs_() {
  const sheet = getWorkloadCarrecsSheet_();
  const lastRow = sheet.getLastRow();
  const teacherCarrecs = new Map();

  if (lastRow < 2) return teacherCarrecs;

  const values = sheet
    .getRange(2, 1, lastRow - 1, CARRECS_COLUMNS.isCarrec)
    .getDisplayValues();

  values.forEach((row) => {
    if (!isTruthyValue_(row[CARRECS_COLUMNS.isCarrec - 1])) return;

    const carrec = toDisplayString_(row[CARRECS_COLUMNS.carrec - 1]);
    if (!carrec) return;

    splitCommaList_(row[CARRECS_COLUMNS.asignado - 1]).forEach((teacherName) => {
      const teacherKey = normalizeText_(teacherName);
      if (teacherKey && !teacherCarrecs.has(teacherKey)) {
        teacherCarrecs.set(teacherKey, carrec);
      }
    });
  });

  return teacherCarrecs;
}

function buildAnnualTeacherTagValues_(teacher, teacherCarrecs, today, scheduleTotals) {
  const fullName = buildFullName_(teacher);
  const situacio = toDisplayString_(teacher[6]);
  const values = {
    cognom1: toDisplayString_(teacher[3]),
    cognom2: toDisplayString_(teacher[4]),
    nom: toDisplayString_(teacher[2]),
    DNI: toDisplayString_(teacher[8]),
    Carrec: teacherCarrecs.get(normalizeText_(fullName)) || '',
    Especialitat: toDisplayString_(teacher[0]),
    'FUNC.DEF': situacio === 'FUNC. DEF' ? 'x' : '',
    CS: situacio === 'CS' ? 'x' : '',
    'FUNC. SNS PLAÇA': situacio === 'FUNC. SNS PLAÇA' ? 'x' : '',
    INT: situacio === 'INT' ? 'x' : '',
    DATA: today,
  };

  ANNUAL_DATA_TOTAL_TAGS.forEach((tagName) => {
    values[tagName] = String(scheduleTotals[tagName] || 0);
  });
  values.TOTAL = String(ANNUAL_DATA_TOTAL_TAGS.reduce((sum, tagName) => {
    return sum + (Number(scheduleTotals[tagName]) || 0);
  }, 0));

  return values;
}

function getScheduleCacheRows_() {
  const spreadsheet = getScheduleSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(CONFIG.scheduleCacheSheetName);
  if (!sheet) {
    throw new Error(
      `No s'ha trobat el full "${CONFIG.scheduleCacheSheetName}" a ${CONFIG.scheduleRegistryName}.`
    );
  }

  const values = sheet.getDataRange().getValues();
  if (!values.length) {
    throw new Error(`El full "${CONFIG.scheduleCacheSheetName}" esta buit.`);
  }

  const requiredHeaders = [
    'row_id',
    'group',
    'source_teacher_code',
    'source_teacher_name',
    'source_teacher_original_code',
    'effective_teacher_code',
    'effective_teacher_name',
    'teacher_was_substituted',
    'subject_code',
    'subject_full_name',
    'classroom',
    'day',
    'slot',
  ];
  const headers = values[0].map((value) => toDisplayString_(value));
  const headerIndexes = new Map();
  headers.forEach((header, index) => {
    if (!header) return;
    if (headerIndexes.has(header)) {
      throw new Error(
        `El full "${CONFIG.scheduleCacheSheetName}" te la capcalera duplicada "${header}".`
      );
    }
    headerIndexes.set(header, index);
  });
  requiredHeaders.forEach((header) => {
    if (!headerIndexes.has(header)) {
      throw new Error(
        `El full "${CONFIG.scheduleCacheSheetName}" no te la capcalera obligatoria "${header}".`
      );
    }
  });

  return values.slice(1)
    .filter((row) => row.some((value) => toDisplayString_(value) !== ''))
    .map((row) => {
      const record = {};
      requiredHeaders.forEach((header) => {
        record[header] = row[headerIndexes.get(header)];
      });
      return record;
    });
}

function getScheduleSpreadsheet_() {
  const properties = PropertiesService.getScriptProperties();
  const registryId = toDisplayString_(
    properties.getProperty(CONFIG.scheduleRegistryPropertyName)
  ) || toDisplayString_(properties.getProperty(CONFIG.tablesPropertyName));
  if (!registryId) {
    throw new Error(
      `Falta la propietat de script "${CONFIG.scheduleRegistryPropertyName}" o `
      + `"${CONFIG.tablesPropertyName}" per llegir els horaris.`
    );
  }

  const registrySpreadsheet = SpreadsheetApp.openById(registryId);
  const registrySheet = registrySpreadsheet.getSheetByName(CONFIG.tablesSheetName);
  if (!registrySheet) {
    throw new Error(`No s'ha trobat el full "${CONFIG.tablesSheetName}" al registre d'horaris.`);
  }

  const lastRow = registrySheet.getLastRow();
  if (lastRow < 1) {
    throw new Error(`El full "${CONFIG.tablesSheetName}" del registre d'horaris esta buit.`);
  }

  const matches = registrySheet
    .getRange(1, 1, lastRow, 2)
    .getDisplayValues()
    .filter((row) => toDisplayString_(row[0]) === CONFIG.scheduleRegistryName);
  if (matches.length !== 1) {
    throw new Error(
      matches.length
        ? `Hi ha mes d'una entrada "${CONFIG.scheduleRegistryName}" al registre d'horaris.`
        : `No s'ha trobat "${CONFIG.scheduleRegistryName}" al registre d'horaris.`
    );
  }

  const spreadsheetId = toDisplayString_(matches[0][1]);
  if (!spreadsheetId) {
    throw new Error(
      `La fila "${CONFIG.scheduleRegistryName}" no te ID a la columna B.`
    );
  }
  return SpreadsheetApp.openById(spreadsheetId);
}

function buildTeacherSchedule_(teacher, scheduleCacheRows) {
  const teacherName = buildFullName_(teacher);
  const teacherCode = toDisplayString_(teacher[5]);
  if (!teacherCode) {
    throw new Error(`El professor "${teacherName}" no te codi REDUIT.`);
  }

  const codeKey = normalizeText_(teacherCode);
  const rows = Array.from({ length: 12 }, (_, index) => ({
    slot: index + 1,
    cells: Array.from({ length: 5 }, () => ({ items: [], itemIndexes: new Map() })),
  }));

  scheduleCacheRows.forEach((cacheRow) => {
    if (normalizeText_(cacheRow.effective_teacher_code) !== codeKey) return;

    const day = Number(cacheRow.day);
    const slot = Number(cacheRow.slot);
    if (
      !Number.isInteger(day) || day < 1 || day > 5
      || !Number.isInteger(slot) || slot < 1 || slot > 12
    ) {
      console.warn(JSON.stringify({
        event: 'annualTeacherData.invalidScheduleRow',
        teacherCode,
        rowId: toDisplayString_(cacheRow.row_id),
        day: cacheRow.day,
        slot: cacheRow.slot,
      }));
      return;
    }

    const cell = rows[slot - 1].cells[day - 1];
    const subjectCode = toDisplayString_(cacheRow.subject_code);
    const subjectName = toDisplayString_(cacheRow.subject_full_name)
      || subjectCode
      || 'Sense assignatura';
    const classroom = toDisplayString_(cacheRow.classroom);
    const itemKey = JSON.stringify([subjectName, classroom]);
    let item = cell.itemIndexes.get(itemKey);
    if (!item) {
      item = {
        subjectName,
        subjectCode,
        groups: [],
        classroom,
        colour: getScheduleSubjectColour_(subjectCode),
      };
      cell.itemIndexes.set(itemKey, item);
      cell.items.push(item);
    }

    const group = toDisplayString_(cacheRow.group);
    if (group && !item.groups.includes(group)) item.groups.push(group);
  });

  rows.forEach((row) => {
    row.cells.forEach((cell) => delete cell.itemIndexes);
  });

  return {
    teacherCode,
    teacherName,
    days: ['Dilluns', 'Dimarts', 'Dimecres', 'Dijous', 'Divendres'],
    printRows: rows.filter((row) => row.cells.some((cell) => cell.items.length > 0)),
  };
}

function getScheduleSubjectColour_(subjectCode) {
  const code = toDisplayString_(subjectCode).toUpperCase();
  if (['GUARDIA', 'GUARDIA_PATI'].includes(code)) {
    return { background: 'FFF4E5', accent: 'F59F00' };
  }
  if (['TUT', 'TUT_FAMILIES'].includes(code)) {
    return { background: 'FFF0F6', accent: 'F06595' };
  }
  if (code === '3R') {
    return { background: 'FFF9DB', accent: 'F2C94C' };
  }
  if ([
    'RC_ESO', 'RC_FP_BAT', 'RDEP', 'RDIM1', 'RDIM2', 'RDIR',
    'REAP', 'REC', 'CARREC', 'FCT', 'TREC',
  ].includes(code)) {
    return { background: 'EEF6FF', accent: '4DABF7' };
  }
  return { background: 'EDF7F6', accent: '7FC7BD' };
}

function calculateScheduleTotals_(schedule) {
  const totals = {};
  const context = buildScheduleClassificationContext_(schedule);
  ANNUAL_DATA_TOTAL_TAGS.forEach((tagName) => {
    totals[tagName] = 0;
  });

  schedule.printRows.forEach((row) => {
    row.cells.forEach((cell) => {
      cell.items.forEach((item) => {
        const classification = classifyScheduleItem_(item, context);
        if (classification.tag) totals[classification.tag] += 1;
      });
    });
  });
  return totals;
}

function classifyScheduleItem_(item, context) {
  const titleKey = normalizeScheduleKey_(item.subjectName);
  const compactTitle = compactScheduleKey_(item.subjectName);
  const groupCategories = getScheduleGroupCategories_(item.groups);
  const groupKeys = item.groups.map(compactScheduleKey_);

  if (compactTitle === 'orientacio') {
    const hasGroup = item.groups.length > 0;
    const hasClassroom = Boolean(toDisplayString_(item.classroom));
    if (!hasGroup && !hasClassroom) return { tag: 'CARREC', reason: '' };
    if (hasGroup && hasClassroom) return { tag: 'HORES_ESO', reason: '' };
    return {
      tag: '',
      reason: 'ORIENTACIÓ ha de tenir grup i classe, o no tenir-ne cap dels dos',
    };
  }

  if (SCHEDULE_EXACT_TITLE_TAGS[compactTitle]) {
    return { tag: SCHEDULE_EXACT_TITLE_TAGS[compactTitle], reason: '' };
  }
  if (compactTitle === 'aulaacollida' && groupKeys.includes('aco')) {
    return { tag: 'HORES_ESO', reason: '' };
  }
  if (compactTitle === 'senseassignatura' && groupKeys.includes('ocu')) {
    return { tag: 'HORES_ESO', reason: '' };
  }
  if (compactTitle === 'siei' && groupKeys.includes('siei')) {
    return { tag: 'HORES_ESO', reason: '' };
  }

  if (compactTitle === 'fct') {
    return { tag: 'HORES_FCT', reason: '' };
  }

  if (compactTitle === 'tutoria' || compactTitle === 'tutfamilies') {
    const ownTutorialClassification = classifyTutorialGroups_(item, groupCategories);
    if (ownTutorialClassification.tag) return ownTutorialClassification;
    if (compactTitle === 'tutfamilies' && context && context.tutorialFallbackTag) {
      return { tag: context.tutorialFallbackTag, reason: '' };
    }
    if (compactTitle === 'tutfamilies' && context && context.tutorialFallbackReason) {
      return { tag: '', reason: context.tutorialFallbackReason };
    }
    return ownTutorialClassification;
  }

  const directiveTitles = SCHEDULE_DIRECTIVE_TITLES.map(compactScheduleKey_);
  if (directiveTitles.includes(compactTitle)) {
    return { tag: 'CARREC_DIRECTIU', reason: '' };
  }
  if (titleKey.includes('carrec')) {
    return { tag: 'CARREC', reason: '' };
  }
  if (titleKey.includes('reunio')) {
    return { tag: 'REUNIONS', reason: '' };
  }
  if (titleKey.includes('guardia')) {
    return { tag: 'GUARDIES', reason: '' };
  }

  if (groupCategories.length > 1) {
    return {
      tag: '',
      reason: `Grups de categories diferents: ${groupCategories.join(', ')}`,
    };
  }
  if (groupCategories.length === 1) {
    return {
      tag: {
        ESO: 'HORES_ESO',
        PFI: 'HORES_PFI',
        BAT: 'HORES_BAT',
        FP: 'HORES_CICLES',
      }[groupCategories[0]] || '',
      reason: '',
    };
  }

  return { tag: '', reason: 'Titol i grups fora de les categories configurades' };
}

function classifyTutorialGroups_(item, groupCategories) {
  const groupKeys = item.groups.map(compactScheduleKey_);
  if (groupKeys.includes('aco')) return { tag: 'TUT_ESO', reason: '' };
  if (groupKeys.includes('pfi')) return { tag: 'TUT_FP', reason: '' };
  if (groupCategories.length > 1) {
    return {
      tag: '',
      reason: `Tutoria amb grups de categories diferents: ${groupCategories.join(', ')}`,
    };
  }
  if (groupCategories.length === 1) {
    const tag = {
      ESO: 'TUT_ESO',
      BAT: 'TUT_BAT',
      FP: 'TUT_FP',
    }[groupCategories[0]];
    if (tag) return { tag, reason: '' };
  }
  return { tag: '', reason: 'Tutoria sense cap grup ESO, BAT o FP reconegut' };
}

function buildScheduleClassificationContext_(schedule) {
  const tutorialTags = new Set();
  schedule.printRows.forEach((row) => {
    row.cells.forEach((cell) => {
      cell.items.forEach((item) => {
        if (compactScheduleKey_(item.subjectName) !== 'tutoria') return;
        const classification = classifyTutorialGroups_(
          item,
          getScheduleGroupCategories_(item.groups)
        );
        if (classification.tag) tutorialTags.add(classification.tag);
      });
    });
  });

  if (tutorialTags.size === 1) {
    return { tutorialFallbackTag: Array.from(tutorialTags)[0], tutorialFallbackReason: '' };
  }
  return {
    tutorialFallbackTag: '',
    tutorialFallbackReason: tutorialTags.size
      ? `TUT_FAMILIES amb TUTORIA de categories diferents: ${Array.from(tutorialTags).join(', ')}`
      : 'TUT_FAMILIES sense cap TUTORIA classificable del mateix professor',
  };
}

function getScheduleGroupCategories_(groups) {
  const categories = new Set();
  const configuredGroups = {};
  Object.keys(SCHEDULE_GROUPS).forEach((category) => {
    SCHEDULE_GROUPS[category].forEach((group) => {
      configuredGroups[compactScheduleKey_(group)] = category;
    });
  });
  groups.forEach((group) => {
    const category = configuredGroups[compactScheduleKey_(group)];
    if (category) categories.add(category);
  });
  return Array.from(categories);
}

function normalizeScheduleKey_(value) {
  return normalizeText_(value).replace(/\s+/g, ' ');
}

function compactScheduleKey_(value) {
  return normalizeScheduleKey_(value).replace(/[^a-z0-9]/g, '');
}

function auditAnnualScheduleClassification() {
  const cacheRows = getScheduleCacheRows_();
  const teachersByCode = new Map();
  cacheRows.forEach((row) => {
    const code = toDisplayString_(row.effective_teacher_code);
    if (!code) return;
    const codeKey = normalizeText_(code);
    if (!teachersByCode.has(codeKey)) {
      teachersByCode.set(codeKey, {
        code,
        name: toDisplayString_(row.effective_teacher_name) || code,
      });
    }
  });

  const unclassified = new Map();
  let bubbleCount = 0;
  let classifiedBubbleCount = 0;
  teachersByCode.forEach((teacher) => {
    const teacherRow = ['', '', teacher.name, '', '', teacher.code];
    const schedule = buildTeacherSchedule_(teacherRow, cacheRows);
    const context = buildScheduleClassificationContext_(schedule);
    schedule.printRows.forEach((row) => {
      row.cells.forEach((cell) => {
        cell.items.forEach((item) => {
          bubbleCount += 1;
          const classification = classifyScheduleItem_(item, context);
          if (classification.tag) {
            classifiedBubbleCount += 1;
            return;
          }

          const key = JSON.stringify([
            normalizeScheduleKey_(item.subjectName),
            item.groups.map(compactScheduleKey_).sort(),
            classification.reason,
          ]);
          if (!unclassified.has(key)) {
            unclassified.set(key, {
              title: item.subjectName,
              groups: item.groups.slice(),
              reason: classification.reason,
              occurrences: 0,
              teachers: new Set(),
            });
          }
          const record = unclassified.get(key);
          record.occurrences += 1;
          record.teachers.add(teacher.name);
        });
      });
    });
  });

  const unresolved = Array.from(unclassified.values())
    .map((record) => ({
      title: record.title,
      groups: record.groups,
      reason: record.reason,
      occurrences: record.occurrences,
      teacherCount: record.teachers.size,
      teachers: Array.from(record.teachers).sort(),
    }))
    .sort((left, right) => left.title.localeCompare(right.title, 'ca'));

  return {
    ok: true,
    teacherCount: teachersByCode.size,
    bubbleCount,
    classifiedBubbleCount,
    unclassifiedBubbleCount: bubbleCount - classifiedBubbleCount,
    unresolved,
  };
}

function exportGoogleDocAsDocx_(fileId) {
  const exportUrl = `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}/export`
    + `?mimeType=${encodeURIComponent(DOCX_MIME_TYPE)}`;
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`,
    },
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(
      `No s'ha pogut exportar la plantilla Dades anuals (${response.getResponseCode()}).`
    );
  }

  return response.getBlob();
}

function buildAnnualTeacherDataDocx_(templateDocx, teacherDocuments) {
  if (!teacherDocuments.length) {
    throw new Error('Cal seleccionar almenys una fila.');
  }

  const templateZip = Utilities.newBlob(
    templateDocx.getBytes(),
    'application/zip',
    'dades-anuals-template.zip'
  );
  const entries = Utilities.unzip(templateZip);
  const documentEntry = entries.find((entry) => entry.getName() === 'word/document.xml');
  if (!documentEntry) {
    throw new Error('La plantilla DOCX no conte word/document.xml.');
  }

  const documentXml = documentEntry.getDataAsString('UTF-8');
  const bodyMatch = documentXml.match(/<w:body\b[^>]*>([\s\S]*?)<\/w:body>/);
  if (!bodyMatch) {
    throw new Error('No s\'ha pogut identificar el cos de la plantilla DOCX.');
  }

  const templateBody = bodyMatch[1];
  const teacherBodies = teacherDocuments.map((teacherDocument, index) => {
    let body = replaceScheduleTagWithTable_(templateBody, teacherDocument.schedule);
    body = replaceAnnualTeacherTagsInXml_(body, teacherDocument.tagValues);
    body = removeDocxBookmarks_(body);
    body = forceDocxSectionsToNextPage_(body);
    body = addBlankFirstPageFooterReference_(body);
    if (index < teacherDocuments.length - 1) {
      body = convertFinalSectionToBreak_(body);
    }
    return body;
  });
  const mergedDocumentXml = documentXml.replace(bodyMatch[1], teacherBodies.join(''));
  const outputEntries = entries.map((entry) => {
    if (entry.getName() === 'word/document.xml') {
      return Utilities.newBlob(mergedDocumentXml, 'application/xml', 'word/document.xml');
    }
    if (entry.getName() === 'word/_rels/document.xml.rels') {
      return Utilities.newBlob(
        addBlankFooterRelationship_(entry.getDataAsString('UTF-8')),
        'application/xml',
        'word/_rels/document.xml.rels'
      );
    }
    if (entry.getName() === '[Content_Types].xml') {
      return Utilities.newBlob(
        addBlankFooterContentType_(entry.getDataAsString('UTF-8')),
        'application/xml',
        '[Content_Types].xml'
      );
    }
    return entry;
  });
  outputEntries.push(createBlankFirstPageFooterBlob_());

  return Utilities.zip(outputEntries);
}

function replaceScheduleTagWithTable_(xml, schedule) {
  const encodedTag = escapeXmlText_('<<HORARI>>');
  let replacements = 0;
  const result = xml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (paragraphXml) => {
    if (!paragraphXml.includes(encodedTag)) return paragraphXml;
    replacements += 1;
    return buildScheduleTableXml_(schedule);
  });

  if (replacements !== 1) {
    throw new Error(
      replacements
        ? 'La plantilla Dades anuals conte mes d\'una etiqueta <<HORARI>>.'
        : 'No s\'ha trobat l\'etiqueta <<HORARI>> a la plantilla Dades anuals.'
    );
  }
  return result;
}

function buildScheduleTableXml_(schedule) {
  if (!schedule.printRows.length) {
    return '<w:p><w:pPr><w:spacing w:after="0" w:before="0"/></w:pPr>'
      + '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>'
      + '<w:i/><w:color w:val="667085"/><w:sz w:val="16"/></w:rPr>'
      + '<w:t>No hi ha cap horari efectiu disponible.</w:t></w:r></w:p>';
  }

  const slotWidth = 400;
  const dayWidth = 1680;
  const tableWidth = slotWidth + (dayWidth * schedule.days.length);
  const grid = [slotWidth, ...schedule.days.map(() => dayWidth)]
    .map((width) => `<w:gridCol w:w="${width}"/>`)
    .join('');
  const headerCells = [
    buildScheduleHeaderCellXml_('', slotWidth),
    ...schedule.days.map((day) => buildScheduleHeaderCellXml_(day, dayWidth)),
  ].join('');
  const headerRow = '<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>'
    + headerCells
    + '</w:tr>';
  const bodyRows = schedule.printRows.map((row) => {
    const cells = [
      buildScheduleSlotCellXml_(row.slot, slotWidth),
      ...row.cells.map((cell) => buildScheduleDayCellXml_(cell, dayWidth)),
    ].join('');
    return `<w:tr><w:trPr><w:cantSplit/></w:trPr>${cells}</w:tr>`;
  }).join('');

  return '<w:tbl><w:tblPr>'
    + `<w:tblW w:w="${tableWidth}" w:type="dxa"/>`
    + '<w:jc w:val="center"/><w:tblLayout w:type="fixed"/>'
    + '<w:tblBorders>'
    + '<w:top w:val="single" w:sz="4" w:color="D8DEE8"/>'
    + '<w:left w:val="single" w:sz="4" w:color="D8DEE8"/>'
    + '<w:bottom w:val="single" w:sz="4" w:color="D8DEE8"/>'
    + '<w:right w:val="single" w:sz="4" w:color="D8DEE8"/>'
    + '<w:insideH w:val="single" w:sz="4" w:color="D8DEE8"/>'
    + '<w:insideV w:val="single" w:sz="4" w:color="D8DEE8"/>'
    + '</w:tblBorders>'
    + '<w:tblCellMar><w:top w:w="40" w:type="dxa"/>'
    + '<w:left w:w="40" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/>'
    + '<w:right w:w="40" w:type="dxa"/></w:tblCellMar>'
    + '</w:tblPr>'
    + `<w:tblGrid>${grid}</w:tblGrid>${headerRow}${bodyRows}</w:tbl>`;
}

function buildScheduleHeaderCellXml_(text, width) {
  return '<w:tc><w:tcPr>'
    + `<w:tcW w:w="${width}" w:type="dxa"/>`
    + '<w:shd w:val="clear" w:fill="E9EEF5"/><w:vAlign w:val="center"/>'
    + '</w:tcPr><w:p><w:pPr><w:jc w:val="center"/>'
    + '<w:spacing w:before="0" w:after="0"/></w:pPr>'
    + '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>'
    + '<w:b/><w:color w:val="344054"/><w:sz w:val="15"/></w:rPr>'
    + `<w:t>${escapeXmlText_(text)}</w:t></w:r></w:p></w:tc>`;
}

function buildScheduleSlotCellXml_(slot, width) {
  return '<w:tc><w:tcPr>'
    + `<w:tcW w:w="${width}" w:type="dxa"/>`
    + '<w:shd w:val="clear" w:fill="F8FAFC"/><w:vAlign w:val="center"/>'
    + '</w:tcPr><w:p><w:pPr><w:jc w:val="center"/>'
    + '<w:spacing w:before="0" w:after="0"/></w:pPr>'
    + '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>'
    + '<w:b/><w:color w:val="344054"/><w:sz w:val="14"/></w:rPr>'
    + `<w:t>${slot}</w:t></w:r></w:p></w:tc>`;
}

function buildScheduleDayCellXml_(cell, width) {
  const content = cell.items.length
    ? cell.items.map(buildScheduleItemParagraphXml_).join('')
    : '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>';
  return '<w:tc><w:tcPr>'
    + `<w:tcW w:w="${width}" w:type="dxa"/>`
    + '<w:vAlign w:val="center"/></w:tcPr>'
    + `${content}</w:tc>`;
}

function buildScheduleItemParagraphXml_(item) {
  const subject = escapeXmlText_(item.subjectName);
  const metaLines = [];
  if (item.groups.length) metaLines.push(`Grup: ${item.groups.join(', ')}`);
  if (item.classroom) metaLines.push(`Classe: ${item.classroom}`);
  const metaRuns = metaLines.map((line) => {
    return '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>'
      + '<w:color w:val="344054"/><w:sz w:val="13"/></w:rPr>'
      + `<w:br/><w:t>${escapeXmlText_(line)}</w:t></w:r>`;
  }).join('');

  return '<w:p><w:pPr>'
    + '<w:keepLines/><w:spacing w:before="20" w:after="20" w:line="160" w:lineRule="atLeast"/>'
    + '<w:jc w:val="left"/>'
    + '<w:ind w:left="50" w:right="25"/>'
    + '<w:pBdr>'
    + `<w:left w:val="single" w:sz="18" w:space="3" w:color="${item.colour.accent}"/>`
    + '</w:pBdr>'
    + `<w:shd w:val="clear" w:fill="${item.colour.background}"/>`
    + '</w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>'
    + '<w:b/><w:color w:val="134E4A"/><w:sz w:val="14"/></w:rPr>'
    + `<w:t>${subject}</w:t></w:r>${metaRuns}</w:p>`;
}

function replaceAnnualTeacherTagsInXml_(xml, tagValues) {
  let result = xml;
  Object.keys(tagValues).forEach((tagName) => {
    const encodedTag = escapeXmlText_(`<<${tagName}>>`);
    if (!result.includes(encodedTag)) {
      throw new Error(`No s'ha trobat l'etiqueta <<${tagName}>> a la plantilla Dades anuals.`);
    }
    const encodedValue = escapeXmlText_(tagValues[tagName]);
    result = result.split(encodedTag).join(encodedValue);
  });
  return result;
}

function removeDocxBookmarks_(xml) {
  return xml
    .replace(/<w:bookmarkStart\b[^>]*\/>/g, '')
    .replace(/<w:bookmarkEnd\b[^>]*\/>/g, '');
}

function forceDocxSectionsToNextPage_(xml) {
  return xml.replace(/<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/g, (sectionXml) => {
    const withoutType = sectionXml.replace(/<w:type\b[^>]*\/>/g, '');
    return withoutType.replace(
      /^(<w:sectPr\b[^>]*>)/,
      '$1<w:type w:val="nextPage"/>'
    );
  });
}

function addBlankFirstPageFooterReference_(xml) {
  return xml.replace(/<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/g, (sectionXml) => {
    if (
      !/<w:titlePg\b/.test(sectionXml)
      || /<w:footerReference\b[^>]*w:type="first"/.test(sectionXml)
    ) {
      return sectionXml;
    }
    return sectionXml.replace(
      /^(<w:sectPr\b[^>]*>)/,
      `$1<w:footerReference w:type="first" r:id="${DOCX_BLANK_FIRST_FOOTER_REL_ID}"/>`
    );
  });
}

function addBlankFooterRelationship_(relationshipsXml) {
  if (relationshipsXml.includes(`Id="${DOCX_BLANK_FIRST_FOOTER_REL_ID}"`)) {
    return relationshipsXml;
  }
  const relationship = '<Relationship'
    + ` Id="${DOCX_BLANK_FIRST_FOOTER_REL_ID}"`
    + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"'
    + ' Target="footer-teacher-admin-blank.xml"/>';
  return relationshipsXml.replace('</Relationships>', `${relationship}</Relationships>`);
}

function addBlankFooterContentType_(contentTypesXml) {
  const partName = '/word/footer-teacher-admin-blank.xml';
  if (contentTypesXml.includes(`PartName="${partName}"`)) {
    return contentTypesXml;
  }
  const override = `<Override PartName="${partName}"`
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>';
  return contentTypesXml.replace('</Types>', `${override}</Types>`);
}

function createBlankFirstPageFooterBlob_() {
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + '<w:p/>'
    + '</w:ftr>';
  return Utilities.newBlob(xml, 'application/xml', DOCX_BLANK_FIRST_FOOTER_PATH);
}

function convertFinalSectionToBreak_(bodyXml) {
  const match = bodyXml.match(/^([\s\S]*)(<w:sectPr\b[\s\S]*?<\/w:sectPr>)(\s*)$/);
  if (!match) {
    throw new Error('No s\'ha pogut conservar el salt de seccio de la plantilla DOCX.');
  }

  return `${match[1]}<w:p><w:pPr>${match[2]}</w:pPr></w:p>${match[3]}`;
}

function escapeXmlText_(value) {
  return toDisplayString_(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
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
  return getRegisteredSpreadsheet_(CONFIG.dbRegistryName);
}

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

  const dbSpreadsheetId = String(match[1] || '').trim();

  if (!dbSpreadsheetId) {
    throw new Error(`La fila "${registryName}" no te ID a la columna B.`);
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

function sendLeaveAbsenceNotificationEmail_(originalTeacherName, substituteTeacherName) {
  const originalTeacher = toDisplayString_(originalTeacherName);
  const substituteTeacher = toDisplayString_(substituteTeacherName);

  if (!originalTeacher || !substituteTeacher) {
    throw new Error('No es pot enviar el correu de baixa: falten noms de professorat.');
  }

  MailApp.sendEmail({
    to: CONFIG.leaveNotificationRecipient,
    subject: CONFIG.leaveNotificationSubject,
    body: [
      'Hola a tots,',
      `A partir de demà s'incorpora en/na ${substituteTeacher}, en substitució de ${originalTeacher}`,
      'Benvingut/da!',
      '',
      'Salut,',
      '',
    ].join('\n'),
  });
}

function notifyScheduleCacheRebuild_(context, endpointOverride) {
  const properties = PropertiesService.getScriptProperties();
  const endpointUrl = toDisplayString_(
    endpointOverride
      || properties.getProperty(CONFIG.cacheRebuildUrlPropertyName)
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

  if (
    statusCode === 404
    && !endpointOverride
    && endpointUrl !== CONFIG.defaultCacheRebuildUrl
  ) {
    const fallbackMessage = 'configured endpoint returned 404; retrying default endpoint';
    console.warn(JSON.stringify({
      event: 'scheduleCacheRebuild.notify.fallback',
      context: logContext,
      endpointUrl: requestLogUrl,
      fallbackEndpointUrl: CONFIG.defaultCacheRebuildUrl,
      statusCode,
    }));
    appendScheduleCacheRebuildLog_('scheduleCacheRebuild.notify.fallback', {
      context: logContext,
      endpointUrl: requestLogUrl,
      hasToken: true,
      statusCode,
      responseText,
      error: fallbackMessage,
    });
    return notifyScheduleCacheRebuild_(logContext, CONFIG.defaultCacheRebuildUrl);
  }

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
