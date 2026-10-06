const CONFIG = {
  tablesPropertyName: 'Tables',
  accessGrantedPropertyName: 'access_granted',
  tablesSheetName: 'tables',
  dbRegistryName: 'Dades de professors',
  workloadRegistryName: 'Càrrega lectiva',
  dbSheetName: 'Llista',
  workloadProfessorsSheetName: 'professors',
  workloadCarrecsSheetName: 'carrecs',
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

const ANNUAL_DATA_PENDING_TAGS = [
  'HORARI',
  'HORES_ESO',
  'HORES_PFI',
  'HORES_BAT',
  'HORES_CICLES',
  'HORES_FCT',
  'TUT_ESO',
  'TUT_BAT',
  'TUT_FP',
  'CARREC_DIRECTIU',
  'CARREC',
  'REUNIONS',
  'GUARDIES',
];

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

  return verifyAnnualDataDrivePermissions_(templateId);
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
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyyMMdd-HHmmss'
  );
  const temporaryFile = templateFile.makeCopy(`teacher-admin-dades-anuals-${timestamp}`);

  try {
    const templateDocx = exportGoogleDocAsDocx_(temporaryFile.getId());
    const mergedDocx = buildAnnualTeacherDataDocx_(
      templateDocx,
      teachers.map((teacher) => {
        return buildAnnualTeacherTagValues_(teacher, teacherCarrecs, today);
      }),
      `dades-anuals-professorat-${timestamp}.docx`
    );

    return {
      fileName: mergedDocx.getName(),
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

function buildAnnualTeacherTagValues_(teacher, teacherCarrecs, today) {
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

  ANNUAL_DATA_PENDING_TAGS.forEach((tagName) => {
    values[tagName] = '';
  });

  return values;
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

  return response.getBlob().setName('dades-anuals-template.docx');
}

function buildAnnualTeacherDataDocx_(templateDocx, teacherTagValues, fileName) {
  if (!teacherTagValues.length) {
    throw new Error('Cal seleccionar almenys una fila.');
  }

  const entries = Utilities.unzip(templateDocx);
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
  const teacherBodies = teacherTagValues.map((tagValues, index) => {
    let body = replaceAnnualTeacherTagsInXml_(templateBody, tagValues);
    body = removeDocxBookmarks_(body);
    body = forceDocxSectionsToNextPage_(body);
    body = addBlankFirstPageFooterReference_(body);
    if (index < teacherTagValues.length - 1) {
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

  return Utilities.zip(outputEntries, fileName);
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
