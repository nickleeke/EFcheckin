/* ============================================================
   Caseload Dashboard — Google Apps Script Backend
   Storage: Google Sheets (per-user, auto-provisioned)
   v3 — multi-user with per-user data isolation

   DEPLOYMENT: Must deploy as "Execute as: User accessing the web app"
   so that Session.getActiveUser() returns the actual user and
   UserProperties are scoped per-user (FERPA compliance).
   ============================================================ */

// ───── Constants ─────
const SHEET_STUDENTS = 'Students';
const SHEET_CHECKINS = 'CheckIns';
const SHEET_COTEACHERS = 'CoTeachers';
const SHEET_EVALUATIONS = 'Evaluations';
const SHEET_PROGRESS = 'ProgressReporting';
const SHEET_IEP_MEETINGS = 'IEPMeetings';

const GPA_MAP = {
  'A':4.0, 'A-':3.7,
  'B+':3.3, 'B':3.0, 'B-':2.7,
  'C+':2.3, 'C':2.0, 'C-':1.7,
  'D+':1.3, 'D':1.0, 'D-':0.7,
  'F':0.0
};

var SUPERUSER_EMAIL = 'nicholas.leeke@rpsmn.org';

// ───── Permission Keys & Role Defaults ─────
var PERMISSION_KEYS = [
  'viewAcademics', 'editAcademics',
  'viewCheckins', 'createCheckins',
  'viewContacts', 'editContacts',
  'viewEvals', 'editEvals',
  'viewProgress', 'editProgress',
  'viewDueProcess',
  'editStudentInfo', 'editGoals'
];

var ROLE_DEFAULT_PERMISSIONS = {
  'co-teacher': {
    viewAcademics: true, editAcademics: true,
    viewCheckins: true, createCheckins: true,
    viewContacts: true, editContacts: true,
    viewEvals: false, editEvals: false,
    viewProgress: false, editProgress: false,
    viewDueProcess: false,
    editStudentInfo: true, editGoals: false
  },
  'service-provider': {
    viewAcademics: true, editAcademics: false,
    viewCheckins: true, createCheckins: false,
    viewContacts: true, editContacts: false,
    viewEvals: true, editEvals: false,
    viewProgress: true, editProgress: true,
    viewDueProcess: true,
    editStudentInfo: false, editGoals: false
  },
  'para': {
    viewAcademics: true, editAcademics: false,
    viewCheckins: false, createCheckins: false,
    viewContacts: true, editContacts: false,
    viewEvals: false, editEvals: false,
    viewProgress: false, editProgress: false,
    viewDueProcess: false,
    editStudentInfo: false, editGoals: false
  }
};

// ───── User Identity ─────

/** Get the authenticated user's email. Requires "Execute as: User accessing the web app". */
function getCurrentUserEmail_() {
  var email = Session.getActiveUser().getEmail();
  if (!email) {
    throw new Error('Unable to determine user identity. Please ensure you are signed in with your school Google account.');
  }
  return email.toLowerCase();
}

/** Check whether the current user has completed onboarding. */
function getUserStatus() {
  var email = getCurrentUserEmail_();

  // Check if SPED Lead
  if (isSpedLead_(email)) {
    return {
      hasData: true,
      role: 'sped-lead',
      permissions: null,  // null = unrestricted due process access
      connectedCaseloads: getSpedLeadCaseloads_(email),
      isSuperuser: false
    };
  }

  var props = PropertiesService.getUserProperties();
  var ssId = props.getProperty('spreadsheet_id');

  // If no spreadsheet but a backup own spreadsheet exists, restore it
  if (!ssId) {
    var ownSsId = props.getProperty('own_spreadsheet_id');
    if (ownSsId) {
      try {
        SpreadsheetApp.openById(ownSsId);
        props.setProperty('spreadsheet_id', ownSsId);
        props.deleteProperty('own_spreadsheet_id');
        ssId = ownSsId;
      } catch(e) {
        props.deleteProperty('own_spreadsheet_id');
      }
    }
  }

  // Check for pending co-teacher invites in ScriptProperties
  var invite = null;
  try {
    var scriptProps = PropertiesService.getScriptProperties();
    var inviteRaw = scriptProps.getProperty('coteacher_invite_' + email);
    if (inviteRaw) {
      invite = JSON.parse(inviteRaw);
    }
  } catch(e) {}

  // Resolve caller's role and permissions
  var callerRole = null;
  var callerPermissions = null;
  if (ssId) {
    try {
      var ss = SpreadsheetApp.openById(ssId);
      var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
      callerRole = getCallerRole_(ctSheet, email);
      if (callerRole !== 'caseload-manager') {
        callerPermissions = getCallerPermissions_(ctSheet, email);
      }
      // null permissions = caseload manager = full access (frontend convention)
    } catch(e) { /* sheet access failed, treat as full access */ }
  }

  return {
    isNewUser: !ssId,
    email: email,
    pendingInvite: invite,
    isSuperuser: email === SUPERUSER_EMAIL,
    role: callerRole,
    permissions: callerPermissions
  };
}

// ───── Per-User Spreadsheet Resolution ─────

/** Get the current user's spreadsheet — works bound or as web app. */
function getSS_() {
  // When running as a bound script (from the spreadsheet menu), use active spreadsheet
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch(e) {}

  // When running as a web app, look up user's spreadsheet from UserProperties
  var props = PropertiesService.getUserProperties();
  var ssId = props.getProperty('spreadsheet_id');
  if (!ssId) {
    throw new Error('No spreadsheet configured. Please complete onboarding.');
  }

  try {
    return SpreadsheetApp.openById(ssId);
  } catch(e) {
    // Spreadsheet was deleted or access revoked; clear stored ID so onboarding re-triggers
    props.deleteProperty('spreadsheet_id');
    throw new Error('Your data spreadsheet is no longer accessible. Please reload to set up a new one.');
  }
}

// ───── Spreadsheet Provisioning ─────

/** Create a new Google Sheet in the user's Drive and store its ID. */
function provisionUserSpreadsheet() {
  var email = getCurrentUserEmail_();
  var props = PropertiesService.getUserProperties();

  // Guard: if already provisioned, return existing
  var existingId = props.getProperty('spreadsheet_id');
  if (existingId) {
    try {
      var existing = SpreadsheetApp.openById(existingId);
      return { success: true, spreadsheetId: existingId, spreadsheetUrl: existing.getUrl() };
    } catch(e) {
      // Spreadsheet was deleted or access revoked; fall through to create new one
    }
  }

  // Create new spreadsheet in user's Drive
  var ss = SpreadsheetApp.create('Caseload Dashboard Data');

  // Initialize Students sheet (rename default Sheet1)
  var studentsSheet = ss.getSheetByName('Sheet1');
  if (studentsSheet) {
    studentsSheet.setName(SHEET_STUDENTS);
  } else {
    studentsSheet = ss.insertSheet(SHEET_STUDENTS);
  }
  ensureHeaders_(studentsSheet, STUDENT_HEADERS);

  // Initialize CheckIns sheet
  var checkInsSheet = ss.insertSheet(SHEET_CHECKINS);
  ensureHeaders_(checkInsSheet, CHECKIN_HEADERS);

  // Initialize CoTeachers sheet with caseload manager entry
  var ctSheet = ss.insertSheet(SHEET_COTEACHERS);
  ensureHeaders_(ctSheet, COTEACHER_HEADERS);
  ctSheet.appendRow([email, 'caseload-manager', new Date().toISOString()]);

  // Initialize Evaluations sheet
  var evalSheet = ss.insertSheet(SHEET_EVALUATIONS);
  ensureHeaders_(evalSheet, EVALUATION_HEADERS);

  // Store the new spreadsheet ID in UserProperties
  var ssId = ss.getId();
  props.setProperty('spreadsheet_id', ssId);

  return {
    success: true,
    spreadsheetId: ssId,
    spreadsheetUrl: ss.getUrl()
  };
}

// ───── Web App Entry ─────
function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('Richfield Public Schools | Caseload Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ───── Initialization ─────

var STUDENT_HEADERS = [
  'id','firstName','lastName','grade','period',
  'focusGoal','accommodations','notes','classesJson',
  'createdAt','updatedAt','iepGoal','goalsJson','caseManagerEmail',
  'online','contactsJson','birthday','accommodationsJson'
];
var CHECKIN_HEADERS = [
  'id','studentId','weekOf',
  'planningRating','followThroughRating','regulationRating',
  'focusGoalRating','effortRating',
  'whatWentWell','barrier',
  'microGoal','microGoalCategory',
  'teacherNotes','academicDataJson','createdAt','goalMet'
];
var COTEACHER_HEADERS = ['email', 'role', 'addedAt', 'permissionsJson'];
var EVALUATION_HEADERS = ['id', 'studentId', 'type', 'itemsJson', 'createdAt', 'updatedAt', 'filesJson', 'meetingDate'];
var VALID_EVAL_TYPES = ['annual-iep', '3-year-reeval', 'initial-eval', 'eval', 'reeval'];
var EVAL_INITIAL_TYPES_ = ['initial-eval', 'eval'];

var PROGRESS_HEADERS = [
  'id', 'studentId', 'goalId', 'objectiveId', 'quarter',
  'progressRating', 'anecdotalNotes', 'dateEntered',
  'enteredBy', 'createdAt', 'lastModified'
];
var VALID_PROGRESS_RATINGS = ['No Progress', 'Adequate Progress', 'Objective Met'];
var VALID_QUARTERS = ['Q1', 'Q2', 'Q3', 'Q4'];

var IEP_MEETING_HEADERS = ['id', 'studentId', 'meetingDate', 'meetingType', 'notes', 'createdAt', 'createdBy'];
var VALID_MEETING_TYPES = ['Annual Review', 'Amendment', 'Progress Conference', 'Other'];
var CONTACT_EMAIL_RE_ = /^[^\s@]+@[^\s@]+\.[a-zA-Z]{2,}$/;

/**
 * Sanitize contacts array: require name, validate email format, normalize phone,
 * enforce single primary. Returns cleaned array (drops entries with empty name).
 */
function sanitizeContacts_(contacts) {
  if (!Array.isArray(contacts)) return [];
  var hasPrimary = false;
  return contacts.map(function(c) {
    var name = String(c.name || '').trim();
    if (!name) return null;
    var email = String(c.email || '').trim();
    if (email && !CONTACT_EMAIL_RE_.test(email)) email = '';
    var phone = String(c.phone || '').replace(/\D/g, '');
    if (phone.length !== 10) phone = '';
    var primary = !!c.primary;
    if (primary && hasPrimary) primary = false;
    if (primary) hasPrimary = true;
    return {
      name: name,
      relationship: String(c.relationship || '').trim(),
      email: email,
      phone: phone,
      primary: primary
    };
  }).filter(function(c) { return c !== null; });
}

function initializeSheets() {
  const ss = getSS_();

  let studentsSheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!studentsSheet) {
    studentsSheet = ss.insertSheet(SHEET_STUDENTS);
  }
  ensureHeaders_(studentsSheet, STUDENT_HEADERS);
  migrateStudentColumns_(studentsSheet);

  let checkInsSheet = ss.getSheetByName(SHEET_CHECKINS);
  if (!checkInsSheet) {
    checkInsSheet = ss.insertSheet(SHEET_CHECKINS);
  }
  ensureHeaders_(checkInsSheet, CHECKIN_HEADERS);

  let ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
  if (!ctSheet) {
    ctSheet = ss.insertSheet(SHEET_COTEACHERS);
    ensureHeaders_(ctSheet, COTEACHER_HEADERS);
    // Add current user as caseload manager if sheet is brand-new
    var email = getCurrentUserEmail_();
    ctSheet.appendRow([email, 'caseload-manager', new Date().toISOString()]);
  } else {
    ensureHeaders_(ctSheet, COTEACHER_HEADERS);
  }

  let evalSheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!evalSheet) evalSheet = ss.insertSheet(SHEET_EVALUATIONS);
  ensureHeaders_(evalSheet, EVALUATION_HEADERS);

  let progressSheet = ss.getSheetByName(SHEET_PROGRESS);
  if (!progressSheet) progressSheet = ss.insertSheet(SHEET_PROGRESS);
  ensureHeaders_(progressSheet, PROGRESS_HEADERS);

  let meetingsSheet = ss.getSheetByName(SHEET_IEP_MEETINGS);
  if (!meetingsSheet) meetingsSheet = ss.insertSheet(SHEET_IEP_MEETINGS);
  ensureHeaders_(meetingsSheet, IEP_MEETING_HEADERS);

  return { success: true, feedbackLinks: getFeedbackLinks() };
}

/** Verify row 1 has the expected headers; overwrite if not. */
function ensureHeaders_(sheet, expectedHeaders) {
  var needsWrite = false;
  if (sheet.getLastRow() === 0) {
    needsWrite = true;
  } else {
    var current = sheet.getRange(1, 1, 1, expectedHeaders.length).getValues()[0];
    for (var i = 0; i < expectedHeaders.length; i++) {
      if (String(current[i]).trim() !== expectedHeaders[i]) {
        needsWrite = true;
        break;
      }
    }
  }
  if (needsWrite) {
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    sheet.getRange('1:1').setFontWeight('bold');
  }
}

/**
 * Fix column misalignment caused by merge of goalsJson and caseManagerEmail.
 * Before the merge, caseManagerEmail was at column index 12. After the merge,
 * goalsJson is at 12 and caseManagerEmail at 13. Existing data rows may still
 * have email values in the goalsJson column. This migrates them to the correct column.
 */
function migrateStudentColumns_(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var goalsIdx = -1, cmIdx = -1;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === 'goalsJson') goalsIdx = i;
    if (headers[i] === 'caseManagerEmail') cmIdx = i;
  }
  if (goalsIdx === -1 || cmIdx === -1 || goalsIdx >= cmIdx) return;

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  var data = sheet.getRange(2, 1, lastRow - 1, Math.max(goalsIdx, cmIdx) + 1).getValues();
  var changed = false;

  for (var i = 0; i < data.length; i++) {
    var goalsVal = String(data[i][goalsIdx] || '').trim();
    var cmVal = String(data[i][cmIdx] || '').trim();
    // If goalsJson column has an email-like value (not JSON) and caseManagerEmail is empty,
    // it's a misplaced caseManagerEmail from before the merge.
    if (goalsVal && goalsVal.indexOf('@') !== -1 &&
        goalsVal.charAt(0) !== '[' && goalsVal.charAt(0) !== '{' && !cmVal) {
      sheet.getRange(i + 2, cmIdx + 1).setValue(goalsVal);
      sheet.getRange(i + 2, goalsIdx + 1).setValue('');
      changed = true;
    }
  }

  if (changed) invalidateStudentCaches_();
}

// ───── Student CRUD ─────

function getStudents() {
  initializeSheetsIfNeeded_();

  var cached = getCache_('students');
  if (cached) return cached;

  const ss = getSS_();
  const sheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const students = [];

  for (let i = 1; i < data.length; i++) {
    const row = {};
    headers.forEach(function(h, idx) { row[h] = data[i][idx]; });
    if (row.birthday instanceof Date) row.birthday = formatDateValue_(row.birthday);
    if (row.createdAt instanceof Date) row.createdAt = row.createdAt.toISOString();
    if (row.updatedAt instanceof Date) row.updatedAt = row.updatedAt.toISOString();
    try { row.classes = JSON.parse(row.classesJson || '[]'); }
    catch(e) { row.classes = []; }
    try { row.goals = JSON.parse(row.goalsJson || '[]'); }
    catch(e) { row.goals = []; }
    try { row.contacts = JSON.parse(row.contactsJson || '[]'); }
    catch(e) { row.contacts = []; }
    students.push(row);
  }

  students.sort(function(a, b) {
    const cmp = String(a.lastName).localeCompare(String(b.lastName));
    return cmp !== 0 ? cmp : String(a.firstName).localeCompare(String(b.firstName));
  });

  setCache_('students', students);
  return students;
}

function saveStudent(profile) {
  initializeSheetsIfNeeded_();

  // SPED Lead guard
  var status = getUserStatus();
  if (status.role === 'sped-lead') {
    throw new Error('SPED Leads have read-only access to student data');
  }

  const ss = getSS_();
  var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
  var perms = getCallerPermissions_(ctSheet, getCurrentUserEmail_());
  requirePermission_(perms, 'editStudentInfo', 'edit student info');
  const sheet = ss.getSheetByName(SHEET_STUDENTS);
  const now = new Date().toISOString();
  const classesJson = JSON.stringify(profile.classes || []);
  const contactsJson = JSON.stringify(sanitizeContacts_(profile.contacts));

  if (profile.id) {
    var found = findRowById_(sheet, profile.id);
    if (found) {
      batchSetValues_(sheet, found.rowIndex, found.colIdx, {
        firstName: profile.firstName || '',
        lastName: profile.lastName || '',
        grade: profile.grade || '',
        period: profile.period || '',
        focusGoal: profile.focusGoal || '',
        accommodations: profile.accommodations || '',
        notes: profile.notes || '',
        classesJson: classesJson,
        iepGoal: profile.iepGoal || '',
        goalsJson: profile.goalsJson || '',
        caseManagerEmail: profile.caseManagerEmail || '',
        online: profile.online ? 'TRUE' : '',
        contactsJson: contactsJson,
        birthday: profile.birthday || '',
        accommodationsJson: profile.accommodationsJson || '',
        updatedAt: now
      });
      invalidateStudentCaches_();
      return { success: true, id: profile.id };
    }
  }

  const id = Utilities.getUuid();
  sheet.appendRow([
    id, profile.firstName||'', profile.lastName||'',
    profile.grade||'', profile.period||'',
    profile.focusGoal||'', profile.accommodations||'',
    profile.notes||'', classesJson, now, now,
    profile.iepGoal||'', profile.goalsJson||'', profile.caseManagerEmail||'',
    profile.online ? 'TRUE' : '', contactsJson, profile.birthday || '',
    profile.accommodationsJson || ''
  ]);
  invalidateStudentCaches_();
  return { success: true, id: id };
}

function saveStudentGoals(studentId, goalsJson) {
  initializeSheetsIfNeeded_();
  const ss = getSS_();
  var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
  var perms = getCallerPermissions_(ctSheet, getCurrentUserEmail_());
  requirePermission_(perms, 'editGoals', 'edit goals');
  const sheet = ss.getSheetByName(SHEET_STUDENTS);
  const now = new Date().toISOString();

  var found = findRowById_(sheet, studentId);
  if (found) {
    batchSetValues_(sheet, found.rowIndex, found.colIdx, {
      goalsJson: goalsJson || '',
      updatedAt: now
    });
    invalidateStudentCaches_();
    return { success: true };
  }
  return { success: false, error: 'Student not found' };
}

function saveStudentContacts(studentId, contactsJson) {
  initializeSheetsIfNeeded_();
  const ss = getSS_();
  var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
  var perms = getCallerPermissions_(ctSheet, getCurrentUserEmail_());
  requirePermission_(perms, 'editContacts', 'edit contacts');
  const sheet = ss.getSheetByName(SHEET_STUDENTS);
  const now = new Date().toISOString();

  var contacts = [];
  try { contacts = JSON.parse(contactsJson || '[]'); } catch(e) { contacts = []; }
  var sanitized = JSON.stringify(sanitizeContacts_(contacts));

  var found = findRowById_(sheet, studentId);
  if (found) {
    batchSetValues_(sheet, found.rowIndex, found.colIdx, {
      contactsJson: sanitized,
      updatedAt: now
    });
    invalidateStudentCaches_();
    return { success: true };
  }
  return { success: false, error: 'Student not found' };
}

function appendStudentNote(studentId, noteText) {
  if (!noteText || !String(noteText).trim()) {
    return { success: false, error: 'Note text cannot be empty' };
  }
  initializeSheetsIfNeeded_();
  const ss = getSS_();
  var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
  var perms = getCallerPermissions_(ctSheet, getCurrentUserEmail_());
  requirePermission_(perms, 'editStudentInfo', 'edit student notes');
  const sheet = ss.getSheetByName(SHEET_STUDENTS);
  var found = findRowById_(sheet, studentId);
  if (!found) return { success: false, error: 'Student not found' };

  var currentNotes = sheet.getRange(found.rowIndex, found.colIdx['notes']).getValue() || '';
  var timestamp = new Date().toLocaleString();
  var separator = currentNotes ? '\n---\n' : '';
  var newNotes = currentNotes + separator + '[' + timestamp + '] ' + noteText;

  batchSetValues_(sheet, found.rowIndex, found.colIdx, {
    notes: newNotes,
    updatedAt: new Date().toISOString()
  });
  invalidateStudentCaches_();
  return { success: true, notes: newNotes };
}

function deleteStudent(studentId) {
  initializeSheetsIfNeeded_();
  const ss = getSS_();
  const studentsSheet = ss.getSheetByName(SHEET_STUDENTS);
  var found = false;
  if (studentsSheet && studentsSheet.getLastRow() > 1) {
    const sData = studentsSheet.getDataRange().getValues();
    const sColIdx = buildColIdx_(sData[0]);
    const sIdCol = sColIdx['id'] - 1;
    for (let i = sData.length - 1; i >= 1; i--) {
      if (sData[i][sIdCol] === studentId) { studentsSheet.deleteRow(i + 1); found = true; break; }
    }
  }
  if (!found) return { success: false, error: 'Student not found' };

  deleteRowsByStudentId_(ss, SHEET_CHECKINS, studentId);
  deleteRowsByStudentId_(ss, SHEET_EVALUATIONS, studentId);
  deleteRowsByStudentId_(ss, SHEET_PROGRESS, studentId);
  deleteRowsByStudentId_(ss, SHEET_IEP_MEETINGS, studentId);

  invalidateCache_();
  return { success: true };
}

/** Delete all rows matching a studentId in a given sheet. */
function deleteRowsByStudentId_(ss, sheetName, studentId) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return;
  var data = sheet.getDataRange().getValues();
  var colIdx = buildColIdx_(data[0]);
  var col = colIdx['studentId'] - 1;
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][col] === studentId) sheet.deleteRow(i + 1);
  }
}

/** Build { id: { firstName, lastName } } lookup map from students array. */
function buildStudentNameMap_(students) {
  var map = {};
  students.forEach(function(s) {
    map[s.id] = { firstName: s.firstName, lastName: s.lastName };
  });
  return map;
}

/** Compute EF average from a check-in row's 5 rating fields. Returns null if no valid ratings. */
function computeEfAvg_(checkIn) {
  var ratings = [
    Number(checkIn.planningRating), Number(checkIn.followThroughRating),
    Number(checkIn.regulationRating), Number(checkIn.focusGoalRating),
    Number(checkIn.effortRating)
  ].filter(function(r) { return !isNaN(r) && r > 0; });
  return ratings.length > 0 ? ratings.reduce(function(a, b) { return a + b; }, 0) / ratings.length : null;
}

// ───── Check-In CRUD ─────

function saveCheckIn(data) {
  initializeSheetsIfNeeded_();

  // SPED Lead guard
  var status = getUserStatus();
  if (status.role === 'sped-lead') {
    throw new Error('SPED Leads cannot create check-ins');
  }

  var ss0 = getSS_();
  var ctSheet = ss0.getSheetByName(SHEET_COTEACHERS);
  var perms = getCallerPermissions_(ctSheet, getCurrentUserEmail_());
  requirePermission_(perms, 'createCheckins', 'create check-ins');
  var lock = LockService.getUserLock();
  lock.waitLock(10000);
  try {
    const ss = getSS_();
    const sheet = ss.getSheetByName(SHEET_CHECKINS);
    const now = new Date().toISOString();
    const academicJson = JSON.stringify(data.academicData || []);

    if (data.id) {
      var found = findRowById_(sheet, data.id);
      if (found) {
        batchSetValues_(sheet, found.rowIndex, found.colIdx, {
          weekOf: data.weekOf || '',
          planningRating: data.planningRating || '',
          followThroughRating: data.followThroughRating || '',
          regulationRating: data.regulationRating || '',
          focusGoalRating: data.focusGoalRating || '',
          effortRating: data.effortRating || '',
          whatWentWell: data.whatWentWell || '',
          barrier: data.barrier || '',
          microGoal: data.microGoal || '',
          microGoalCategory: data.microGoalCategory || '',
          teacherNotes: data.teacherNotes || '',
          academicDataJson: academicJson,
          goalMet: data.goalMet || ''
        });
        invalidateCheckInCaches_();
        return { success: true, id: data.id };
      }
    }

    const id = Utilities.getUuid();
    sheet.appendRow([
      id, data.studentId, data.weekOf||'',
      data.planningRating||'', data.followThroughRating||'',
      data.regulationRating||'', data.focusGoalRating||'',
      data.effortRating||'',
      data.whatWentWell||'', data.barrier||'',
      data.microGoal||'', data.microGoalCategory||'',
      data.teacherNotes||'', academicJson, now,
      data.goalMet||''
    ]);
    invalidateCheckInCaches_();
    return { success: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

function getCheckIns(studentId) {
  initializeSheetsIfNeeded_();
  const ss = getSS_();
  var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
  if (!checkPermission_(getCallerPermissions_(ctSheet, getCurrentUserEmail_()), 'viewCheckins')) return [];
  const sheet = ss.getSheetByName(SHEET_CHECKINS);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIdx = buildColIdx_(headers);
  const results = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][colIdx['studentId'] - 1] === studentId) {
      const row = {};
      headers.forEach(function(h, idx) { row[h] = data[i][idx]; });
      row.weekOf = formatDateValue_(row.weekOf);
      try { row.academicData = JSON.parse(row.academicDataJson || '[]'); }
      catch(e) { row.academicData = []; }
      results.push(row);
    }
  }

  results.sort(function(a, b) { return b.weekOf.localeCompare(a.weekOf); });
  return results;
}

function deleteCheckIn(checkInId) {
  initializeSheetsIfNeeded_();
  const ss = getSS_();
  var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
  var perms = getCallerPermissions_(ctSheet, getCurrentUserEmail_());
  requirePermission_(perms, 'createCheckins', 'delete check-ins');
  const sheet = ss.getSheetByName(SHEET_CHECKINS);
  if (!sheet) return { success: false };
  const data = sheet.getDataRange().getValues();
  const colIdx = buildColIdx_(data[0]);
  const idCol = colIdx['id'] - 1;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][idCol] === checkInId) { sheet.deleteRow(i + 1); invalidateCheckInCaches_(); return { success: true }; }
  }
  return { success: false };
}

function updateCheckInAcademicData(checkInId, academicData) {
  initializeSheetsIfNeeded_();
  const ss = getSS_();
  var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
  var perms = getCallerPermissions_(ctSheet, getCurrentUserEmail_());
  requirePermission_(perms, 'editAcademics', 'edit academic data');
  const sheet = ss.getSheetByName(SHEET_CHECKINS);
  if (!sheet) return { success: false };

  var found = findRowById_(sheet, checkInId);
  if (found && found.colIdx['academicDataJson']) {
    sheet.getRange(found.rowIndex, found.colIdx['academicDataJson']).setValue(JSON.stringify(academicData || []));
    invalidateCheckInCaches_();
    return { success: true };
  }
  return { success: false };
}

// ───── Dashboard / Analytics ─────

function getDashboardData() {
  var cached = getCache_('dashboard');
  if (cached) return cached;

  initializeSheetsIfNeeded_();
  const students = getStudents();
  const ss = getSS_();
  const ciSheet = ss.getSheetByName(SHEET_CHECKINS);

  // Build eval type lookup (studentId -> 'eval' or 'reeval')
  const evalTypeMap = {};
  const evalSheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (evalSheet && evalSheet.getLastRow() > 1) {
    const evalData = evalSheet.getDataRange().getValues();
    const evalHeaders = evalData[0];
    const evalColIdx = buildColIdx_(evalHeaders);
    for (let i = 1; i < evalData.length; i++) {
      const sid = evalData[i][evalColIdx['studentId'] - 1];
      const etype = evalData[i][evalColIdx['type'] - 1];
      if (sid) evalTypeMap[sid] = etype || 'eval';
    }
  }

  // Build all check-ins once
  const allCheckIns = [];
  if (ciSheet && ciSheet.getLastRow() > 1) {
    const ciData = ciSheet.getDataRange().getValues();
    const ciHeaders = ciData[0];
    for (let i = 1; i < ciData.length; i++) {
      const r = {};
      ciHeaders.forEach(function(h, idx) { r[h] = ciData[i][idx]; });
      r.weekOf = formatDateValue_(r.weekOf);
      try { r.academicData = JSON.parse(r.academicDataJson || '[]'); }
      catch(e) { r.academicData = []; }
      allCheckIns.push(r);
    }
  }

  const summary = students.map(function(s) {
    const checkIns = allCheckIns.filter(function(ci) { return ci.studentId === s.id; });
    checkIns.sort(function(a, b) { return b.weekOf.localeCompare(a.weekOf); });

    const latest = checkIns[0] || null;
    const totalCheckIns = checkIns.length;

    // EF average
    let avgRating = latest ? computeEfAvg_(latest) : null;

    // Trend
    let trend = 'none';
    if (checkIns.length >= 2 && avgRating !== null) {
      const prevAvg = computeEfAvg_(checkIns[1]);
      if (prevAvg !== null) {
        if (avgRating > prevAvg + 0.3) trend = 'up';
        else if (avgRating < prevAvg - 0.3) trend = 'down';
        else trend = 'flat';
      }
    }

    // Days since last check-in
    let daysSinceCheckIn = null;
    if (latest && latest.weekOf) {
      const lastParts = latest.weekOf.split('-');
      const lastDate = new Date(Number(lastParts[0]), Number(lastParts[1]) - 1, Number(lastParts[2]));
      const today = new Date();
      today.setHours(0,0,0,0);
      daysSinceCheckIn = Math.floor((today - lastDate) / 86400000);
    }

    // EF history (last 6 weeks, oldest first)
    const efHistory = [];
    checkIns.slice(0, 6).reverse().forEach(function(ci) {
      var avg = computeEfAvg_(ci);
      if (avg !== null) efHistory.push(avg);
    });

    // Academic from latest
    let gpa = null;
    let totalMissing = 0;
    let academicData = [];
    if (latest && latest.academicData && latest.academicData.length > 0) {
      academicData = latest.academicData;
      const gpaValues = [];
      latest.academicData.forEach(function(c) {
        if (c.grade && GPA_MAP.hasOwnProperty(c.grade)) gpaValues.push(GPA_MAP[c.grade]);
        totalMissing += Number(c.missing) || 0;
      });
      if (gpaValues.length > 0) {
        gpa = gpaValues.reduce(function(a, b) { return a + b; }, 0) / gpaValues.length;
      }
    }

    return {
      id: s.id,
      firstName: s.firstName, lastName: s.lastName,
      grade: s.grade, period: s.period,
      focusGoal: s.focusGoal,
      accommodations: s.accommodations || '',
      accommodationsJson: s.accommodationsJson || '',
      notes: s.notes || '',
      iepGoal: s.iepGoal || '',
      goalsJson: s.goalsJson || '',
      caseManagerEmail: s.caseManagerEmail || '',
      contactsJson: s.contactsJson || '',
      contacts: s.contacts || [],
      classes: s.classes || [],
      totalCheckIns: totalCheckIns,
      latestCheckInId: latest ? latest.id : null,
      latestWeek: latest ? latest.weekOf : null,
      latestMicroGoal: latest ? latest.microGoal : null,
      avgRating: avgRating,
      trend: trend,
      gpa: gpa,
      totalMissing: totalMissing,
      academicData: academicData,
      birthday: s.birthday || '',
      evalType: evalTypeMap[s.id] || null,
      daysSinceCheckIn: daysSinceCheckIn,
      efHistory: efHistory
    };
  });

  setCache_('dashboard', summary);
  return summary;
}

// ───── User Properties Cache (FERPA-compliant: per-user isolated) ─────
// Sheets = source of truth; User Properties = fast read cache.
// Pattern: read-through cache, invalidate on write.

var CACHE_PREFIX = 'cache_';
var CACHE_TTL_MS = 120000; // 2-minute TTL — keeps co-teacher views reasonably fresh
var CALENDAR_CACHE_TTL_MS = 300000; // 5-minute TTL — calendar events change less frequently

function getCache_(key) {
  try {
    var raw = PropertiesService.getUserProperties().getProperty(CACHE_PREFIX + key);
    if (raw) {
      var cached = JSON.parse(raw);
      if (cached && typeof cached === 'object' && cached.hasOwnProperty('_ts')) {
        if (Date.now() - cached._ts > CACHE_TTL_MS) return null;
        return cached._data;
      }
      return cached; // legacy format (no TTL wrapper)
    }
  } catch(e) {}
  return null;
}

function setCache_(key, data) {
  try {
    var wrapped = { _data: data, _ts: Date.now() };
    PropertiesService.getUserProperties().setProperty(CACHE_PREFIX + key, JSON.stringify(wrapped));
  } catch(e) { /* exceeds 9KB property limit — skip silently */ }
}

function invalidateCache_() {
  try {
    var props = PropertiesService.getUserProperties();
    props.deleteProperty(CACHE_PREFIX + 'students');
    props.deleteProperty(CACHE_PREFIX + 'dashboard');
    props.deleteProperty(CACHE_PREFIX + 'eval_summary');
    props.deleteProperty(CACHE_PREFIX + 'progress');
    props.deleteProperty(CACHE_PREFIX + 'calendar_meetings');
    VALID_QUARTERS.forEach(function(q) {
      props.deleteProperty(CACHE_PREFIX + 'due_process_' + q);
    });
  } catch(e) {}
}

function invalidateCalendarCache_() {
  try {
    PropertiesService.getUserProperties().deleteProperty(CACHE_PREFIX + 'calendar_meetings');
  } catch(e) {}
}

// Targeted invalidation — only clear caches affected by the write
function invalidateStudentCaches_() {
  try {
    var props = PropertiesService.getUserProperties();
    props.deleteProperty(CACHE_PREFIX + 'students');
    props.deleteProperty(CACHE_PREFIX + 'dashboard');
  } catch(e) {}
}

function invalidateCheckInCaches_() {
  try {
    var props = PropertiesService.getUserProperties();
    props.deleteProperty(CACHE_PREFIX + 'dashboard');
  } catch(e) {}
}

function invalidateEvalCaches_() {
  try {
    var props = PropertiesService.getUserProperties();
    props.deleteProperty(CACHE_PREFIX + 'eval_summary');
    props.deleteProperty(CACHE_PREFIX + 'dashboard');
    VALID_QUARTERS.forEach(function(q) {
      props.deleteProperty(CACHE_PREFIX + 'due_process_' + q);
    });
  } catch(e) {}
}

function invalidateMeetingCaches_() {
  try {
    var props = PropertiesService.getUserProperties();
    props.deleteProperty(CACHE_PREFIX + 'eval_summary');
    VALID_QUARTERS.forEach(function(q) {
      props.deleteProperty(CACHE_PREFIX + 'due_process_' + q);
    });
  } catch(e) {}
}

function invalidateProgressCaches_() {
  try {
    var props = PropertiesService.getUserProperties();
    props.deleteProperty(CACHE_PREFIX + 'progress');
    VALID_QUARTERS.forEach(function(q) {
      props.deleteProperty(CACHE_PREFIX + 'due_process_' + q);
    });
  } catch(e) {}
}

// ───── Eval Task Summary (cross-student aggregation) ─────

function getEvalTaskSummary() {
  initializeSheetsIfNeeded_();
  var _ss = getSS_();
  var _ctSheet = _ss.getSheetByName(SHEET_COTEACHERS);
  if (!checkPermission_(getCallerPermissions_(_ctSheet, getCurrentUserEmail_()), 'viewEvals')) {
    return { activeEvals: [], overdueCount: 0, dueThisWeekCount: 0, timeline: [] };
  }
  var cached = getCache_('eval_summary');
  if (cached) return cached;
  var result = getEvalTimelineExtended_(7);
  setCache_('eval_summary', result);
  return result;
}

// ───── IEP Meeting CRUD ─────

function saveIEPMeeting(data) {
  if (!data || !data.studentId || !data.meetingDate || !data.meetingType) {
    return { success: false, error: 'studentId, meetingDate, and meetingType are required' };
  }
  if (VALID_MEETING_TYPES.indexOf(data.meetingType) === -1) {
    return { success: false, error: 'Invalid meetingType. Valid: ' + VALID_MEETING_TYPES.join(', ') };
  }

  initializeSheetsIfNeeded_();
  var status = getUserStatus();
  var ss;

  if (status.role === 'sped-lead') {
    // SPED Lead writes to target CM's spreadsheet
    var targetSpreadsheetId = data.caseManagerSpreadsheetId;
    if (!targetSpreadsheetId) {
      return {success: false, error: 'caseManagerSpreadsheetId required for SPED Lead writes'};
    }

    // Verify SPED Lead has access to this caseload
    var caseloads = getSpedLeadCaseloads_(getCurrentUserEmail_());
    var hasAccess = caseloads.some(function(cm) {
      return cm.spreadsheetId === targetSpreadsheetId;
    });
    if (!hasAccess) {
      return {success: false, error: 'Access denied to this caseload'};
    }

    ss = SpreadsheetApp.openById(targetSpreadsheetId);
  } else {
    ss = getSS_();
    var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
    var perms = getCallerPermissions_(ctSheet, getCurrentUserEmail_());
    requirePermission_(perms, 'editEvals', 'manage IEP meetings');
  }

  var sheet = ss.getSheetByName(SHEET_IEP_MEETINGS);
  if (!sheet) return { success: false, error: 'IEPMeetings sheet not found' };

  var lock = LockService.getUserLock();
  try {
    lock.waitLock(10000);
  } catch(e) {
    return { success: false, error: 'Could not acquire lock' };
  }

  try {
    var now = new Date().toISOString();
    var email = getCurrentUserEmail_();

    if (data.id) {
      // Update existing
      var found = findRowById_(sheet, data.id);
      if (!found) {
        return { success: false, error: 'Meeting not found' };
      }
      var colIdx = found.colIdx;
      batchSetValues_(sheet, found.rowIndex, colIdx, {
        meetingDate: data.meetingDate,
        meetingType: data.meetingType,
        notes: data.notes || ''
      });
      invalidateMeetingCaches_();
      return { success: true, id: data.id, updated: true };
    }

    // Create new
    var id = Utilities.getUuid();
    sheet.appendRow([
      id, data.studentId, data.meetingDate, data.meetingType,
      data.notes || '', now, email
    ]);
    invalidateMeetingCaches_();
    return { success: true, id: id, updated: false };
  } finally {
    lock.releaseLock();
  }
}

function getIEPMeetings(studentId) {
  initializeSheetsIfNeeded_();
  var ss = getSS_();
  var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
  if (!checkPermission_(getCallerPermissions_(ctSheet, getCurrentUserEmail_()), 'viewEvals')) return [];
  var sheet = ss.getSheetByName(SHEET_IEP_MEETINGS);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var allData = sheet.getDataRange().getValues();
  var headers = allData[0];
  var results = [];
  for (var i = 1; i < allData.length; i++) {
    var row = {};
    headers.forEach(function(h, idx) { row[h] = allData[i][idx]; });
    if (row.meetingDate instanceof Date) row.meetingDate = formatDateValue_(row.meetingDate);
    if (row.createdAt instanceof Date) row.createdAt = row.createdAt.toISOString();
    if (!studentId || row.studentId === studentId) results.push(row);
  }
  return results;
}

function deleteIEPMeeting(meetingId) {
  if (!meetingId) return { success: false, error: 'meetingId is required' };
  initializeSheetsIfNeeded_();
  var ss = getSS_();
  var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
  var perms = getCallerPermissions_(ctSheet, getCurrentUserEmail_());
  requirePermission_(perms, 'editEvals', 'delete IEP meetings');
  var sheet = ss.getSheetByName(SHEET_IEP_MEETINGS);
  if (!sheet) return { success: false, error: 'Sheet not found' };

  var found = findRowById_(sheet, meetingId);
  if (!found) return { success: false, error: 'Meeting not found' };

  sheet.deleteRow(found.rowIndex);
  invalidateMeetingCaches_();
  return { success: true };
}

// ───── Due Process Dashboard Data ─────

function getDueProcessData(quarter) {
  if (!quarter || VALID_QUARTERS.indexOf(quarter) === -1) {
    quarter = getCurrentQuarter();
  }
  initializeSheetsIfNeeded_();
  var _ss = getSS_();
  var _ctSheet = _ss.getSheetByName(SHEET_COTEACHERS);
  if (!checkPermission_(getCallerPermissions_(_ctSheet, getCurrentUserEmail_()), 'viewDueProcess')) {
    return { error: 'Permission denied' };
  }
  var cacheKey = 'due_process_' + quarter;
  var cached = getCache_(cacheKey);
  if (cached) return cached;
  var email = (getCurrentUserEmail_() || '').toLowerCase();
  var students = getStudents();

  // 1. Eval summary with extended 30-day timeline
  var evalSummary = getEvalTimelineExtended_(30);

  // 2. Merged meetings from eval meetingDate + standalone IEPMeetings
  var meetings = getAllMeetingsForCalendar_(students);

  // 2b. Count meetings this M-F week for summary stats
  var now = new Date();
  var todayDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var dow = todayDate.getDay();
  var mondayOffset = (dow === 0) ? -6 : 1 - dow;
  var weekMonday = new Date(todayDate.getTime() + mondayOffset * 86400000);
  var weekFriday = new Date(weekMonday.getTime() + 4 * 86400000);
  var weekMondayStr = formatDateValue_(weekMonday);
  var weekFridayStr = formatDateValue_(weekFriday);
  var meetingsThisWeek = meetings.filter(function(m) {
    return m.date >= weekMondayStr && m.date <= weekFridayStr;
  });

  // 3. Progress responsibility: goals where current user is responsible
  var progressAssignments = [];
  students.forEach(function(s) {
    var goals = [];
    try { goals = JSON.parse(s.goalsJson || '[]'); } catch(e) {}
    goals.forEach(function(goal) {
      if ((goal.responsibleEmail || '').toLowerCase() === email) {
        progressAssignments.push({
          studentId: s.id,
          studentName: s.firstName + ' ' + s.lastName,
          goalId: goal.id,
          goalArea: goal.goalArea || 'General',
          goalText: goal.text || '',
          caseManagerEmail: s.caseManagerEmail || '',
          online: s.online === 'TRUE',
          birthday: s.birthday || '',
          objectiveCount: (goal.objectives || []).length,
          objectives: (goal.objectives || []).map(function(obj) {
            return { id: obj.id, text: obj.text || '' };
          })
        });
      }
    });
  });

  // 4. Completion map: for each student, check if all assigned objectives have entries
  var completionMap = buildCompletionMap_(email, students, quarter);

  var completionFlags = getDPCompletionFlags_();

  var result = {
    evalSummary: evalSummary,
    meetings: meetings,
    meetingsThisWeek: meetingsThisWeek,
    progressAssignments: progressAssignments,
    completionMap: completionMap,
    completionFlags: completionFlags,
    currentQuarter: quarter
  };

  setCache_(cacheKey, result);
  return result;
}

/**
 * Lightweight quarter-switch endpoint — returns only the completion map for the
 * requested quarter. Used by switchDPProgressQuarter to avoid re-fetching evals,
 * meetings, and progress assignments that don't change with quarter.
 */
function getDPCompletionForQuarter(quarter) {
  if (!quarter || VALID_QUARTERS.indexOf(quarter) === -1) {
    quarter = getCurrentQuarter();
  }
  initializeSheetsIfNeeded_();
  var _ss = getSS_();
  var _ctSheet = _ss.getSheetByName(SHEET_COTEACHERS);
  if (!checkPermission_(getCallerPermissions_(_ctSheet, getCurrentUserEmail_()), 'viewDueProcess')) {
    return { error: 'Permission denied' };
  }
  var email = (getCurrentUserEmail_() || '').toLowerCase();
  var students = getStudents();
  var completionMap = buildCompletionMap_(email, students, quarter);
  return { completionMap: completionMap, currentQuarter: quarter };
}

/**
 * Extended eval timeline — same as getEvalTaskSummary() but builds N days instead of 7.
 */
function getEvalTimelineExtended_(numDays) {
  initializeSheetsIfNeeded_();
  var ss = getSS_();
  var evalSheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!evalSheet || evalSheet.getLastRow() <= 1) {
    return { dueThisWeekCount: 0, overdueCount: 0, timeline: [], activeEvals: [], overdueTasks: [], dueThisWeekTasks: [] };
  }

  var students = getStudents();
  var studentMap = buildStudentNameMap_(students);

  var evalData = evalSheet.getDataRange().getValues();
  var evalHeaders = evalData[0];
  var evalColIdx = buildColIdx_(evalHeaders);

  var now = new Date();
  var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var todayStr = formatDateValue_(today);

  var dayAbbrs = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  var monthAbbrs = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  var timelineDays = [];
  for (var d = 0; d < numDays; d++) {
    var dayDate = new Date(today.getTime() + d * 86400000);
    timelineDays.push({
      date: formatDateValue_(dayDate),
      dayOfWeek: dayDate.getDay(),
      dayAbbr: dayAbbrs[dayDate.getDay()],
      dayNum: dayDate.getDate(),
      monthShort: monthAbbrs[dayDate.getMonth()],
      tasks: []
    });
  }
  var endDateStr = timelineDays[numDays - 1].date;

  // Build date-to-index map for fast lookup
  var dateToIdx = {};
  timelineDays.forEach(function(td, idx) { dateToIdx[td.date] = idx; });

  var overdueCount = 0;
  var dueThisWeekCount = 0;
  var activeEvals = [];
  var overdueTasks = [];
  var dueThisWeekTasks = [];
  var seenStudents = {};

  // Week boundary for "due this week" count (7 days from today)
  var weekEndDate = new Date(today.getTime() + 6 * 86400000);
  var weekEndStr = formatDateValue_(weekEndDate);

  for (var i = 1; i < evalData.length; i++) {
    var studentId = evalData[i][evalColIdx['studentId'] - 1];
    var itemsRaw = evalData[i][evalColIdx['itemsJson'] - 1];
    var evalType = evalData[i][evalColIdx['type'] - 1];
    var evalId = evalData[i][evalColIdx['id'] - 1];
    var items;
    try { items = JSON.parse(itemsRaw || '[]'); } catch(e) { items = []; }

    var studentInfo = studentMap[studentId] || { firstName: 'Unknown', lastName: '' };
    var studentFullName = studentInfo.firstName + ' ' + studentInfo.lastName;

    var evalDone = 0;
    var evalOverdue = 0;
    var evalNextDue = null; // earliest unchecked due date
    items.forEach(function(item) {
      if (item.checked) { evalDone++; return; }
      if (!item.dueDate) return;

      // Track earliest unchecked due date for this eval
      if (!evalNextDue || item.dueDate < evalNextDue) evalNextDue = item.dueDate;

      if (item.dueDate < todayStr) {
        overdueCount++;
        evalOverdue++;
        overdueTasks.push({
          itemId: item.id, text: item.text, dueDate: item.dueDate,
          studentId: studentId, evalId: evalId, studentName: studentFullName, evalType: evalType
        });
        return;
      }

      if (item.dueDate >= todayStr && item.dueDate <= weekEndStr) {
        dueThisWeekCount++;
        dueThisWeekTasks.push({
          itemId: item.id, text: item.text, dueDate: item.dueDate,
          studentId: studentId, evalId: evalId, studentName: studentFullName, evalType: evalType
        });
      }

      // Map to timeline day if in range
      var dayIdx = dateToIdx[item.dueDate];
      if (dayIdx !== undefined) {
        timelineDays[dayIdx].tasks.push({
          itemId: item.id, text: item.text,
          studentId: studentId, evalId: evalId, studentName: studentFullName, evalType: evalType
        });
      }
    });

    if (!seenStudents[studentId]) {
      seenStudents[studentId] = true;
      activeEvals.push({
        evalId: evalId, studentId: studentId, studentName: studentFullName,
        type: evalType, done: evalDone, total: items.length, overdueCount: evalOverdue,
        nextDueDate: evalNextDue
      });
    }
  }

  overdueTasks.sort(function(a, b) { return a.dueDate < b.dueDate ? -1 : a.dueDate > b.dueDate ? 1 : 0; });
  dueThisWeekTasks.sort(function(a, b) { return a.dueDate < b.dueDate ? -1 : a.dueDate > b.dueDate ? 1 : 0; });

  return {
    dueThisWeekCount: dueThisWeekCount,
    overdueCount: overdueCount,
    timeline: timelineDays,
    activeEvals: activeEvals,
    overdueTasks: overdueTasks,
    dueThisWeekTasks: dueThisWeekTasks
  };
}

/**
 * Merge eval meetingDate values with standalone IEPMeetings rows.
 */
/**
 * Public endpoint: returns all meetings (eval + standalone + Google Calendar) for calendar views.
 * Pass forceRefresh=true to bypass the calendar cache (e.g. from a manual refresh button).
 */
function getAllMeetings(forceRefresh) {
  initializeSheetsIfNeeded_();
  if (forceRefresh) invalidateCalendarCache_();
  var students = getStudents();
  return getAllMeetingsForCalendar_(students);
}

function getAllMeetingsForCalendar_(students) {
  var ss = getSS_();
  var meetings = [];

  // Student name lookup
  var studentMap = {};
  students.forEach(function(s) { studentMap[s.id] = s.firstName + ' ' + s.lastName; });

  // Source 1: Eval meetingDate fields
  var evalSheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (evalSheet && evalSheet.getLastRow() > 1) {
    var evalData = evalSheet.getDataRange().getValues();
    var evalHeaders = evalData[0];
    var evalColIdx = buildColIdx_(evalHeaders);

    for (var i = 1; i < evalData.length; i++) {
      var md = evalData[i][evalColIdx['meetingDate'] - 1];
      if (md) {
        var sid = evalData[i][evalColIdx['studentId'] - 1];
        var evalType = evalData[i][evalColIdx['type'] - 1];
        meetings.push({
          date: formatDateValue_(md),
          studentId: sid,
          studentName: studentMap[sid] || 'Unknown',
          meetingType: getEvalMeetingLabel_(evalType),
          meetingCategory: getEvalMeetingCategory_(evalType),
          source: 'eval'
        });
      }
    }
  }

  // Source 2: IEPMeetings sheet
  var meetSheet = ss.getSheetByName(SHEET_IEP_MEETINGS);
  if (meetSheet && meetSheet.getLastRow() > 1) {
    var meetData = meetSheet.getDataRange().getValues();
    var meetHeaders = meetData[0];
    for (var j = 1; j < meetData.length; j++) {
      var row = {};
      meetHeaders.forEach(function(h, idx) { row[h] = meetData[j][idx]; });
      meetings.push({
        id: row.id,
        date: formatDateValue_(row.meetingDate),
        studentId: row.studentId,
        studentName: studentMap[row.studentId] || 'Unknown',
        meetingType: row.meetingType || 'Other',
        meetingCategory: 'iep',
        notes: row.notes || '',
        source: 'standalone'
      });
    }
  }

  // Source 3: Google Calendar events (IEP meetings matched by student initials)
  var today = new Date();
  var endDate = new Date(today.getTime() + 30 * 24 * 60 * 60 * 1000); // 30 days out
  var calMeetings = getCalendarMeetings_(students, today, endDate);
  if (calMeetings.length > 0) {
    // Auto-populate eval meeting dates from calendar events
    autoPopulateEvalMeetingDates_(calMeetings);

    // Deduplicate: skip gcal meetings where an eval/standalone meeting
    // already exists for the same student on the same date
    var existingKey = {};
    meetings.forEach(function(m) {
      existingKey[m.studentId + '|' + m.date] = true;
    });
    calMeetings.forEach(function(cm) {
      if (!existingKey[cm.studentId + '|' + cm.date]) {
        meetings.push(cm);
      }
    });
  }

  return meetings;
}

function getEvalMeetingLabel_(evalType) {
  if (evalType === 'annual-iep') return 'Annual IEP Meeting';
  if (evalType === '3-year-reeval' || evalType === 'reeval') return 'Re-Eval Meeting';
  if (evalType === 'initial-eval' || evalType === 'eval') return 'Initial Eval Meeting';
  return 'Eval Meeting';
}

function getEvalMeetingCategory_(evalType) {
  if (evalType === 'annual-iep') return 'iep';
  return 'eval'; // initial-eval, eval, 3-year-reeval, reeval
}

// ───── Google Calendar Integration ─────

/**
 * Fetch IEP and Eval meetings from the user's default Google Calendar.
 * Searches for events matching "IEP Meeting" or "Eval Meeting" and parses:
 *   {Initials} {optional format} {IEP|Eval} Meeting
 *   e.g. "RT Virtual IEP Meeting", "JS In Person Eval Meeting"
 *
 * Returns array in the standard meeting shape:
 *   { date, studentId, studentName, meetingType, meetingCategory, source: 'gcal', calendarEventTitle }
 *
 * Uses a 5-minute cache in UserProperties to avoid repeated CalendarApp calls.
 */
function getCalendarMeetings_(students, startDate, endDate) {
  // Check cache first
  var cached = getCalendarCache_();
  if (cached) return cached;

  var results = [];
  try {
    var cal = CalendarApp.getDefaultCalendar();

    // Search for both IEP and Eval meetings (CalendarApp has no OR query)
    var iepEvents = cal.getEvents(startDate, endDate, { search: 'IEP Meeting' });
    var evalEvents = cal.getEvents(startDate, endDate, { search: 'Eval Meeting' });

    // Deduplicate by event ID
    var seenIds = {};
    var allEvents = [];
    iepEvents.concat(evalEvents).forEach(function(evt) {
      var eid = evt.getId();
      if (!seenIds[eid]) {
        seenIds[eid] = true;
        allEvents.push(evt);
      }
    });

    // Build initials → student(s) lookup
    var initialsMap = {};
    students.forEach(function(s) {
      var fn = (s.firstName || '').trim();
      var ln = (s.lastName || '').trim();
      if (!fn || !ln) return;
      var key = (fn.charAt(0) + ln.charAt(0)).toUpperCase();
      if (!initialsMap[key]) initialsMap[key] = [];
      initialsMap[key].push({ id: s.id, name: fn + ' ' + ln });
    });

    allEvents.forEach(function(evt) {
      var title = (evt.getTitle() || '').trim();
      if (!title) return;

      // Extract initials (first token) and meeting type (rest of title)
      var parts = title.split(/\s+/);
      var initials = (parts[0] || '').toUpperCase();

      // Derive meeting type from everything after initials
      // e.g. "RT Virtual IEP Meeting" → "Virtual IEP Meeting"
      var meetingType = parts.slice(1).join(' ') || 'IEP Meeting';

      // Determine category: does title contain "Eval Meeting" or "IEP Meeting"?
      var titleUpper = title.toUpperCase();
      var meetingCategory = titleUpper.indexOf('EVAL MEETING') !== -1 ? 'eval' : 'iep';

      var matched = initialsMap[initials];
      if (!matched || matched.length === 0) return; // no caseload student matched

      var eventDate = formatDateValue_(evt.getStartTime());

      matched.forEach(function(student) {
        results.push({
          date: eventDate,
          studentId: student.id,
          studentName: student.name,
          meetingType: meetingType,
          meetingCategory: meetingCategory,
          source: 'gcal',
          calendarEventTitle: title
        });
      });
    });
  } catch(e) {
    // CalendarApp may not be authorized yet — fail silently
    Logger.log('getCalendarMeetings_ error: ' + e.message);
  }

  setCalendarCache_(results);
  return results;
}

/** Read calendar meetings cache (5-min TTL). */
function getCalendarCache_() {
  try {
    var raw = PropertiesService.getUserProperties().getProperty(CACHE_PREFIX + 'calendar_meetings');
    if (raw) {
      var cached = JSON.parse(raw);
      if (cached && cached._ts && (Date.now() - cached._ts < CALENDAR_CACHE_TTL_MS)) {
        return cached._data;
      }
    }
  } catch(e) {}
  return null;
}

/** Write calendar meetings cache. */
function setCalendarCache_(data) {
  try {
    var wrapped = { _data: data, _ts: Date.now() };
    PropertiesService.getUserProperties().setProperty(
      CACHE_PREFIX + 'calendar_meetings', JSON.stringify(wrapped)
    );
  } catch(e) { /* exceeds property limit — skip silently */ }
}

/**
 * Auto-populate meeting dates on active evals from calendar events.
 * Only sets the date if the eval's meetingDate is empty and the calendar
 * event date is today or in the future.
 */
function autoPopulateEvalMeetingDates_(calendarMeetings) {
  if (!calendarMeetings || calendarMeetings.length === 0) return;

  var ss = getSS_();
  var evalSheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!evalSheet || evalSheet.getLastRow() <= 1) return;

  var evalData = evalSheet.getDataRange().getValues();
  var evalHeaders = evalData[0];
  var colIdx = buildColIdx_(evalHeaders);

  var todayStr = formatDateValue_(new Date());
  var didWrite = false;

  // Build lookup: studentId → [{rowIndex, meetingDate, evalId}]
  var evalsByStudent = {};
  for (var i = 1; i < evalData.length; i++) {
    var sid = evalData[i][colIdx['studentId'] - 1];
    var md = evalData[i][colIdx['meetingDate'] - 1];
    if (!evalsByStudent[sid]) evalsByStudent[sid] = [];
    evalsByStudent[sid].push({
      rowIndex: i + 1,
      meetingDate: md ? formatDateValue_(md) : '',
      evalId: evalData[i][colIdx['id'] - 1]
    });
  }

  calendarMeetings.forEach(function(m) {
    if (m.date < todayStr) return; // skip past events
    var evals = evalsByStudent[m.studentId];
    if (!evals) return;

    evals.forEach(function(ev) {
      if (ev.meetingDate) return; // already has a meeting date — don't overwrite
      // Set the meeting date from the calendar event
      batchSetValues_(evalSheet, ev.rowIndex, colIdx, {
        meetingDate: m.date,
        updatedAt: new Date().toISOString()
      });
      ev.meetingDate = m.date; // prevent duplicate writes for same eval
      didWrite = true;
    });
  });

  if (didWrite) {
    invalidateEvalCaches_();
  }
}

/**
 * Build completion map: for each student with goals assigned to the user,
 * count total objectives vs objectives with progress entries for the quarter.
 */
function buildCompletionMap_(email, students, quarter) {
  var ss = getSS_();
  var progressSheet = ss.getSheetByName(SHEET_PROGRESS);
  var allEntries = [];

  if (progressSheet && progressSheet.getLastRow() > 1) {
    var pData = progressSheet.getDataRange().getValues();
    var pHeaders = pData[0];
    for (var i = 1; i < pData.length; i++) {
      var entry = {};
      pHeaders.forEach(function(h, idx) { entry[h] = pData[i][idx]; });
      if (entry.quarter === quarter) allEntries.push(entry);
    }
  }

  // Build lookup set: "studentId|goalId|objectiveId" → true
  var entrySet = {};
  allEntries.forEach(function(e) {
    if (e.progressRating) {
      entrySet[e.studentId + '|' + e.goalId + '|' + e.objectiveId] = true;
    }
  });

  var completionMap = {};

  students.forEach(function(s) {
    var goals = [];
    try { goals = JSON.parse(s.goalsJson || '[]'); } catch(e) {}

    var myGoals = goals.filter(function(g) {
      return (g.responsibleEmail || '').toLowerCase() === email;
    });

    if (myGoals.length === 0) return;

    var total = 0;
    var completed = 0;
    var goalDetail = {};
    myGoals.forEach(function(goal) {
      var objectives = goal.objectives || [];
      var objDetail = {};
      if (objectives.length === 0) {
        // Goal with no objectives counts as 1 item
        total++;
        var done = !!entrySet[s.id + '|' + goal.id + '|'];
        if (done) completed++;
        goalDetail[goal.id] = { completed: done, objectives: {} };
      } else {
        var goalAllDone = true;
        objectives.forEach(function(obj) {
          total++;
          var objDone = !!entrySet[s.id + '|' + goal.id + '|' + obj.id];
          if (objDone) completed++;
          else goalAllDone = false;
          objDetail[obj.id] = objDone;
        });
        goalDetail[goal.id] = { completed: goalAllDone, objectives: objDetail };
      }
    });

    completionMap[s.id] = {
      total: total,
      completed: completed,
      allDone: total > 0 && completed >= total,
      goals: goalDetail
    };
  });

  return completionMap;
}

// ───── Teacher Feedback Links (per-user) ─────

function getFeedbackLinks() {
  var props = PropertiesService.getUserProperties();
  return {
    formUrl: props.getProperty('feedback_form_url') || '',
    sheetUrl: props.getProperty('feedback_sheet_url') || ''
  };
}

function saveFeedbackLinks(links) {
  var props = PropertiesService.getUserProperties();
  props.setProperty('feedback_form_url', links.formUrl || '');
  props.setProperty('feedback_sheet_url', links.sheetUrl || '');
  return { success: true };
}

// ───── Dashboard Config (per-user) ─────

/**
 * Retrieve the user's dashboard config (row order + custom widgets).
 * Returns default config if none saved.
 */
function getDashboardConfig() {
  var props = PropertiesService.getUserProperties();
  var json = props.getProperty('dashboard_config');
  if (!json) {
    return getDefaultDashboardConfig_();
  }
  try {
    var config = JSON.parse(json);
    // Ensure required keys exist
    if (!Array.isArray(config.rowOrder)) config.rowOrder = getDefaultDashboardConfig_().rowOrder;
    if (!Array.isArray(config.widgets)) config.widgets = [];
    return config;
  } catch (e) {
    return getDefaultDashboardConfig_();
  }
}

/**
 * Save the user's dashboard config (row order + custom widgets).
 * Stored in UserProperties as JSON — no cache invalidation needed
 * since this doesn't affect student/eval data.
 */
function saveDashboardConfig(config) {
  if (!config || typeof config !== 'object') {
    throw new Error('Invalid config');
  }
  if (!Array.isArray(config.rowOrder)) {
    throw new Error('rowOrder must be an array');
  }
  if (!Array.isArray(config.widgets)) {
    throw new Error('widgets must be an array');
  }
  // Cap widgets at 10 to prevent abuse of UserProperties storage
  if (config.widgets.length > 10) {
    throw new Error('Maximum 10 widgets allowed');
  }
  var props = PropertiesService.getUserProperties();
  props.setProperty('dashboard_config', JSON.stringify(config));
  return { success: true };
}

/** Default config: built-in rows in standard order, no widgets. */
function getDefaultDashboardConfig_() {
  return {
    rowOrder: ['needs-attention', 'at-a-glance', 'evals', 'missing', 'recent'],
    widgets: []
  };
}

// ───── Co-Teacher Management ─────

/** Return team members and the current user's role for the active spreadsheet. */
function getTeamInfo() {
  var email = getCurrentUserEmail_();
  var ss = getSS_();
  var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);

  if (!ctSheet || ctSheet.getLastRow() <= 1) {
    return { members: [{ email: email, role: 'caseload-manager', addedAt: '' }], currentUserRole: 'caseload-manager' };
  }

  var data = ctSheet.getDataRange().getValues();
  var headers = data[0];
  var members = [];
  var currentUserRole = 'caseload-manager';

  for (var i = 1; i < data.length; i++) {
    var member = {};
    headers.forEach(function(h, idx) { member[h] = data[i][idx]; });
    if (member.addedAt instanceof Date) member.addedAt = member.addedAt.toISOString();
    member.role = normalizeRole_(member.role);
    // Parse permissions for display in team management UI
    if (member.permissionsJson) {
      try { member.permissions = JSON.parse(member.permissionsJson); } catch(e) { member.permissions = null; }
    }
    delete member.permissionsJson; // Don't send raw JSON string
    // If no stored permissions, resolve from role defaults
    if (!member.permissions && member.role !== 'caseload-manager') {
      member.permissions = resolveDefaultPermissions_(member.role);
    }
    members.push(member);
    if (String(member.email).toLowerCase() === email) {
      currentUserRole = member.role;
    }
  }

  var found = members.some(function(m) { return String(m.email).toLowerCase() === email; });
  if (!found) {
    members.unshift({ email: email, role: 'caseload-manager', addedAt: '' });
    currentUserRole = 'caseload-manager';
  }

  return {
    members: members,
    currentUserRole: currentUserRole,
    roleDefaults: ROLE_DEFAULT_PERMISSIONS,
    permissionKeys: PERMISSION_KEYS
  };
}

/** Valid roles that a caseload manager can assign to team members. */
var ASSIGNABLE_ROLES = ['service-provider', 'para', 'co-teacher'];

/** Invite a team member by email with a specified role. Only caseload managers can add members. */
function addTeamMember(email, role, customPermissions) {
  email = String(email || '').trim().toLowerCase();
  if (!email || email.indexOf('@') === -1) {
    return { success: false, error: 'Please enter a valid email address.' };
  }

  role = String(role || '').trim().toLowerCase();
  if (ASSIGNABLE_ROLES.indexOf(role) === -1) {
    return { success: false, error: 'Invalid role. Must be one of: Service Provider, Para, or Co-Teacher.' };
  }

  var currentEmail = getCurrentUserEmail_();
  if (email === currentEmail) {
    return { success: false, error: 'You cannot add yourself as a team member.' };
  }

  var ss = getSS_();

  // Ensure CoTeachers sheet exists
  var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
  if (!ctSheet) {
    ctSheet = ss.insertSheet(SHEET_COTEACHERS);
    ensureHeaders_(ctSheet, COTEACHER_HEADERS);
    ctSheet.appendRow([currentEmail, 'caseload-manager', new Date().toISOString()]);
  }

  // Enforce: only caseload managers can add team members
  var callerRole = getCallerRole_(ctSheet, currentEmail);
  if (callerRole !== 'caseload-manager') {
    return { success: false, error: 'Only Caseload Managers can add team members.' };
  }

  // Check for duplicates
  var data = ctSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === email) {
      return { success: false, error: 'This person is already on your team.' };
    }
  }

  // Resolve permissions: start with role defaults, apply any custom overrides
  var perms = resolveDefaultPermissions_(role);
  if (customPermissions && typeof customPermissions === 'object') {
    PERMISSION_KEYS.forEach(function(k) {
      if (customPermissions.hasOwnProperty(k)) {
        perms[k] = !!customPermissions[k];
      }
    });
  }

  // Add to CoTeachers sheet with the specified role and permissions
  ctSheet.appendRow([email, role, new Date().toISOString(), JSON.stringify(perms)]);

  // Share the spreadsheet with the new team member
  try {
    ss.addEditor(email);
  } catch(e) {
    return { success: false, error: 'Could not share the spreadsheet: ' + e.message };
  }

  // Store invite in ScriptProperties so the member sees it on next load
  try {
    var scriptProps = PropertiesService.getScriptProperties();
    scriptProps.setProperty('coteacher_invite_' + email, JSON.stringify({
      spreadsheetId: ss.getId(),
      inviterEmail: currentEmail,
      invitedAt: new Date().toISOString()
    }));
  } catch(e) {}

  return { success: true };
}

/** Look up a user's role from the CoTeachers sheet. */
function getCallerRole_(ctSheet, email) {
  if (!ctSheet || ctSheet.getLastRow() <= 1) return 'caseload-manager';
  var data = ctSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === email) {
      return normalizeRole_(data[i][1]);
    }
  }
  return 'caseload-manager'; // fallback for spreadsheet creator
}

/** Compute default permissions from role name. */
function resolveDefaultPermissions_(role) {
  var defaults = ROLE_DEFAULT_PERMISSIONS[role];
  if (!defaults) {
    // Unknown role gets view-only permissions
    var viewOnly = {};
    PERMISSION_KEYS.forEach(function(k) { viewOnly[k] = k.indexOf('view') === 0; });
    return viewOnly;
  }
  var result = {};
  PERMISSION_KEYS.forEach(function(k) { result[k] = !!defaults[k]; });
  return result;
}

/** Resolve the full permission set for a caller from the CoTeachers sheet. */
function getCallerPermissions_(ctSheet, email) {
  var role = getCallerRole_(ctSheet, email);

  // Caseload managers always get all permissions
  if (role === 'caseload-manager') {
    var allTrue = {};
    PERMISSION_KEYS.forEach(function(k) { allTrue[k] = true; });
    return allTrue;
  }

  // Look up stored permissions from the sheet
  if (ctSheet && ctSheet.getLastRow() > 1) {
    var data = ctSheet.getDataRange().getValues();
    var headers = data[0];
    var colIdx = buildColIdx_(headers);
    var permCol = colIdx['permissionsJson'];

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === email) {
        var permJson = permCol ? String(data[i][permCol - 1] || '') : '';
        if (permJson) {
          try {
            var stored = JSON.parse(permJson);
            var result = {};
            PERMISSION_KEYS.forEach(function(k) {
              result[k] = stored.hasOwnProperty(k) ? !!stored[k] : false;
            });
            return result;
          } catch(e) { /* fall through to defaults */ }
        }
        break;
      }
    }
  }

  // Fallback: compute from role defaults
  return resolveDefaultPermissions_(role);
}

/** Check a single permission key. Returns true if granted. */
function checkPermission_(permissions, key) {
  return permissions && permissions[key] === true;
}

/** Return error result if permission is denied. For use in public endpoints. */
function requirePermission_(permissions, key, actionLabel) {
  if (!checkPermission_(permissions, key)) {
    throw new Error('Permission denied: you do not have access to ' + (actionLabel || key) + '.');
  }
}

/** Remove a team member. Only caseload managers can do this. Revokes spreadsheet access. */
function removeTeamMember(email) {
  email = String(email || '').trim().toLowerCase();
  var currentEmail = getCurrentUserEmail_();
  var ss = getSS_();
  var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
  if (!ctSheet) return { success: false, error: 'No team members configured.' };

  // Enforce: only caseload managers can remove team members
  var callerRole = getCallerRole_(ctSheet, currentEmail);
  if (callerRole !== 'caseload-manager') {
    return { success: false, error: 'Only Caseload Managers can remove team members.' };
  }

  var data = ctSheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    var memberRole = String(data[i][1]).toLowerCase();
    if (String(data[i][0]).toLowerCase() === email && memberRole !== 'caseload-manager' && memberRole !== 'owner') {
      ctSheet.deleteRow(i + 1);
      break;
    }
  }

  try { ss.removeEditor(email); } catch(e) {}

  // Remove any pending invite
  try {
    PropertiesService.getScriptProperties().deleteProperty('coteacher_invite_' + email);
  } catch(e) {}

  return { success: true };
}

/** Update permissions for an existing team member. Only caseload managers can do this. */
function updateTeamMemberPermissions(memberEmail, permissionsObj) {
  memberEmail = String(memberEmail || '').trim().toLowerCase();
  if (!memberEmail) return { success: false, error: 'Email is required.' };

  var currentEmail = getCurrentUserEmail_();
  var ss = getSS_();
  var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
  if (!ctSheet) return { success: false, error: 'No team configured.' };

  var callerRole = getCallerRole_(ctSheet, currentEmail);
  if (callerRole !== 'caseload-manager') {
    return { success: false, error: 'Only Caseload Managers can edit permissions.' };
  }

  var data = ctSheet.getDataRange().getValues();
  var colIdx = buildColIdx_(data[0]);
  var permCol = colIdx['permissionsJson'];
  if (!permCol) return { success: false, error: 'Sheet schema needs update. Please reload.' };

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === memberEmail) {
      var memberRole = normalizeRole_(data[i][colIdx['role'] - 1]);
      if (memberRole === 'caseload-manager') {
        return { success: false, error: 'Cannot modify caseload manager permissions.' };
      }

      // Validate and sanitize the permissions object
      var sanitized = {};
      PERMISSION_KEYS.forEach(function(k) {
        sanitized[k] = permissionsObj && permissionsObj.hasOwnProperty(k) ? !!permissionsObj[k] : false;
      });

      batchSetValues_(ctSheet, i + 1, colIdx, {
        permissionsJson: JSON.stringify(sanitized)
      });

      return { success: true };
    }
  }

  return { success: false, error: 'Team member not found.' };
}

/** Co-teacher accepts a pending invite and links to the shared spreadsheet. */
function acceptCoTeacherInvite() {
  var email = getCurrentUserEmail_();
  var scriptProps = PropertiesService.getScriptProperties();
  var inviteRaw = scriptProps.getProperty('coteacher_invite_' + email);

  if (!inviteRaw) {
    return { success: false, error: 'No pending invite found.' };
  }

  var invite = JSON.parse(inviteRaw);

  // Verify access
  try {
    var ss = SpreadsheetApp.openById(invite.spreadsheetId);
  } catch(e) {
    return { success: false, error: 'Cannot access the shared spreadsheet. The Caseload Manager may have revoked access.' };
  }

  var props = PropertiesService.getUserProperties();
  var oldSsId = props.getProperty('spreadsheet_id');
  if (oldSsId) {
    props.setProperty('own_spreadsheet_id', oldSsId);
  }

  props.setProperty('spreadsheet_id', invite.spreadsheetId);
  scriptProps.deleteProperty('coteacher_invite_' + email);
  invalidateCache_();

  return { success: true, spreadsheetUrl: ss.getUrl() };
}

/** Co-teacher declines a pending invite. */
function declineCoTeacherInvite() {
  var email = getCurrentUserEmail_();
  PropertiesService.getScriptProperties().deleteProperty('coteacher_invite_' + email);
  return { success: true };
}

/** Co-teacher voluntarily leaves the shared spreadsheet. */
function leaveCoTeacherTeam() {
  var email = getCurrentUserEmail_();
  var props = PropertiesService.getUserProperties();

  // Remove self from CoTeachers sheet
  try {
    var ss = getSS_();
    var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
    if (ctSheet && ctSheet.getLastRow() > 1) {
      var data = ctSheet.getDataRange().getValues();
      for (var i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]).toLowerCase() === email) {
          ctSheet.deleteRow(i + 1);
          break;
        }
      }
    }
  } catch(e) {}

  // Clear shared spreadsheet and restore own if available
  props.deleteProperty('spreadsheet_id');
  var ownSsId = props.getProperty('own_spreadsheet_id');
  if (ownSsId) {
    try {
      SpreadsheetApp.openById(ownSsId);
      props.setProperty('spreadsheet_id', ownSsId);
      props.deleteProperty('own_spreadsheet_id');
    } catch(e) {
      props.deleteProperty('own_spreadsheet_id');
    }
  }

  invalidateCache_();
  return { success: true };
}

/** Force-refresh: invalidate cache and return fresh dashboard data. */

// ───── Evaluation Checklist CRUD ─────

function getEvaluation(studentId) {
  initializeSheetsIfNeeded_();
  var ss = getSS_();
  var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
  if (!checkPermission_(getCallerPermissions_(ctSheet, getCurrentUserEmail_()), 'viewEvals')) return null;
  var sheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!sheet || sheet.getLastRow() <= 1) return null;

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colIdx = buildColIdx_(headers);
  var studentIdCol = colIdx['studentId'] - 1;

  for (var i = 1; i < data.length; i++) {
    if (data[i][studentIdCol] === studentId) {
      var row = {};
      headers.forEach(function(h, idx) { row[h] = data[i][idx]; });
      // Convert Date objects to strings for google.script.run serialization
      // (Sheets auto-formats date-like strings as native Date objects)
      if (row.meetingDate instanceof Date) row.meetingDate = formatDateValue_(row.meetingDate);
      if (row.createdAt instanceof Date) row.createdAt = row.createdAt.toISOString();
      if (row.updatedAt instanceof Date) row.updatedAt = row.updatedAt.toISOString();
      try { row.items = JSON.parse(row.itemsJson || '[]'); }
      catch(e) { row.items = []; }
      try { row.files = JSON.parse(row.filesJson || '[]'); }
      catch(e) { row.files = []; }
      return row;
    }
  }
  return null;
}

/** Look up an evaluation by its eval ID (not studentId). Returns parsed eval object or null. */
function getEvaluationById_(evalId) {
  initializeSheetsIfNeeded_();
  var ss = getSS_();
  var sheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!sheet || sheet.getLastRow() <= 1) return null;

  var found = findRowById_(sheet, evalId);
  if (!found) return null;

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rowData = data[found.rowIndex - 1];
  var row = {};
  headers.forEach(function(h, idx) { row[h] = rowData[idx]; });

  if (row.meetingDate instanceof Date) row.meetingDate = formatDateValue_(row.meetingDate);
  if (row.createdAt instanceof Date) row.createdAt = row.createdAt.toISOString();
  if (row.updatedAt instanceof Date) row.updatedAt = row.updatedAt.toISOString();
  try { row.items = JSON.parse(row.itemsJson || '[]'); } catch(e) { row.items = []; }
  try { row.files = JSON.parse(row.filesJson || '[]'); } catch(e) { row.files = []; }
  return row;
}

function createEvaluation(studentId, type) {
  initializeSheetsIfNeeded_();

  // SPED Lead guard
  var status = getUserStatus();
  if (status.role === 'sped-lead') {
    throw new Error('SPED Leads have read-only access to evaluations');
  }

  var _ss = getSS_();
  var _ctSheet = _ss.getSheetByName(SHEET_COTEACHERS);
  var _perms = getCallerPermissions_(_ctSheet, getCurrentUserEmail_());
  requirePermission_(_perms, 'editEvals', 'create evaluations');
  if (VALID_EVAL_TYPES.indexOf(type) === -1) {
    return { success: false, error: 'Invalid evaluation type.' };
  }

  var lock = LockService.getUserLock();
  try {
    lock.waitLock(10000);
  } catch(e) {
    return { success: false, error: 'Could not acquire lock. Please try again.' };
  }

  try {
    var existing = getEvaluation(studentId);
    if (existing) {
      return { success: false, error: 'This student already has an active checklist.' };
    }

    var ss = getSS_();
    var sheet = ss.getSheetByName(SHEET_EVALUATIONS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_EVALUATIONS);
      ensureHeaders_(sheet, EVALUATION_HEADERS);
    }

    var items = EVAL_INITIAL_TYPES_.indexOf(type) !== -1 ? getEvalTemplateItems_() : getReEvalTemplateItems_();
    var now = new Date().toISOString();
    var id = Utilities.getUuid();

    sheet.appendRow([id, studentId, type, JSON.stringify(items), now, now, JSON.stringify([]), '']);
    invalidateEvalCaches_();

    return { success: true, id: id, studentId: studentId, type: type, items: items, files: [], meetingDate: '' };
  } finally {
    lock.releaseLock();
  }
}

function updateEvalMeetingDate(evalId, meetingDate) {
  initializeSheetsIfNeeded_();
  var ss = getSS_();
  var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
  var perms = getCallerPermissions_(ctSheet, getCurrentUserEmail_());
  requirePermission_(perms, 'editEvals', 'update eval meeting dates');
  var sheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!sheet) return { success: false, error: 'Evaluations sheet not found.' };

  var found = findRowById_(sheet, evalId);
  if (!found) return { success: false, error: 'Evaluation not found.' };

  batchSetValues_(sheet, found.rowIndex, found.colIdx, {
    meetingDate: meetingDate || '',
    updatedAt: new Date().toISOString()
  });
  invalidateEvalCaches_();
  return { success: true };
}

function updateEvaluationType(evalId, newType) {
  if (VALID_EVAL_TYPES.indexOf(newType) === -1) {
    return { success: false, error: 'Invalid evaluation type.' };
  }
  initializeSheetsIfNeeded_();
  var ss = getSS_();
  var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
  var perms = getCallerPermissions_(ctSheet, getCurrentUserEmail_());
  requirePermission_(perms, 'editEvals', 'change eval types');
  var sheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!sheet) return { success: false, error: 'Evaluations sheet not found.' };

  var found = findRowById_(sheet, evalId);
  if (!found) return { success: false, error: 'Evaluation not found.' };

  batchSetValues_(sheet, found.rowIndex, found.colIdx, {
    type: newType,
    updatedAt: new Date().toISOString()
  });
  invalidateEvalCaches_();
  return { success: true };
}

// ─── Primary items-save endpoint (replaces granular item CRUD) ───

function saveEvaluationItems(evalId, items) {
  initializeSheetsIfNeeded_();

  // SPED Lead guard
  var status = getUserStatus();
  if (status.role === 'sped-lead') {
    throw new Error('SPED Leads have read-only access to evaluations');
  }

  var ss = getSS_();
  var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
  var perms = getCallerPermissions_(ctSheet, getCurrentUserEmail_());
  requirePermission_(perms, 'editEvals', 'edit eval checklist items');
  var sheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!sheet) return { success: false, error: 'Evaluations sheet not found.' };

  var found = findRowById_(sheet, evalId);
  if (!found) return { success: false, error: 'Evaluation not found.' };

  var sanitized = (items || []).map(function(item) {
    var files = [];
    if (Array.isArray(item.files)) {
      files = item.files.map(function(f) {
        return { id: String(f.id || ''), name: String(f.name || '').trim(), url: String(f.url || '').trim() };
      }).filter(function(f) { return f.name && f.url; });
    }
    return {
      id: String(item.id || ''),
      text: String(item.text || '').trim(),
      checked: !!item.checked,
      completedAt: item.completedAt || null,
      dueDate: item.dueDate || null,
      files: files
    };
  });

  batchSetValues_(sheet, found.rowIndex, found.colIdx, {
    itemsJson: JSON.stringify(sanitized),
    updatedAt: new Date().toISOString()
  });
  invalidateEvalCaches_();

  return { success: true, items: sanitized };
}


function deleteEvaluation(evalId) {
  initializeSheetsIfNeeded_();

  // SPED Lead guard
  var status = getUserStatus();
  if (status.role === 'sped-lead') {
    throw new Error('SPED Leads have read-only access to evaluations');
  }

  var ss = getSS_();
  var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
  var perms = getCallerPermissions_(ctSheet, getCurrentUserEmail_());
  requirePermission_(perms, 'editEvals', 'delete evaluations');
  var sheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!sheet) return { success: false };

  var found = findRowById_(sheet, evalId);
  if (found) {
    sheet.deleteRow(found.rowIndex);
    invalidateEvalCaches_();
    return { success: true };
  }
  return { success: false };
}


// ───── Drive File Browser & Eval Files ─────

/** Trigger Drive scope — never called, but ensures the scope is added. */
function triggerDriveScope_() { DriveApp.getRootFolder(); }

/** Trigger DocumentApp scope — never called, ensures the scope is added. */
function triggerDocScope_() { DocumentApp.create(''); }

/**
 * Search the user's Google Drive for files matching a query string.
 * Returns up to 20 results with id, name, mimeType, url, and iconUrl.
 */
function searchDriveFiles(query) {
  var results = [];
  var maxResults = 20;
  try {
    // Use Drive Advanced Service (v2) for ordering support
    var params = {
      maxResults: maxResults,
      orderBy: 'modifiedDate desc',
      q: 'trashed = false'
    };
    if (query && query.trim()) {
      params.q = 'title contains \'' + query.replace(/'/g, "\\'") + '\' and trashed = false';
    }
    var resp = Drive.Files.list(params);
    var items = resp.items || [];
    for (var i = 0; i < items.length; i++) {
      var f = items[i];
      results.push({
        id: f.id,
        name: f.title,
        mimeType: f.mimeType,
        url: f.alternateLink || ''
      });
    }
  } catch (e) {
    // Fallback to DriveApp if Advanced Service not enabled
    try {
      var files;
      if (query && query.trim()) {
        files = DriveApp.searchFiles('title contains \'' + query.replace(/'/g, "\\'") + '\' and trashed = false');
      } else {
        files = DriveApp.searchFiles('trashed = false');
      }
      while (files.hasNext() && results.length < maxResults) {
        var file = files.next();
        results.push({
          id: file.getId(),
          name: file.getName(),
          mimeType: file.getMimeType(),
          url: file.getUrl()
        });
      }
    } catch (e2) {
      Logger.log('searchDriveFiles fallback error: ' + e2.message);
    }
  }
  return results;
}

function addEvalFile(evalId, fileData) {
  initializeSheetsIfNeeded_();
  var ss = getSS_();
  var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
  var perms = getCallerPermissions_(ctSheet, getCurrentUserEmail_());
  requirePermission_(perms, 'editEvals', 'add eval files');
  var sheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!sheet) return { success: false, error: 'Evaluations sheet not found.' };

  var found = findRowById_(sheet, evalId);
  if (!found) return { success: false, error: 'Evaluation not found.' };

  var filesCol = found.colIdx['filesJson'];
  var filesRaw = filesCol ? sheet.getRange(found.rowIndex, filesCol).getValue() : '';
  var files;
  try { files = JSON.parse(filesRaw || '[]'); } catch(e) { files = []; }

  var newFile = {
    id: 'file-' + Utilities.getUuid().substr(0, 8),
    driveFileId: fileData.driveFileId || '',
    name: fileData.name || 'Untitled',
    mimeType: fileData.mimeType || '',
    url: fileData.url || '',
    addedAt: new Date().toISOString()
  };

  files.push(newFile);

  batchSetValues_(sheet, found.rowIndex, found.colIdx, {
    filesJson: JSON.stringify(files),
    updatedAt: new Date().toISOString()
  });

  return { success: true, files: files, newFile: newFile };
}

function removeEvalFile(evalId, fileId) {
  initializeSheetsIfNeeded_();
  var ss = getSS_();
  var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
  var perms = getCallerPermissions_(ctSheet, getCurrentUserEmail_());
  requirePermission_(perms, 'editEvals', 'remove eval files');
  var sheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!sheet) return { success: false, error: 'Evaluations sheet not found.' };

  var found = findRowById_(sheet, evalId);
  if (!found) return { success: false, error: 'Evaluation not found.' };

  var filesCol = found.colIdx['filesJson'];
  var filesRaw = filesCol ? sheet.getRange(found.rowIndex, filesCol).getValue() : '';
  var files;
  try { files = JSON.parse(filesRaw || '[]'); } catch(e) { files = []; }

  files = files.filter(function(f) { return f.id !== fileId; });

  batchSetValues_(sheet, found.rowIndex, found.colIdx, {
    filesJson: JSON.stringify(files),
    updatedAt: new Date().toISOString()
  });

  return { success: true, files: files };
}

// ───── Eval Template Items ─────

function getEvalTemplateItems_() {
  var items = [
    'Review referral and obtain parent consent',
    'Conduct record review',
    'Conduct classroom observations',
    'Administer academic assessments',
    'Administer cognitive assessments',
    'Administer behavioral/social-emotional assessments',
    'Gather teacher input',
    'Gather parent input',
    'Write evaluation report',
    'Schedule eligibility determination meeting',
    'Hold eligibility determination meeting',
    'Finalize evaluation documentation'
  ];
  return items.map(function(text, idx) {
    return { id: 'item-' + (idx + 1), text: text, checked: false, completedAt: null, dueDate: null };
  });
}

function getReEvalTemplateItems_() {
  var items = [
    'Review existing evaluation data',
    'Send parent notification and obtain consent',
    'Gather teacher input on current performance',
    'Gather parent input on current concerns',
    'Review progress monitoring data',
    'Determine if additional assessments are needed',
    'Conduct additional assessments (if needed)',
    'Write re-evaluation report',
    'Schedule re-evaluation meeting',
    'Hold re-evaluation meeting',
    'Finalize re-evaluation documentation'
  ];
  return items.map(function(text, idx) {
    return { id: 'item-' + (idx + 1), text: text, checked: false, completedAt: null, dueDate: null };
  });
}

// ───── Progress Reporting ─────

/** Determine the current school-year quarter based on today's date.
 *  Q1 = Sep–Nov, Q2 = Dec–Feb, Q3 = Mar–May, Q4 = Jun–Aug */
function getCurrentQuarter() {
  var month = new Date().getMonth(); // 0-indexed
  if (month >= 8 && month <= 10) return 'Q1';   // Sep-Nov
  if (month === 11 || month <= 1) return 'Q2';  // Dec-Feb
  if (month >= 2 && month <= 4) return 'Q3';    // Mar-May
  return 'Q4';                                   // Jun-Aug
}

/** Human-readable quarter label with season and approximate school year. */
function getQuarterLabel_(quarter) {
  var now = new Date();
  var year = now.getFullYear();
  var month = now.getMonth();
  // School year straddles calendar years: 2025-26
  var startYear = month >= 8 ? year : year - 1;
  var endYear = startYear + 1;
  var yearStr = String(startYear) + '-' + String(endYear).slice(-2);

  var seasons = { Q1: 'Fall', Q2: 'Winter', Q3: 'Spring', Q4: 'Summer' };
  var season = seasons[quarter] || '';
  return quarter + ' \u2014 ' + season + ' ' + yearStr;
}

/** Save or update a progress entry for a specific goal+objective+quarter. */
function saveProgressEntry(data) {
  // Validate required fields
  if (!data || !data.studentId) return { success: false, error: 'studentId is required.' };
  if (!data.goalId) return { success: false, error: 'goalId is required.' };
  if (!data.objectiveId) return { success: false, error: 'objectiveId is required.' };

  // Validate quarter
  var quarter = String(data.quarter || '');
  if (VALID_QUARTERS.indexOf(quarter) === -1) {
    return { success: false, error: 'Invalid quarter. Must be one of: ' + VALID_QUARTERS.join(', ') };
  }

  // Validate progressRating
  var rating = String(data.progressRating || '');
  if (VALID_PROGRESS_RATINGS.indexOf(rating) === -1) {
    return { success: false, error: 'Invalid progress rating. Must be one of: ' + VALID_PROGRESS_RATINGS.join(', ') };
  }

  // Validate anecdotalNotes
  var notes = String(data.anecdotalNotes || '');
  if (!notes || notes.trim().length < 10) {
    return { success: false, error: 'Anecdotal notes are required (minimum 10 characters).' };
  }

  initializeSheetsIfNeeded_();
  var status = getUserStatus();
  var ss;

  if (status.role === 'sped-lead') {
    // SPED Lead writes to target CM's spreadsheet
    var targetSpreadsheetId = data.caseManagerSpreadsheetId;
    if (!targetSpreadsheetId) {
      return {success: false, error: 'caseManagerSpreadsheetId required for SPED Lead writes'};
    }

    // Verify SPED Lead has access to this caseload
    var caseloads = getSpedLeadCaseloads_(getCurrentUserEmail_());
    var hasAccess = caseloads.some(function(cm) {
      return cm.spreadsheetId === targetSpreadsheetId;
    });
    if (!hasAccess) {
      return {success: false, error: 'Access denied to this caseload'};
    }

    ss = SpreadsheetApp.openById(targetSpreadsheetId);
  } else {
    ss = getSS_();
    var ctSheet = ss.getSheetByName(SHEET_COTEACHERS);
    var perms = getCallerPermissions_(ctSheet, getCurrentUserEmail_());
    requirePermission_(perms, 'editProgress', 'edit progress entries');
  }

  // Validate student exists
  var students = getStudents();
  var studentExists = false;
  for (var si = 0; si < students.length; si++) {
    if (students[si].id === data.studentId) { studentExists = true; break; }
  }
  if (!studentExists) return { success: false, error: 'Student not found.' };

  var sheet = ss.getSheetByName(SHEET_PROGRESS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PROGRESS);
    ensureHeaders_(sheet, PROGRESS_HEADERS);
  }

  var email = getCurrentUserEmail_();
  var now = new Date().toISOString();

  // Use LockService to prevent duplicate entries from concurrent saves
  var lock = LockService.getUserLock();
  try {
    lock.waitLock(10000); // Wait up to 10 seconds
  } catch(e) {
    return { success: false, error: 'Could not acquire lock. Please try again.' };
  }

  try {
    // Check for existing entry (upsert by studentId+goalId+objectiveId+quarter)
    var allData = sheet.getDataRange().getValues();
    var headers = allData[0];
    var colIdx = buildColIdx_(headers);
    var existingRow = null;
    var existingId = null;

    for (var i = 1; i < allData.length; i++) {
      var row = allData[i];
      if (row[colIdx.studentId - 1] === data.studentId &&
          row[colIdx.goalId - 1] === data.goalId &&
          row[colIdx.objectiveId - 1] === data.objectiveId &&
          row[colIdx.quarter - 1] === quarter) {
        existingRow = i + 1; // 1-based sheet row
        existingId = row[colIdx.id - 1];
        break;
      }
    }

    if (existingRow) {
      // Update existing entry
      batchSetValues_(sheet, existingRow, colIdx, {
        progressRating: rating,
        anecdotalNotes: notes.trim(),
        dateEntered: now.split('T')[0],
        enteredBy: email,
        lastModified: now
      });
      invalidateProgressCaches_();
      return { success: true, id: existingId, updated: true };
    } else {
      // Create new entry
      var id = Utilities.getUuid();
      sheet.appendRow([
        id,
      data.studentId,
      data.goalId,
      data.objectiveId,
      quarter,
      rating,
      notes.trim(),
      now.split('T')[0],
      email,
      now,
      now
    ]);
      invalidateProgressCaches_();
      return { success: true, id: id, updated: false };
    }
  } finally {
    lock.releaseLock();
  }
}

/** Internal: fetch progress entries for a student, optionally filtered by quarter. */
function getProgressEntries_(studentId, quarter) {
  initializeSheetsIfNeeded_();
  var ss = getSS_();
  var sheet = ss.getSheetByName(SHEET_PROGRESS);
  if (!sheet) return [];

  var allData = sheet.getDataRange().getValues();
  if (allData.length <= 1) return [];

  var headers = allData[0];
  var results = [];
  for (var i = 1; i < allData.length; i++) {
    var row = {};
    headers.forEach(function(h, idx) { row[h] = allData[i][idx]; });
    if (row.dateEntered instanceof Date) row.dateEntered = formatDateValue_(row.dateEntered);
    if (row.createdAt instanceof Date) row.createdAt = row.createdAt.toISOString();
    if (row.lastModified instanceof Date) row.lastModified = row.lastModified.toISOString();
    if (row.studentId === studentId && (!quarter || row.quarter === quarter)) {
      results.push(row);
    }
  }
  return results;
}

/** Public: get progress entries for a student in a specific quarter. */
function getProgressEntries(studentId, quarter) {
  initializeSheetsIfNeeded_();
  var _ss = getSS_();
  var _ctSheet = _ss.getSheetByName(SHEET_COTEACHERS);
  if (!checkPermission_(getCallerPermissions_(_ctSheet, getCurrentUserEmail_()), 'viewProgress')) return [];
  return getProgressEntries_(studentId, quarter);
}

/** Public: get all progress entries across all quarters for a student. */
function getAllProgressForStudent(studentId) {
  initializeSheetsIfNeeded_();
  var _ss = getSS_();
  var _ctSheet = _ss.getSheetByName(SHEET_COTEACHERS);
  if (!checkPermission_(getCallerPermissions_(_ctSheet, getCurrentUserEmail_()), 'viewProgress')) return [];
  return getProgressEntries_(studentId, null);
}

/** Delete a progress entry by ID (used by tests for cleanup). */
function deleteProgressEntry_(entryId) {
  var ss = getSS_();
  var sheet = ss.getSheetByName(SHEET_PROGRESS);
  if (!sheet) return;
  var found = findRowById_(sheet, entryId);
  if (found) {
    sheet.deleteRow(found.rowIndex);
    invalidateProgressCaches_();
  }
}

/** Calculate GPA from academicData array for report purposes.
 *  Returns { raw: number, rounded: string, excludedCount: number } or null if no gradeable classes. */
function calculateGpaForReport_(academicData) {
  if (!academicData || academicData.length === 0) return null;

  var gpaValues = [];
  var excluded = 0;
  academicData.forEach(function(c) {
    if (c.grade && GPA_MAP.hasOwnProperty(c.grade)) {
      gpaValues.push(GPA_MAP[c.grade]);
    } else {
      excluded++;
    }
  });

  if (gpaValues.length === 0) return null;

  var raw = gpaValues.reduce(function(a, b) { return a + b; }, 0) / gpaValues.length;
  return {
    raw: raw,
    rounded: raw.toFixed(2),
    excludedCount: excluded
  };
}

/** Extract and sort grades for report display.
 *  Returns array of { className, grade, missing } sorted by className. */
function getGradesForReport_(student) {
  var data = student.academicData || [];
  var grades = data.map(function(c) {
    return {
      className: c.className || '',
      grade: c.grade || '',
      missing: Number(c.missing) || 0
    };
  });
  grades.sort(function(a, b) {
    return a.className.localeCompare(b.className);
  });
  return grades;
}

/** Assemble all data needed for a progress report.
 *  Pure data function — no HTML generation.
 *  @param {Object} student — student record from dashboard data
 *  @param {string} quarter — e.g. 'Q2'
 *  @param {Array} allEntries — progress entries (can span multiple quarters for history)
 *  @returns {Object} report data object */
function assembleReportData_(student, quarter, allEntries) {
  // Parse goals
  var goals = [];
  try { goals = JSON.parse(student.goalsJson || '[]'); }
  catch(e) { goals = []; }

  // Build a lookup: goalId+objectiveId+quarter → entry
  var entryMap = {};
  (allEntries || []).forEach(function(e) {
    var key = e.goalId + '|' + e.objectiveId + '|' + e.quarter;
    entryMap[key] = e;
  });

  // Determine prior quarters (before current)
  var qOrder = ['Q1', 'Q2', 'Q3', 'Q4'];
  var currentIdx = qOrder.indexOf(quarter);
  var priorQuarters = currentIdx > 0 ? qOrder.slice(0, currentIdx) : [];

  // Group goals by goalArea
  var areaMap = {};
  var areaOrder = [];
  goals.forEach(function(goal) {
    var area = goal.goalArea || 'General';
    if (!areaMap[area]) {
      areaMap[area] = [];
      areaOrder.push(area);
    }

    var objectives = (goal.objectives || []).map(function(obj) {
      // Current quarter progress
      var currentKey = goal.id + '|' + obj.id + '|' + quarter;
      var currentEntry = entryMap[currentKey];
      var currentProgress = currentEntry
        ? { rating: currentEntry.progressRating, notes: currentEntry.anecdotalNotes }
        : { rating: 'Not yet reported', notes: '' };

      // Prior quarter history
      var history = [];
      priorQuarters.forEach(function(pq) {
        var pKey = goal.id + '|' + obj.id + '|' + pq;
        var pEntry = entryMap[pKey];
        if (pEntry) {
          history.push({ quarter: pq, rating: pEntry.progressRating, notes: pEntry.anecdotalNotes });
        }
      });

      return {
        id: obj.id,
        text: obj.text,
        currentProgress: currentProgress,
        progressHistory: history
      };
    });

    // Compute goal-level rating from objectives
    var goalLevelRating = null;
    if (objectives.length > 0) {
      var allMet = objectives.every(function(o) { return o.currentProgress.rating === 'Objective Met'; });
      var anyNoProgress = objectives.some(function(o) { return o.currentProgress.rating === 'No Progress'; });
      var anyReported = objectives.some(function(o) {
        return o.currentProgress.rating !== 'Not yet reported';
      });
      if (allMet) {
        goalLevelRating = 'Goal Met';
      } else if (anyNoProgress) {
        goalLevelRating = 'Insufficient Progress';
      } else if (anyReported) {
        goalLevelRating = 'Adequate Progress';
      }
    }

    areaMap[area].push({
      id: goal.id,
      text: goal.text,
      goalLevelRating: goalLevelRating,
      objectives: objectives
    });
  });

  var goalGroups = areaOrder.map(function(area) {
    return { goalArea: area, goals: areaMap[area] };
  });

  // Grades
  var grades = getGradesForReport_(student);
  var gpaResult = calculateGpaForReport_(student.academicData);

  // Summary counts
  var totalGoals = goals.length;
  var goalsWithAdequateOrMet = 0;
  var goalsWithNoProgress = 0;

  goals.forEach(function(goal) {
    var objs = goal.objectives || [];
    if (objs.length === 0) return;
    var allAdequateOrMet = true;
    var anyNoProgress = false;
    objs.forEach(function(obj) {
      var key = goal.id + '|' + obj.id + '|' + quarter;
      var entry = entryMap[key];
      if (!entry || entry.progressRating === 'No Progress') {
        allAdequateOrMet = false;
        anyNoProgress = true;
      }
    });
    // Mutually exclusive: a goal is either on track or needs attention, not both
    if (allAdequateOrMet) {
      goalsWithAdequateOrMet++;
    } else if (anyNoProgress) {
      goalsWithNoProgress++;
    }
  });

  return {
    summary: {
      studentName: (student.firstName || '') + ' ' + (student.lastName || ''),
      gradeLevel: student.grade || '',
      caseManager: student.caseManagerEmail || '',
      reportingPeriod: getQuarterLabel_(quarter),
      totalGoals: totalGoals,
      goalsWithAdequateOrMet: goalsWithAdequateOrMet,
      goalsWithNoProgress: goalsWithNoProgress
    },
    goalGroups: goalGroups,
    grades: grades,
    gpa: gpaResult
  };
}


/** Escape HTML special characters for safe embedding in report. */
function escHtml_(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Generate a printable HTML progress report for one student.
 *  Styled to match the official Richfield Public Schools progress report
 *  format, using the Lato font (Google Fonts) and RPS brand colors.
 *  @param {Object} student - student record
 *  @param {string} quarter - e.g. 'Q2'
 *  @param {Array} allEntries - progress entries (all quarters for history)
 *  @param {string} overallSummary - optional teacher-written summary
 *  @returns {string} complete HTML document string */
function generateProgressReportHtml_(student, quarter, allEntries, overallSummary) {
  var data = assembleReportData_(student, quarter, allEntries);
  var s = data.summary;
  var today = new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });

  // Checkbox helper: ☑ if match, ☐ otherwise
  function chk(goalRating, value) {
    return goalRating === value ? '&#x2611;' : '&#x2610;';
  }

  // Rating color helper (for timeline chips)
  function ratingColor(rating) {
    if (rating === 'Objective Met') return '#1B5E20';
    if (rating === 'Adequate Progress') return '#7A5900';
    if (rating === 'No Progress') return '#BA1A1A';
    return '#666';
  }
  function ratingBg(rating) {
    if (rating === 'Objective Met') return '#D6F5D6';
    if (rating === 'Adequate Progress') return '#FFF3CD';
    if (rating === 'No Progress') return '#FFDAD6';
    return '#F5F5F5';
  }

  var html = '';
  html += '<!DOCTYPE html><html><head><meta charset="utf-8">';
  html += '<title>Progress Report \u2014 ' + escHtml_(s.studentName) + '</title>';
  html += '<link rel="preconnect" href="https://fonts.googleapis.com">';
  html += '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>';
  html += '<link href="https://fonts.googleapis.com/css2?family=Lato:wght@400;700;900&display=swap" rel="stylesheet">';
  html += '<style>';

  // Base
  html += 'body { font-family: "Lato", "Segoe UI", Arial, sans-serif; font-size: 11pt; line-height: 1.6; color: #1C1B1F; max-width: 8.5in; margin: 0 auto; padding: 0.5in; }';
  html += 'p { margin: 4px 0; }';

  // Header
  html += '.report-header { display: flex; align-items: flex-start; justify-content: space-between; border-bottom: 3px solid #942022; padding-bottom: 14px; margin-bottom: 20px; }';
  html += '.header-left { display: flex; align-items: flex-start; gap: 14px; }';
  html += '.school-logo { width: 40px; height: 40px; background: #000; border-radius: 4px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }';
  html += '.school-logo svg { display: block; }';
  html += '.school-name { font-size: 15pt; font-weight: 700; margin: 0 0 2px 0; color: #1C1B1F; }';
  html += '.school-address { font-size: 9pt; color: #49454F; margin: 0; line-height: 1.4; }';
  html += '.report-title { font-size: 18pt; font-weight: 700; color: #942022; margin: 0; white-space: nowrap; align-self: center; }';

  // Student info table
  html += '.info-table { width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 10.5pt; }';
  html += '.info-table td { padding: 6px 12px; border: 1px solid #D8D8D8; }';
  html += '.info-table .label { font-weight: 700; color: #49454F; width: 130px; }';

  // Summary section
  html += '.summary-section { background: #F8F8F8; border: 1px solid #E0E0E0; border-radius: 8px; padding: 16px 20px; margin-bottom: 24px; }';
  html += '.summary-section p { font-size: 10.5pt; line-height: 1.7; margin: 0; }';

  // Goal sections
  html += '.goal-section { margin-bottom: 20px; break-inside: avoid; }';
  html += '.goal-heading { font-size: 12pt; font-weight: 700; margin: 20px 0 4px 0; color: #1C1B1F; }';
  html += '.goal-area-label { font-size: 9pt; font-weight: 700; color: #942022; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px; }';
  html += '.goal-text { margin: 4px 0 10px 0; font-size: 10.5pt; }';
  html += '.goal-date { font-size: 10pt; margin-bottom: 6px; }';

  // Progress extent checkboxes
  html += '.extent-line { font-size: 10pt; margin-bottom: 12px; color: #49454F; }';
  html += '.extent-line span { margin-right: 16px; }';
  html += '.extent-checked { font-weight: 700; color: #1C1B1F; }';

  // Objectives
  html += '.objective-block { margin: 8px 0 8px 20px; padding: 8px 0; }';
  html += '.objective-label { font-size: 10pt; font-weight: 700; color: #49454F; }';
  html += '.objective-text { font-size: 10.5pt; margin: 2px 0 4px 0; }';
  html += '.objective-progress { font-size: 10.5pt; font-weight: 700; font-style: italic; margin: 4px 0; }';
  html += '.progress-timeline { margin-top: 4px; font-size: 9pt; color: #666; }';
  html += '.timeline-chip { display: inline-block; padding: 1px 6px; border-radius: 8px; font-size: 8pt; margin-right: 4px; }';

  // Grades table
  html += '.grades-section { margin-top: 24px; }';
  html += '.grades-heading { font-size: 14pt; font-weight: 700; margin: 0 0 8px 0; padding-bottom: 6px; border-bottom: 2px solid #942022; }';
  html += '.grades-table { width: 100%; border-collapse: collapse; margin-top: 8px; font-size: 10pt; }';
  html += '.grades-table th { background: #F5F5F5; text-align: left; padding: 6px 12px; border: 1px solid #D8D8D8; font-weight: 700; }';
  html += '.grades-table td { padding: 6px 12px; border: 1px solid #E8E8E8; }';
  html += '.gpa-display { margin-top: 10px; font-size: 11pt; }';
  html += '.gpa-value { font-weight: 700; font-size: 14pt; }';

  // Footer
  html += '.report-footer { margin-top: 32px; padding-top: 12px; border-top: 2px solid #942022; font-size: 9pt; color: #49454F; }';
  html += '.report-footer p { margin: 4px 0; }';

  // Print
  html += '@media print {';
  html += '  body { padding: 0.25in; margin: 0; }';
  html += '  .goal-section { break-inside: avoid; }';
  html += '  .objective-block { break-inside: avoid; }';
  html += '  .summary-section { background: #F8F8F8 !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }';
  html += '  .extent-checked { font-weight: 700 !important; }';
  html += '  .timeline-chip { -webkit-print-color-adjust: exact; print-color-adjust: exact; }';
  html += '  .grades-table th { background: #F5F5F5 !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }';
  html += '}';

  html += '</style></head><body>';

  // ── Header ──
  html += '<div class="report-header">';
  html += '<div class="header-left">';
  html += '<div class="school-logo"><svg width="28" height="28" viewBox="0 0 40 40" xmlns="http://www.w3.org/2000/svg"><rect width="40" height="40" rx="4" fill="#000"/><text x="50%" y="54%" dominant-baseline="central" text-anchor="middle" font-family="Lato,sans-serif" font-weight="700" font-size="28" fill="#fff">R</text></svg></div>';
  html += '<div>';
  html += '<p class="school-name">Richfield Public Schools</p>';
  html += '<p class="school-address">401 70th Street W (door 26)<br>Richfield MN 55423-3061<br>Tel: 612-798-6000</p>';
  html += '</div>';
  html += '</div>';
  html += '<div class="report-title">Progress report</div>';
  html += '</div>';

  // ── Student Info Table ──
  html += '<table class="info-table">';
  html += '<tr><td class="label">Student:</td><td>' + escHtml_(s.studentName) + '</td>';
  html += '<td class="label">Plan manager:</td><td>' + escHtml_(s.caseManager) + '</td></tr>';
  html += '<tr><td class="label">School:</td><td>Richfield Middle School</td>';
  html += '<td class="label">Grade:</td><td>' + escHtml_(s.gradeLevel) + '</td></tr>';
  html += '<tr><td class="label">Reporting period:</td><td>' + escHtml_(s.reportingPeriod) + '</td>';
  html += '<td class="label">Date:</td><td>' + today + '</td></tr>';
  html += '</table>';

  // ── Summary (if provided) ──
  if (overallSummary && String(overallSummary).trim()) {
    html += '<div class="summary-section">';
    html += '<p>' + escHtml_(overallSummary) + '</p>';
    html += '</div>';
  }

  // ── Goals ──
  if (data.goalGroups.length === 0) {
    html += '<p style="color:#49454F;margin-top:16px;">No IEP goals have been entered for this student.</p>';
  }

  var goalNum = 0;
  data.goalGroups.forEach(function(group) {
    group.goals.forEach(function(goal) {
      goalNum++;
      html += '<div class="goal-section">';
      html += '<p class="goal-area-label">' + escHtml_(group.goalArea) + '</p>';
      html += '<p class="goal-heading">Goal ' + goalNum + ':</p>';
      html += '<p class="goal-text">' + escHtml_(goal.text) + '</p>';
      html += '<p class="goal-date">Date: ' + today + '</p>';

      // Progress extent checkboxes
      html += '<p class="extent-line">';
      html += 'The extent to which that progress is sufficient to enable the pupil to achieve the goals by the end of the year:<br>';
      var ratings = [
        { value: 'Insufficient Progress', label: 'Insufficient progress' },
        { value: 'Adequate Progress', label: 'Adequate progress' },
        { value: 'Goal Met', label: 'Goal met' }
      ];
      ratings.forEach(function(r) {
        var isChecked = goal.goalLevelRating === r.value;
        html += '<span class="' + (isChecked ? 'extent-checked' : '') + '">' + chk(goal.goalLevelRating, r.value) + ' ' + r.label + '</span>';
      });
      html += '</p>';

      // Objectives
      goal.objectives.forEach(function(obj, idx) {
        html += '<div class="objective-block">';
        html += '<p class="objective-label">Objective ' + (idx + 1) + ':</p>';
        html += '<p class="objective-text">' + escHtml_(obj.text) + '</p>';

        var r = obj.currentProgress.rating;
        if (r !== 'Not yet reported' || obj.currentProgress.notes) {
          html += '<p class="objective-progress">Progress: ';
          if (obj.currentProgress.notes) {
            html += escHtml_(obj.currentProgress.notes);
          } else {
            html += escHtml_(r);
          }
          html += '</p>';
        }

        // Progress timeline (prior quarters)
        if (obj.progressHistory.length > 0) {
          html += '<div class="progress-timeline">Prior: ';
          obj.progressHistory.forEach(function(h) {
            html += '<span class="timeline-chip" style="background:' + ratingBg(h.rating) + ';color:' + ratingColor(h.rating) + ';">' + h.quarter + ': ' + escHtml_(h.rating) + '</span> ';
          });
          html += '</div>';
        }

        html += '</div>';
      });

      html += '</div>';
    });
  });

  // ── Grades Section ──
  html += '<div class="grades-section">';
  html += '<p class="grades-heading">Academic Snapshot</p>';

  if (data.grades.length === 0) {
    html += '<p style="color:#49454F;">Grades not yet available.</p>';
  } else {
    html += '<table class="grades-table">';
    html += '<thead><tr><th>Class</th><th>Grade</th><th>Missing Assignments</th></tr></thead>';
    html += '<tbody>';
    data.grades.forEach(function(g) {
      var isPassFail = !GPA_MAP.hasOwnProperty(g.grade) && g.grade;
      html += '<tr>';
      html += '<td>' + escHtml_(g.className) + '</td>';
      html += '<td>' + escHtml_(g.grade) + (isPassFail ? ' <em style="font-size:9pt;color:#666;">(P/F)</em>' : '') + '</td>';
      html += '<td>' + g.missing + '</td>';
      html += '</tr>';
    });
    html += '</tbody></table>';

    html += '<div class="gpa-display">';
    if (data.gpa) {
      html += 'GPA: <span class="gpa-value">' + data.gpa.rounded + '</span>';
      if (data.gpa.excludedCount > 0) {
        html += ' <em style="font-size:9pt;color:#666;">(' + data.gpa.excludedCount + ' pass/fail class' + (data.gpa.excludedCount > 1 ? 'es' : '') + ' excluded)</em>';
      }
    } else {
      html += 'GPA: <span class="gpa-value">N/A</span>';
    }
    html += '</div>';
  }
  html += '</div>';

  // ── Footer ──
  html += '<div class="report-footer">';
  html += '<p>This progress report is provided in accordance with the Individuals with Disabilities Education Act (IDEA). Parents/guardians are encouraged to contact the plan manager listed above with any questions regarding their child\'s progress or IEP.</p>';
  html += '<p><strong>Questions?</strong> Contact ' + escHtml_(s.caseManager) + '</p>';
  html += '</div>';

  html += '</body></html>';
  return html;
}

/** Public endpoint: generate a progress report for one student.
 *  Called from the frontend. */
function generateProgressReport(studentId, quarter, overallSummary) {
  if (VALID_QUARTERS.indexOf(String(quarter)) === -1) {
    return { success: false, error: 'Invalid quarter.' };
  }

  initializeSheetsIfNeeded_();
  var _ss = getSS_();
  var _ctSheet = _ss.getSheetByName(SHEET_COTEACHERS);
  var _perms = getCallerPermissions_(_ctSheet, getCurrentUserEmail_());
  requirePermission_(_perms, 'viewProgress', 'generate progress reports');
  var students = getStudents();
  var student = null;
  for (var i = 0; i < students.length; i++) {
    if (students[i].id === studentId) { student = students[i]; break; }
  }
  if (!student) return { success: false, error: 'Student not found.' };

  // Get dashboard data for GPA/academic info
  var dashData = getDashboardData();
  var dashStudent = null;
  for (var j = 0; j < dashData.length; j++) {
    if (dashData[j].id === studentId) { dashStudent = dashData[j]; break; }
  }

  // Merge academic data into student record
  if (dashStudent) {
    student.academicData = dashStudent.academicData || [];
    student.gpa = dashStudent.gpa;
  } else {
    student.academicData = [];
    student.gpa = null;
  }

  var allEntries = getAllProgressForStudent(studentId);
  var html = generateProgressReportHtml_(student, quarter, allEntries, overallSummary || '');
  return { success: true, html: html };
}

/** Public endpoint: batch generate reports for entire caseload. */
function generateBatchReports(quarter, overallSummaries) {
  if (VALID_QUARTERS.indexOf(quarter) === -1) {
    return { success: false, error: 'Invalid quarter: ' + quarter };
  }

  try {
    initializeSheetsIfNeeded_();
    var _ss = getSS_();
    var _ctSheet = _ss.getSheetByName(SHEET_COTEACHERS);
    var _perms = getCallerPermissions_(_ctSheet, getCurrentUserEmail_());
    requirePermission_(_perms, 'viewProgress', 'generate batch reports');
    var dashData = getDashboardData();
    var summaries = overallSummaries || {};
    var results = [];

    dashData.forEach(function(student) {
      var allEntries = getAllProgressForStudent(student.id);
      var summary = summaries[student.id] || '';

      // Ensure goalsJson is available
      if (!student.goalsJson) student.goalsJson = '[]';

      var html = generateProgressReportHtml_(student, quarter, allEntries, summary);
      results.push({
        studentId: student.id,
        studentName: (student.firstName || '') + ' ' + (student.lastName || ''),
        html: html
      });
    });

    return { success: true, reports: results };
  } catch (e) {
    return { success: false, error: 'Failed to generate batch reports: ' + e.message };
  }
}

// ───── Due Process Completion Flags ─────

/**
 * Toggle a due process report completion flag.
 * Flags are stored in UserProperties as JSON: { "studentId|goalId|quarter": true }
 */
function toggleDPReportComplete(studentId, goalId, quarter, complete) {
  if (!studentId || !goalId || VALID_QUARTERS.indexOf(quarter) === -1) {
    return { success: false, error: 'Invalid parameters' };
  }
  initializeSheetsIfNeeded_();
  var _ss = getSS_();
  var _ctSheet = _ss.getSheetByName(SHEET_COTEACHERS);
  var _perms = getCallerPermissions_(_ctSheet, getCurrentUserEmail_());
  requirePermission_(_perms, 'editProgress', 'mark progress complete');
  var lock = LockService.getUserLock();
  try {
    lock.waitLock(5000);
    var props = PropertiesService.getUserProperties();
    var flagsJson = props.getProperty('dp_completion_flags') || '{}';
    var flags = JSON.parse(flagsJson);
    var key = studentId + '|' + goalId + '|' + quarter;
    if (complete) {
      flags[key] = true;
    } else {
      delete flags[key];
    }
    props.setProperty('dp_completion_flags', JSON.stringify(flags));
  } finally {
    lock.releaseLock();
  }
  invalidateProgressCaches_();
  return { success: true };
}

/** Read completion flags from UserProperties. */
function getDPCompletionFlags_() {
  var props = PropertiesService.getUserProperties();
  var flagsJson = props.getProperty('dp_completion_flags') || '{}';
  return JSON.parse(flagsJson);
}

// ───── Gemini AI ─────

var GEMINI_MODEL_ = 'gemini-2.0-flash';
var GEMINI_API_URL_ = 'https://generativelanguage.googleapis.com/v1beta/models/';

/**
 * One-time setup: run this function from the Script Editor to store
 * your Gemini API key in Script Properties. It is shared across all
 * users of the web app but never sent to the frontend.
 *
 * Usage: In the Script Editor, select setGeminiApiKey and click Run.
 *        When prompted, enter your API key.
 */
function setGeminiApiKey(key) {
  if (!key) {
    // When run manually from Script Editor, prompt via Logger
    throw new Error('Usage: call setGeminiApiKey("your-api-key-here") from the Script Editor console, or set the script property GEMINI_API_KEY manually in Project Settings > Script Properties.');
  }
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
  Logger.log('Gemini API key stored successfully.');
  return { success: true };
}

/** Retrieve the stored API key. Returns null if not set. */
function getGeminiApiKey_() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || null;
}

/**
 * Core Gemini API call. All AI features should use this function.
 *
 * @param {string} prompt - The user/system prompt text
 * @param {Object} [opts] - Optional configuration
 * @param {string} [opts.model] - Model name (default: GEMINI_MODEL_)
 * @param {string} [opts.systemInstruction] - System instruction for the model
 * @param {number} [opts.temperature] - Temperature 0-2 (default: 0.7)
 * @param {number} [opts.maxOutputTokens] - Max tokens in response (default: 1024)
 * @returns {string} The generated text response
 * @throws {Error} If API key is missing or the API call fails
 */
function callGemini_(prompt, opts) {
  var apiKey = getGeminiApiKey_();
  if (!apiKey) {
    throw new Error('Gemini API key not configured. Run setGeminiApiKey() or add GEMINI_API_KEY in Script Properties.');
  }

  opts = opts || {};
  var model = opts.model || GEMINI_MODEL_;
  var temperature = opts.temperature != null ? opts.temperature : 0.7;
  var maxOutputTokens = opts.maxOutputTokens || 1024;

  var url = GEMINI_API_URL_ + model + ':generateContent?key=' + apiKey;

  var requestBody = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: temperature,
      maxOutputTokens: maxOutputTokens
    }
  };

  if (opts.systemInstruction) {
    requestBody.systemInstruction = {
      parts: [{ text: opts.systemInstruction }]
    };
  }

  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var body = JSON.parse(response.getContentText());

  if (code !== 200) {
    var errMsg = (body.error && body.error.message) || ('Gemini API error: HTTP ' + code);
    throw new Error(errMsg);
  }

  if (!body.candidates || !body.candidates.length ||
      !body.candidates[0].content || !body.candidates[0].content.parts ||
      !body.candidates[0].content.parts.length) {
    throw new Error('Gemini returned an empty response.');
  }

  return body.candidates[0].content.parts[0].text;
}

/**
 * Structured Gemini call that returns parsed JSON.
 * Use for features that need structured data (e.g., suggested goals, categorization).
 *
 * @param {string} prompt - The prompt (should instruct the model to return JSON)
 * @param {Object} [opts] - Same options as callGemini_, plus:
 * @param {Object} [opts.responseSchema] - JSON schema for structured output
 * @returns {Object} Parsed JSON response
 */

/** Lightweight check: does a Gemini API key exist? No network call. */
function hasGeminiKey() {
  return { available: !!getGeminiApiKey_() };
}

/**
 * Public endpoint: check whether Gemini is configured and reachable.
 * Called from frontend to show/hide AI features.
 */
function getAiStatus() {
  var key = getGeminiApiKey_();
  if (!key) {
    return { available: false, reason: 'API key not configured' };
  }
  try {
    var result = callGemini_('Respond with exactly: OK', {
      temperature: 0,
      maxOutputTokens: 10
    });
    return { available: true, model: GEMINI_MODEL_ };
  } catch (e) {
    return { available: false, reason: e.message };
  }
}

/**
 * Public endpoint: generate an AI-powered progress summary for one student.
 * Assembles all available data (goals, ratings, check-ins, academics) and
 * sends a structured prompt to Gemini.
 */
function generateAiProgressSummary(studentId, quarter) {
  if (!studentId) return { success: false, error: 'Missing studentId.' };
  if (VALID_QUARTERS.indexOf(String(quarter)) === -1) {
    return { success: false, error: 'Invalid quarter.' };
  }

  // ── Fetch student + academic data (same pattern as generateProgressReport) ──
  initializeSheetsIfNeeded_();
  var _ss = getSS_();
  var _ctSheet = _ss.getSheetByName(SHEET_COTEACHERS);
  var _perms = getCallerPermissions_(_ctSheet, getCurrentUserEmail_());
  requirePermission_(_perms, 'editProgress', 'generate AI summaries');
  var students = getStudents();
  var student = null;
  for (var i = 0; i < students.length; i++) {
    if (students[i].id === studentId) { student = students[i]; break; }
  }
  if (!student) return { success: false, error: 'Student not found.' };

  var dashData = getDashboardData();
  var dashStudent = null;
  for (var j = 0; j < dashData.length; j++) {
    if (dashData[j].id === studentId) { dashStudent = dashData[j]; break; }
  }
  if (dashStudent) {
    student.academicData = dashStudent.academicData || [];
    student.gpa = dashStudent.gpa;
  } else {
    student.academicData = [];
    student.gpa = null;
  }

  // ── Assemble progress data ──
  var allEntries = getAllProgressForStudent(studentId);
  var data = assembleReportData_(student, quarter, allEntries);

  // ── Gather recent check-in context ──
  var checkIns = getCheckIns(studentId);
  var recentCheckIns = checkIns.slice(0, 4);

  // ── Build the prompt ──
  var lines = [];
  lines.push('Write a quarterly progress summary for the following student.');
  lines.push('');
  lines.push('Student: ' + (student.firstName || '') + ' ' + (student.lastName || '') + ', Grade ' + (student.grade || ''));
  lines.push('Reporting Period: ' + data.summary.reportingPeriod);
  lines.push('GPA: ' + (data.gpa ? data.gpa.rounded : 'N/A'));
  lines.push('');

  // Goals & objectives
  lines.push('IEP Goals & Progress:');
  var goalNum = 0;
  data.goalGroups.forEach(function(group) {
    group.goals.forEach(function(goal) {
      goalNum++;
      lines.push('');
      lines.push('Goal ' + goalNum + ' (' + group.goalArea + '): ' + goal.text);
      goal.objectives.forEach(function(obj, idx) {
        lines.push('  Objective ' + (idx + 1) + ': ' + obj.text);
        lines.push('    Current Rating: ' + obj.currentProgress.rating);
        if (obj.currentProgress.notes) {
          lines.push('    Teacher Notes: ' + obj.currentProgress.notes);
        }
        if (obj.progressHistory.length > 0) {
          var hist = obj.progressHistory.map(function(h) { return h.quarter + ': ' + h.rating; }).join(', ');
          lines.push('    Prior Quarters: ' + hist);
        }
      });
    });
  });

  // Grades
  if (data.grades.length > 0) {
    lines.push('');
    lines.push('Current Grades:');
    data.grades.forEach(function(g) {
      lines.push('  ' + g.className + ': ' + g.grade + (g.missing > 0 ? ' (' + g.missing + ' missing)' : ''));
    });
  }

  // Recent check-in trends
  if (recentCheckIns.length > 0) {
    lines.push('');
    lines.push('Recent Check-In Trends (last ' + recentCheckIns.length + ' weeks):');
    recentCheckIns.forEach(function(ci) {
      var avg = 0;
      var count = 0;
      ['planningRating', 'followThroughRating', 'regulationRating', 'focusGoalRating', 'effortRating'].forEach(function(k) {
        var v = Number(ci[k]);
        if (v > 0) { avg += v; count++; }
      });
      avg = count > 0 ? (avg / count).toFixed(1) : 'N/A';
      var parts = ['Week of ' + ci.weekOf + ': EF avg ' + avg + '/5'];
      if (ci.barrier) parts.push('Barrier: ' + ci.barrier);
      if (ci.whatWentWell) parts.push('Positive: ' + ci.whatWentWell);
      if (ci.microGoal) parts.push('Micro-goal: ' + ci.microGoal);
      lines.push('  ' + parts.join(' | '));
    });
  }

  var prompt = lines.join('\n');

  var systemInstruction = 'You are a special education case manager writing a brief, parent-friendly progress summary for an IEP progress report. Write 2-3 short paragraphs in plain language. Be encouraging but honest about areas needing growth. Reference specific data (ratings, accuracy percentages, grades) where available. Do not use markdown, bullet points, or headers \u2014 write flowing prose only.';

  try {
    var summary = callGemini_(prompt, {
      systemInstruction: systemInstruction,
      temperature: 0.7,
      maxOutputTokens: 1024
    });
    return { success: true, summary: summary };
  } catch (e) {
    return { success: false, error: e.message || 'Failed to generate AI summary.' };
  }
}

/**
 * Public endpoint: read a Google Sheets spreadsheet, send to Gemini for a
 * special-ed summary, create a Google Doc, and attach it to the eval's files.
 *
 * @param {string} evalId - Evaluation ID to attach the doc to
 * @param {string} driveFileId - Drive file ID of the source spreadsheet
 * @returns {Object} { success, files, newFile, docUrl } or { success: false, error }
 */
function generateSheetSummary(evalId, driveFileId) {
  if (!evalId) return { success: false, error: 'Missing evalId.' };
  if (!driveFileId) return { success: false, error: 'Missing driveFileId.' };

  var apiKey = getGeminiApiKey_();
  if (!apiKey) return { success: false, error: 'Gemini API key not configured.' };

  // Look up eval and student
  var evalObj = getEvaluationById_(evalId);
  if (!evalObj) return { success: false, error: 'Evaluation not found.' };

  var students = getStudents();
  var student = null;
  for (var i = 0; i < students.length; i++) {
    if (students[i].id === evalObj.studentId) { student = students[i]; break; }
  }

  // Open and read the spreadsheet
  var ss;
  try {
    ss = SpreadsheetApp.openById(driveFileId);
  } catch (e) {
    return { success: false, error: 'Cannot open spreadsheet. Check sharing permissions.' };
  }

  var sheet = ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: false, error: 'Spreadsheet has no data rows.' };

  // Convert to header:value pairs for the prompt
  var headers = data[0];
  var rows = data.slice(1);
  var records = rows.map(function(row) {
    var record = {};
    headers.forEach(function(h, idx) {
      if (h && row[idx] !== '') {
        var val = row[idx];
        if (val instanceof Date) val = formatDateValue_(val);
        record[String(h).trim()] = String(val).trim();
      }
    });
    return record;
  });

  var fileName = ss.getName();

  // Build the prompt
  var prompt = 'You are given survey/assessment data exported from a Google Sheets file called "' + fileName + '".\n\n';
  prompt += 'The data contains ' + records.length + ' response(s).\n\n';
  records.forEach(function(rec, idx) {
    prompt += 'Response ' + (idx + 1) + ':\n';
    Object.keys(rec).forEach(function(key) {
      prompt += '  ' + key + ': ' + rec[key] + '\n';
    });
    prompt += '\n';
  });

  if (student) {
    prompt += 'Student context: ' + (student.firstName || '') + ' ' + (student.lastName || '') + ', Grade ' + (student.grade || '') + '\n';
  }

  prompt += '\nWrite a professional narrative summary of these results suitable for pasting into a special education evaluation report (3-year reevaluation or IEP). ';
  prompt += 'Organize by transition domains or assessment domains as appropriate to the data. ';
  prompt += 'Include a section on strengths/protective factors and a section on areas of need. ';
  prompt += 'End with recommendations for the IEP team. ';
  prompt += 'Write in flowing prose paragraphs \u2014 no bullet points, no markdown headers, no numbered lists. ';
  prompt += 'Use professional, objective language appropriate for a legal educational document. ';
  prompt += 'Reference the student by full name (not pronouns alone) in the first mention of each paragraph.';

  var systemInstruction = 'You are a Minnesota special education case manager writing assessment summaries for inclusion in IEP evaluation reports. You follow IDEA federal requirements and Minnesota Rules Chapter 3525. Your writing is data-driven, objective, and uses person-first language. You frame findings through a strengths-based lens while honestly identifying areas of need. You understand that for 7th-8th grade students, age-appropriate transition assessment data informs early transition planning that formally begins no later than age 14 or 9th grade in Minnesota.';

  // Call Gemini
  var summaryText;
  try {
    summaryText = callGemini_(prompt, {
      systemInstruction: systemInstruction,
      temperature: 0.5,
      maxOutputTokens: 2048
    });
  } catch (e) {
    return { success: false, error: 'AI generation failed: ' + (e.message || String(e)) };
  }

  // Create a Google Doc with the summary
  var docTitle = fileName + ' Summary';
  var doc = DocumentApp.create(docTitle);
  var body = doc.getBody();
  body.clear();

  var title = body.appendParagraph(docTitle);
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);

  var meta = 'Generated: ' + new Date().toLocaleDateString() + ' | Source: ' + fileName;
  if (student) meta += ' | Student: ' + student.firstName + ' ' + student.lastName;
  var metaPara = body.appendParagraph(meta);
  metaPara.setFontSize(9);
  metaPara.setForegroundColor('#666666');
  body.appendParagraph('');

  var paragraphs = summaryText.split(/\n\n+/);
  paragraphs.forEach(function(p) {
    if (p.trim()) {
      var para = body.appendParagraph(p.trim());
      para.setFontSize(11);
      para.setLineSpacing(1.15);
    }
  });

  doc.saveAndClose();
  var docUrl = doc.getUrl();
  var docId = doc.getId();

  // Attach the new doc to the evaluation's files
  var result = addEvalFile(evalId, {
    driveFileId: docId,
    name: docTitle,
    mimeType: 'application/vnd.google-apps.document',
    url: docUrl
  });

  invalidateEvalCaches_();

  if (!result.success) return { success: false, error: 'Summary created but failed to attach: ' + result.error };
  return { success: true, files: result.files, newFile: result.newFile, docUrl: docUrl };
}

// ───── Helpers ─────

/** Normalize a value that may be a Date object into YYYY-MM-DD string */
function formatDateValue_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var y = val.getFullYear();
    var m = ('0' + (val.getMonth() + 1)).slice(-2);
    var d = ('0' + val.getDate()).slice(-2);
    return y + '-' + m + '-' + d;
  }
  return String(val);
}

function initializeSheetsIfNeeded_() {
  const ss = getSS_();
  if (!ss.getSheetByName(SHEET_STUDENTS) || !ss.getSheetByName(SHEET_CHECKINS) ||
      !ss.getSheetByName(SHEET_COTEACHERS) || !ss.getSheetByName(SHEET_IEP_MEETINGS) ||
      !ss.getSheetByName(SHEET_EVALUATIONS) || !ss.getSheetByName(SHEET_PROGRESS)) {
    initializeSheets();
  }
}

/** Build a {headerName: 1-based-column-index} map from a headers array. */
function buildColIdx_(headers) {
  var colIdx = {};
  headers.forEach(function(h, i) { colIdx[h] = i + 1; });
  return colIdx;
}

/**
 * Find a data row by its ID (column 0). Returns {rowIndex, colIdx} or null.
 * rowIndex is 1-based (sheet row number). colIdx maps header names to 1-based columns.
 */
function findRowById_(sheet, id) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colIdx = buildColIdx_(headers);
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      return { rowIndex: i + 1, colIdx: colIdx };
    }
  }
  return null;
}

/** Update multiple cells in a single row. fields is {headerName: value}. */
function batchSetValues_(sheet, rowIndex, colIdx, fields) {
  for (var key in fields) {
    if (fields.hasOwnProperty(key) && colIdx[key]) {
      sheet.getRange(rowIndex, colIdx[key]).setValue(fields[key]);
    }
  }
}

/** Normalize legacy 'owner' role to 'caseload-manager'. */
function normalizeRole_(role) {
  role = String(role || '').toLowerCase();
  return role === 'owner' ? 'caseload-manager' : role;
}

// ───── Menu (for bound-script usage) ─────
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Caseload Dashboard')
      .addItem('Open Caseload Dashboard', 'openWebApp')
      .addItem('Initialize / Reset Sheets', 'initializeSheets')
      .addItem('Get Web App URL', 'showWebAppUrl')
      .addToUi();
  } catch(e) {
    // Not running as a bound script; silently skip
  }
}

function openWebApp() {
  const html = HtmlService.createHtmlOutput(
    '<p style="font-family:sans-serif;">Opening\u2026</p>' +
    '<script>window.open("' + ScriptApp.getService().getUrl() + '");google.script.host.close();</script>'
  ).setWidth(300).setHeight(80);
  SpreadsheetApp.getUi().showModalDialog(html, 'Caseload Dashboard');
}

function showWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  const html = HtmlService.createHtmlOutput(
    '<p style="font-family:sans-serif;margin-bottom:12px;">Your webapp URL:</p>' +
    '<input type="text" value="' + url + '" style="width:100%;padding:8px;font-size:13px;" onclick="this.select()" readonly>'
  ).setWidth(450).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Web App URL');
}

// ───── Case Manager Management (Superuser Admin) ─────

/** Get the global list of case managers. Accessible to all users. */
function getCaseManagers() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty('case_managers');
    if (raw) return JSON.parse(raw);
  } catch(e) {}
  return [];
}

/** Check if the current user is the superuser. */
function isSuperuser() {
  return getCurrentUserEmail_() === SUPERUSER_EMAIL;
}

/** Add a case manager to the global list. Superuser only. */
function addCaseManager(email, name) {
  if (getCurrentUserEmail_() !== SUPERUSER_EMAIL) {
    return { success: false, error: 'Unauthorized' };
  }
  email = String(email || '').trim().toLowerCase();
  name = String(name || '').trim();
  if (!email || email.indexOf('@') === -1) {
    return { success: false, error: 'Please enter a valid email address.' };
  }
  if (!name) {
    return { success: false, error: 'Please enter a display name.' };
  }

  var scriptProps = PropertiesService.getScriptProperties();
  var list = [];
  try {
    var raw = scriptProps.getProperty('case_managers');
    if (raw) list = JSON.parse(raw);
  } catch(e) {}

  var exists = list.some(function(cm) { return cm.email === email; });
  if (exists) {
    return { success: false, error: 'This case manager already exists.' };
  }

  list.push({ email: email, name: name });
  scriptProps.setProperty('case_managers', JSON.stringify(list));
  return { success: true };
}

/** Remove a case manager from the global list. Superuser only. */
function removeCaseManager(email) {
  if (getCurrentUserEmail_() !== SUPERUSER_EMAIL) {
    return { success: false, error: 'Unauthorized' };
  }
  email = String(email || '').trim().toLowerCase();

  var scriptProps = PropertiesService.getScriptProperties();
  var list = [];
  try {
    var raw = scriptProps.getProperty('case_managers');
    if (raw) list = JSON.parse(raw);
  } catch(e) {}

  list = list.filter(function(cm) { return cm.email !== email; });
  scriptProps.setProperty('case_managers', JSON.stringify(list));
  return { success: true };
}

/** Get all students for the admin assignment view. Superuser only. */
function getAllStudentsForAdmin() {
  if (getCurrentUserEmail_() !== SUPERUSER_EMAIL) {
    return { success: false, error: 'Unauthorized' };
  }
  var students = getStudents();
  return students.map(function(s) {
    return {
      id: s.id,
      firstName: s.firstName,
      lastName: s.lastName,
      grade: s.grade,
      caseManagerEmail: s.caseManagerEmail || ''
    };
  });
}

/** Assign a case manager to a student. Superuser only. */
function assignCaseManager(studentId, caseManagerEmail) {
  if (getCurrentUserEmail_() !== SUPERUSER_EMAIL) {
    return { success: false, error: 'Unauthorized' };
  }
  var ss = getSS_();
  var sheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheet) return { success: false, error: 'Students sheet not found.' };

  var found = findRowById_(sheet, studentId);
  if (!found || !found.colIdx['caseManagerEmail']) {
    return { success: false, error: found ? 'caseManagerEmail column not found. Please reload.' : 'Student not found.' };
  }

  sheet.getRange(found.rowIndex, found.colIdx['caseManagerEmail']).setValue(caseManagerEmail || '');
  invalidateStudentCaches_();
  return { success: true };
}

// ─── SPED Lead Functions ───

/** Get list of SPED Lead emails from ScriptProperties. */
function getSpedLeads_() {
  var raw = PropertiesService.getScriptProperties().getProperty('sped_leads');
  if (!raw) return [];
  try {
    return JSON.parse(raw);
  } catch(e) {
    Logger.log('Error parsing sped_leads: ' + e.message);
    return [];
  }
}

/** Get case manager caseloads for a SPED Lead. */
function getSpedLeadCaseloads_(spedLeadEmail) {
  var raw = PropertiesService.getScriptProperties().getProperty('sped_lead_caseloads_' + spedLeadEmail);
  if (!raw) return [];
  try {
    return JSON.parse(raw);
  } catch(e) {
    Logger.log('Error parsing sped_lead_caseloads: ' + e.message);
    return [];
  }
}

/** Get SPED Lead's spreadsheet ID. */
function getSpedLeadSpreadsheetId_(spedLeadEmail) {
  return PropertiesService.getScriptProperties().getProperty('sped_lead_spreadsheet_' + spedLeadEmail);
}

/** Get eval metrics from a case manager's spreadsheet. */
function getEvalMetrics_(ss) {
  var evalsSheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!evalsSheet) {
    return {activeCount: 0, overdueCount: 0, dueThisWeekCount: 0};
  }

  var data = evalsSheet.getDataRange().getValues();
  if (data.length <= 1) {
    return {activeCount: 0, overdueCount: 0, dueThisWeekCount: 0};
  }

  var headers = data[0];
  var colIdx = buildColIdx_(headers);
  if (!colIdx.meetingDate) {
    return {activeCount: 0, overdueCount: 0, dueThisWeekCount: 0};
  }

  var now = new Date();
  var oneWeekFromNow = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);

  var activeCount = 0;
  var overdueCount = 0;
  var dueThisWeekCount = 0;

  for (var i = 1; i < data.length; i++) {
    var meetingDate = data[i][colIdx.meetingDate - 1];
    if (!meetingDate) continue;

    activeCount++;

    if (meetingDate < now) {
      overdueCount++;
    } else if (meetingDate <= oneWeekFromNow) {
      dueThisWeekCount++;
    }
  }

  return {
    activeCount: activeCount,
    overdueCount: overdueCount,
    dueThisWeekCount: dueThisWeekCount
  };
}

/** Get due process metrics from a case manager's spreadsheet. */
function getDueProcessMetrics_(ss) {
  var iepSheet = ss.getSheetByName(SHEET_IEP_MEETINGS);
  var progressSheet = ss.getSheetByName(SHEET_PROGRESS);

  var upcomingIEPs = 0;
  var progressCompletionRate = 0;

  // Count upcoming IEPs (next 7 days)
  if (iepSheet) {
    var iepData = iepSheet.getDataRange().getValues();
    if (iepData.length > 1) {
      var headers = iepData[0];
      var colIdx = buildColIdx_(headers);
      if (colIdx.meetingDate) {
        var now = new Date();
        var oneWeekFromNow = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);

        for (var i = 1; i < iepData.length; i++) {
          var meetingDate = iepData[i][colIdx.meetingDate - 1];
          if (meetingDate && meetingDate >= now && meetingDate <= oneWeekFromNow) {
            upcomingIEPs++;
          }
        }
      }
    }
  }

  // Calculate progress completion rate (simplified - just check if entries exist)
  if (progressSheet) {
    var progressData = progressSheet.getDataRange().getValues();
    if (progressData.length > 1) {
      var headers = progressData[0];
      var colIdx = buildColIdx_(headers);

      if (colIdx.progressRating) {
        var totalEntries = progressData.length - 1;
        var completedEntries = 0;

        for (var i = 1; i < progressData.length; i++) {
          var progressRating = progressData[i][colIdx.progressRating - 1];
          if (progressRating) completedEntries++;
        }

        progressCompletionRate = totalEntries > 0 ? Math.round((completedEntries / totalEntries) * 100) : 0;
      }
    }
  }

  return {
    upcomingIEPs: upcomingIEPs,
    progressCompletionRate: progressCompletionRate
  };
}

/** Get student summaries from a case manager's spreadsheet. */
function getStudentSummaries_(ss) {
  var studentsSheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!studentsSheet) return [];

  var data = studentsSheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var headers = data[0];
  var colIdx = buildColIdx_(headers);

  var students = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    students.push({
      studentId: row[colIdx.id - 1],
      firstName: row[colIdx.firstName - 1] || '',
      lastName: row[colIdx.lastName - 1] || '',
      grade: row[colIdx.grade - 1] || '',
      evalType: '', // Will be enriched from Evaluations sheet
      evalStatus: '',
      nextIEPDate: null,
      gpa: null
    });
  }

  // Enrich with eval data
  var evalsSheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (evalsSheet) {
    var evalData = evalsSheet.getDataRange().getValues();
    if (evalData.length > 1) {
      var evalHeaders = evalData[0];
      var evalColIdx = buildColIdx_(evalHeaders);

      for (var i = 1; i < evalData.length; i++) {
        var studentId = evalData[i][evalColIdx.studentId - 1];
        var student = null;
        for (var j = 0; j < students.length; j++) {
          if (students[j].studentId === studentId) {
            student = students[j];
            break;
          }
        }
        if (student) {
          student.evalType = evalData[i][evalColIdx.type - 1] || '';
        }
      }
    }
  }

  return students;
}

/** Sync SPED Lead dashboard from all connected case manager spreadsheets. */
function syncSpedLeadDashboard(spedLeadEmail) {
  var caseloads = getSpedLeadCaseloads_(spedLeadEmail);
  var spedLeadSSId = getSpedLeadSpreadsheetId_(spedLeadEmail);

  if (!spedLeadSSId) {
    return {success: false, error: 'SPED Lead spreadsheet not provisioned'};
  }

  var spedLeadSS = SpreadsheetApp.openById(spedLeadSSId);
  var aggregateSheet = spedLeadSS.getSheetByName('AggregateMetrics');
  var studentsSheet = spedLeadSS.getSheetByName('AllStudents');
  var timelineSheet = spedLeadSS.getSheetByName('ComplianceTimeline');

  if (!aggregateSheet || !studentsSheet || !timelineSheet) {
    return {success: false, error: 'SPED Lead spreadsheet missing required sheets'};
  }

  // Clear existing data (keep headers)
  if (aggregateSheet.getLastRow() > 1) {
    aggregateSheet.deleteRows(2, aggregateSheet.getLastRow() - 1);
  }
  if (studentsSheet.getLastRow() > 1) {
    studentsSheet.deleteRows(2, studentsSheet.getLastRow() - 1);
  }
  if (timelineSheet.getLastRow() > 1) {
    timelineSheet.deleteRows(2, timelineSheet.getLastRow() - 1);
  }

  var syncedCount = 0;
  var failedCount = 0;
  var errors = [];

  caseloads.forEach(function(cm) {
    try {
      var cmSS = SpreadsheetApp.openById(cm.spreadsheetId);

      // Get metrics
      var evalMetrics = getEvalMetrics_(cmSS);
      var dpMetrics = getDueProcessMetrics_(cmSS);
      var students = getStudentSummaries_(cmSS);

      // Write to AggregateMetrics
      aggregateSheet.appendRow([
        cm.email,
        cm.name,
        students.length,
        evalMetrics.activeCount,
        evalMetrics.overdueCount,
        dpMetrics.upcomingIEPs,
        dpMetrics.progressCompletionRate,
        'OK',
        new Date()
      ]);

      // Write to AllStudents
      students.forEach(function(s) {
        studentsSheet.appendRow([
          s.studentId,
          s.firstName,
          s.lastName,
          cm.email,
          cm.name,
          s.grade,
          s.evalType,
          s.evalStatus,
          s.nextIEPDate,
          s.gpa
        ]);
      });

      syncedCount++;
      Utilities.sleep(500); // Rate limiting

    } catch(e) {
      failedCount++;
      errors.push({cm: cm.email, error: e.message});

      // Write failed status to AggregateMetrics
      aggregateSheet.appendRow([
        cm.email,
        cm.name,
        0,
        0,
        0,
        0,
        0,
        'Failed',
        new Date()
      ]);
    }
  });

  // Update last sync timestamp
  PropertiesService.getScriptProperties().setProperty(
    'sped_lead_last_sync_' + spedLeadEmail,
    new Date().toISOString()
  );

  return {
    success: true,
    syncedCount: syncedCount,
    failedCount: failedCount,
    errors: errors
  };
}

/** Install daily 2am sync trigger for a SPED Lead. */
function installSpedLeadSyncTrigger(spedLeadEmail) {
  // Delete existing trigger if any
  removeSpedLeadSyncTrigger(spedLeadEmail);

  // Create daily trigger at 2am
  var trigger = ScriptApp.newTrigger('onSpedLeadDailySyncTrigger')
    .timeBased()
    .atHour(2)
    .everyDays(1)
    .create();

  // Store trigger ID
  PropertiesService.getScriptProperties().setProperty(
    'sped_lead_trigger_' + spedLeadEmail,
    trigger.getUniqueId()
  );

  return {success: true, triggerId: trigger.getUniqueId()};
}

/** Remove SPED Lead sync trigger. */
function removeSpedLeadSyncTrigger(spedLeadEmail) {
  var triggerId = PropertiesService.getScriptProperties().getProperty('sped_lead_trigger_' + spedLeadEmail);
  if (!triggerId) return;

  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getUniqueId() === triggerId) {
      ScriptApp.deleteTrigger(triggers[i]);
      break;
    }
  }

  PropertiesService.getScriptProperties().deleteProperty('sped_lead_trigger_' + spedLeadEmail);
}

/** Daily sync trigger handler - syncs all SPED Leads. */
function onSpedLeadDailySyncTrigger(e) {
  var spedLeads = getSpedLeads_();
  spedLeads.forEach(function(email) {
    try {
      syncSpedLeadDashboard(email);
    } catch(err) {
      Logger.log('Failed to sync SPED Lead ' + email + ': ' + err.message);
    }
  });
}

/** Check if email is a SPED Lead. */
function isSpedLead_(email) {
  var spedLeads = getSpedLeads_();
  return spedLeads.indexOf(email) !== -1;
}

/** Add a SPED Lead (superuser only). */
function addSpedLead(email, _testAuthorized) {
  // Validate caller is superuser (bypass for tests)
  if (!_testAuthorized) {
    var currentEmail = getCurrentUserEmail_();
    if (currentEmail !== SUPERUSER_EMAIL) {
      return {success: false, error: 'Only superuser can add SPED Leads'};
    }
  }

  email = String(email || '').trim().toLowerCase();
  if (!email || email.indexOf('@') === -1) {
    return {success: false, error: 'Invalid email address'};
  }

  var spedLeads = getSpedLeads_();
  if (spedLeads.indexOf(email) !== -1) {
    return {success: false, error: 'Email already registered as SPED Lead'};
  }

  spedLeads.push(email);
  PropertiesService.getScriptProperties().setProperty('sped_leads', JSON.stringify(spedLeads));

  return {success: true};
}

function removeSpedLead(email, _testAuthorized) {
  // Validate caller is superuser (bypass for tests)
  if (!_testAuthorized) {
    var currentEmail = getCurrentUserEmail_();
    if (currentEmail !== SUPERUSER_EMAIL) {
      return {success: false, error: 'Only superuser can remove SPED Leads'};
    }
  }

  email = String(email || '').trim().toLowerCase();

  var spedLeads = getSpedLeads_();
  var index = spedLeads.indexOf(email);
  if (index === -1) {
    return {success: false, error: 'Email not found in SPED Leads'};
  }

  spedLeads.splice(index, 1);
  PropertiesService.getScriptProperties().setProperty('sped_leads', JSON.stringify(spedLeads));

  // Cleanup related properties
  PropertiesService.getScriptProperties().deleteProperty('sped_lead_caseloads_' + email);
  PropertiesService.getScriptProperties().deleteProperty('sped_lead_spreadsheet_' + email);
  PropertiesService.getScriptProperties().deleteProperty('sped_lead_last_sync_' + email);

  // Remove trigger
  removeSpedLeadSyncTrigger(email);

  return {success: true};
}

function updateSpedLeadConnections(spedLeadEmail, caseManagerEmails, _testAuthorized) {
  // Validate caller is superuser (bypass for tests)
  if (!_testAuthorized) {
    var currentEmail = getCurrentUserEmail_();
    if (currentEmail !== SUPERUSER_EMAIL) {
      return {success: false, error: 'Only superuser can update connections'};
    }
  }

  spedLeadEmail = String(spedLeadEmail || '').trim().toLowerCase();
  if (!spedLeadEmail) {
    return {success: false, error: 'Invalid SPED Lead email'};
  }

  var spedLeads = getSpedLeads_();
  if (spedLeads.indexOf(spedLeadEmail) === -1) {
    return {success: false, error: 'Email not registered as SPED Lead'};
  }

  if (!caseManagerEmails || !Array.isArray(caseManagerEmails)) {
    return {success: false, error: 'caseManagerEmails must be an array'};
  }

  // Build caseload array with CM details
  var caseloads = [];
  var caseManagers = getCaseManagers(); // Existing function

  caseManagerEmails.forEach(function(cmEmail) {
    cmEmail = String(cmEmail || '').trim().toLowerCase();
    var cm = caseManagers.find(function(c) { return c.email === cmEmail; });
    if (cm && cm.spreadsheetId) {
      caseloads.push({
        email: cm.email,
        name: cm.name || cm.email,
        spreadsheetId: cm.spreadsheetId
      });

      // Grant SPED Lead editor access to CM's spreadsheet
      try {
        var cmSS = SpreadsheetApp.openById(cm.spreadsheetId);
        cmSS.addEditor(spedLeadEmail);
      } catch(e) {
        Logger.log('Failed to share ' + cm.email + ' spreadsheet with SPED Lead: ' + e.message);
      }
    }
  });

  // Store connections
  PropertiesService.getScriptProperties().setProperty(
    'sped_lead_caseloads_' + spedLeadEmail,
    JSON.stringify(caseloads)
  );

  return {success: true, connectedCount: caseloads.length};
}

function provisionSpedLeadSpreadsheet(spedLeadEmail, _testAuthorized) {
  // Validate caller is superuser (bypass for tests)
  if (!_testAuthorized) {
    var currentEmail = getCurrentUserEmail_();
    if (currentEmail !== SUPERUSER_EMAIL) {
      return {success: false, error: 'Only superuser can provision spreadsheets'};
    }
  }

  spedLeadEmail = String(spedLeadEmail || '').trim().toLowerCase();
  if (!spedLeadEmail) {
    return {success: false, error: 'Invalid SPED Lead email'};
  }

  var spedLeads = getSpedLeads_();
  if (spedLeads.indexOf(spedLeadEmail) === -1) {
    return {success: false, error: 'Email not registered as SPED Lead'};
  }

  // Check if already provisioned
  var existingId = getSpedLeadSpreadsheetId_(spedLeadEmail);
  if (existingId) {
    return {success: false, error: 'Spreadsheet already provisioned'};
  }

  // Create spreadsheet
  var name = 'SPED Lead Dashboard - ' + spedLeadEmail.split('@')[0];
  var ss = SpreadsheetApp.create(name);
  var ssId = ss.getId();

  // Delete default sheet
  var defaultSheet = ss.getSheets()[0];
  if (defaultSheet.getName() === 'Sheet1') {
    ss.deleteSheet(defaultSheet);
  }

  // Create AggregateMetrics sheet
  var aggregateSheet = ss.insertSheet('AggregateMetrics');
  aggregateSheet.appendRow([
    'caseManagerEmail',
    'caseManagerName',
    'studentCount',
    'activeEvals',
    'overdueEvals',
    'upcomingIEPs',
    'progressCompletionRate',
    'lastSyncStatus',
    'lastSyncDate'
  ]);

  // Create AllStudents sheet
  var studentsSheet = ss.insertSheet('AllStudents');
  studentsSheet.appendRow([
    'studentId',
    'firstName',
    'lastName',
    'caseManagerEmail',
    'caseManagerName',
    'grade',
    'evalType',
    'evalStatus',
    'nextIEPDate',
    'gpa'
  ]);

  // Create ComplianceTimeline sheet
  var timelineSheet = ss.insertSheet('ComplianceTimeline');
  timelineSheet.appendRow([
    'date',
    'studentId',
    'studentName',
    'caseManagerEmail',
    'type',
    'meetingType',
    'evalType'
  ]);

  // Store spreadsheet ID
  PropertiesService.getScriptProperties().setProperty(
    'sped_lead_spreadsheet_' + spedLeadEmail,
    ssId
  );

  // Run initial sync
  var syncResult = syncSpedLeadDashboard(spedLeadEmail);

  // Install daily trigger
  var triggerResult = installSpedLeadSyncTrigger(spedLeadEmail);

  return {
    success: true,
    spreadsheetId: ssId,
    syncResult: syncResult,
    triggerResult: triggerResult
  };
}

/** Get SPED Lead dashboard data (public endpoint). */
function getSpedLeadDashboardData() {
  var userEmail = getCurrentUserEmail_();
  if (!isSpedLead_(userEmail)) {
    return {error: 'Access denied'};
  }

  // Auto-sync if stale (>24 hours)
  var lastSync = PropertiesService.getScriptProperties().getProperty('sped_lead_last_sync_' + userEmail);
  if (lastSync) {
    var lastSyncDate = new Date(lastSync);
    var hoursSinceSync = (Date.now() - lastSyncDate.getTime()) / 1000 / 60 / 60;
    if (hoursSinceSync > 24) {
      syncSpedLeadDashboard(userEmail);
    }
  }

  // Placeholder - return empty data for now
  return {
    caseManagers: [],
    lastSyncDate: lastSync
  };
}
