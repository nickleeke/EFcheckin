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

  return {
    isNewUser: !ssId,
    email: email,
    pendingInvite: invite,
    isSuperuser: email === SUPERUSER_EMAIL
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
  'online','contactsJson','birthday'
];
var CHECKIN_HEADERS = [
  'id','studentId','weekOf',
  'planningRating','followThroughRating','regulationRating',
  'focusGoalRating','effortRating',
  'whatWentWell','barrier',
  'microGoal','microGoalCategory',
  'teacherNotes','academicDataJson','createdAt','goalMet'
];
var COTEACHER_HEADERS = ['email', 'role', 'addedAt'];
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
  const ss = getSS_();
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
    profile.online ? 'TRUE' : '', contactsJson, profile.birthday || ''
  ]);
  invalidateStudentCaches_();
  return { success: true, id: id };
}

function saveStudentGoals(studentId, goalsJson) {
  initializeSheetsIfNeeded_();
  const ss = getSS_();
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
  const ss = getSS_();
  const studentsSheet = ss.getSheetByName(SHEET_STUDENTS);
  if (studentsSheet) {
    const sData = studentsSheet.getDataRange().getValues();
    for (let i = sData.length - 1; i >= 1; i--) {
      if (sData[i][0] === studentId) { studentsSheet.deleteRow(i + 1); break; }
    }
  }
  const checkInsSheet = ss.getSheetByName(SHEET_CHECKINS);
  if (checkInsSheet && checkInsSheet.getLastRow() > 1) {
    const cData = checkInsSheet.getDataRange().getValues();
    for (let i = cData.length - 1; i >= 1; i--) {
      if (cData[i][1] === studentId) checkInsSheet.deleteRow(i + 1);
    }
  }
  const evalSheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (evalSheet && evalSheet.getLastRow() > 1) {
    const eData = evalSheet.getDataRange().getValues();
    for (let i = eData.length - 1; i >= 1; i--) {
      if (eData[i][1] === studentId) evalSheet.deleteRow(i + 1);
    }
  }
  invalidateCache_();
  return { success: true };
}

// ───── Check-In CRUD ─────

function saveCheckIn(data) {
  initializeSheetsIfNeeded_();
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
}

function getCheckIns(studentId) {
  initializeSheetsIfNeeded_();
  const ss = getSS_();
  const sheet = ss.getSheetByName(SHEET_CHECKINS);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const results = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === studentId) {
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
  const ss = getSS_();
  const sheet = ss.getSheetByName(SHEET_CHECKINS);
  if (!sheet) return { success: false };
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === checkInId) { sheet.deleteRow(i + 1); invalidateCheckInCaches_(); return { success: true }; }
  }
  return { success: false };
}

function updateCheckInAcademicData(checkInId, academicData) {
  initializeSheetsIfNeeded_();
  const ss = getSS_();
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
    let avgRating = null;
    if (latest) {
      const ratings = [
        Number(latest.planningRating), Number(latest.followThroughRating),
        Number(latest.regulationRating), Number(latest.focusGoalRating),
        Number(latest.effortRating)
      ].filter(function(r) { return !isNaN(r) && r > 0; });
      if (ratings.length > 0) {
        avgRating = ratings.reduce(function(a, b) { return a + b; }, 0) / ratings.length;
      }
    }

    // Trend
    let trend = 'none';
    if (checkIns.length >= 2 && avgRating !== null) {
      const prev = checkIns[1];
      const prevRatings = [
        Number(prev.planningRating), Number(prev.followThroughRating),
        Number(prev.regulationRating), Number(prev.focusGoalRating),
        Number(prev.effortRating)
      ].filter(function(r) { return !isNaN(r) && r > 0; });
      if (prevRatings.length > 0) {
        const prevAvg = prevRatings.reduce(function(a, b) { return a + b; }, 0) / prevRatings.length;
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
    const historySlice = checkIns.slice(0, 6).reverse();
    historySlice.forEach(function(ci) {
      const rs = [
        Number(ci.planningRating), Number(ci.followThroughRating),
        Number(ci.regulationRating), Number(ci.focusGoalRating),
        Number(ci.effortRating)
      ].filter(function(r) { return !isNaN(r) && r > 0; });
      if (rs.length > 0) {
        efHistory.push(rs.reduce(function(a, b) { return a + b; }, 0) / rs.length);
      }
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
    VALID_QUARTERS.forEach(function(q) {
      props.deleteProperty(CACHE_PREFIX + 'due_process_' + q);
    });
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
  var cached = getCache_('eval_summary');
  if (cached) return cached;

  initializeSheetsIfNeeded_();
  var ss = getSS_();
  var evalSheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!evalSheet || evalSheet.getLastRow() <= 1) {
    return { dueThisWeekCount: 0, overdueCount: 0, timeline: [] };
  }

  // Build student name lookup
  var students = getStudents();
  var studentMap = {};
  students.forEach(function(s) {
    studentMap[s.id] = { firstName: s.firstName, lastName: s.lastName };
  });

  var evalData = evalSheet.getDataRange().getValues();
  var evalHeaders = evalData[0];
  var evalColIdx = buildColIdx_(evalHeaders);

  var now = new Date();
  var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var todayStr = formatDateValue_(today);

  // Build 7-day range
  var dayAbbrs = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  var monthAbbrs = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var timelineDays = [];
  for (var d = 0; d < 7; d++) {
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
  var endDateStr = timelineDays[6].date;

  var overdueCount = 0;
  var dueThisWeekCount = 0;
  var activeEvals = [];
  var overdueTasks = [];
  var dueThisWeekTasks = [];

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
    items.forEach(function(item) {
      if (item.checked) { evalDone++; return; }
      if (!item.dueDate) return;

      if (item.dueDate < todayStr) {
        overdueCount++;
        evalOverdue++;
        overdueTasks.push({
          itemId: item.id,
          text: item.text,
          dueDate: item.dueDate,
          studentId: studentId,
          evalId: evalId,
          studentName: studentFullName,
          evalType: evalType
        });
        return;
      }

      if (item.dueDate >= todayStr && item.dueDate <= endDateStr) {
        dueThisWeekCount++;
        dueThisWeekTasks.push({
          itemId: item.id,
          text: item.text,
          dueDate: item.dueDate,
          studentId: studentId,
          evalId: evalId,
          studentName: studentFullName,
          evalType: evalType
        });
        for (var t = 0; t < timelineDays.length; t++) {
          if (timelineDays[t].date === item.dueDate) {
            timelineDays[t].tasks.push({
              itemId: item.id,
              text: item.text,
              studentId: studentId,
              evalId: evalId,
              studentName: studentFullName,
              evalType: evalType
            });
            break;
          }
        }
      }
    });

    activeEvals.push({
      evalId: evalId,
      studentId: studentId,
      studentName: studentFullName,
      type: evalType,
      done: evalDone,
      total: items.length,
      overdueCount: evalOverdue
    });
  }

  // Sort overdue by date ascending (oldest first)
  overdueTasks.sort(function(a, b) { return a.dueDate < b.dueDate ? -1 : a.dueDate > b.dueDate ? 1 : 0; });
  // Sort due this week by date ascending
  dueThisWeekTasks.sort(function(a, b) { return a.dueDate < b.dueDate ? -1 : a.dueDate > b.dueDate ? 1 : 0; });

  var result = {
    dueThisWeekCount: dueThisWeekCount,
    overdueCount: overdueCount,
    timeline: timelineDays,
    activeEvals: activeEvals,
    overdueTasks: overdueTasks,
    dueThisWeekTasks: dueThisWeekTasks
  };

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
  var ss = getSS_();
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
  var sheet = ss.getSheetByName(SHEET_IEP_MEETINGS);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var allData = sheet.getDataRange().getValues();
  var headers = allData[0];
  var results = [];
  for (var i = 1; i < allData.length; i++) {
    var row = {};
    headers.forEach(function(h, idx) { row[h] = allData[i][idx]; });
    if (!studentId || row.studentId === studentId) results.push(row);
  }
  return results;
}

function deleteIEPMeeting(meetingId) {
  if (!meetingId) return { success: false, error: 'meetingId is required' };
  initializeSheetsIfNeeded_();
  var ss = getSS_();
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
  var cacheKey = 'due_process_' + quarter;
  var cached = getCache_(cacheKey);
  if (cached) return cached;

  initializeSheetsIfNeeded_();
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
  var studentMap = {};
  students.forEach(function(s) {
    studentMap[s.id] = { firstName: s.firstName, lastName: s.lastName };
  });

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

    activeEvals.push({
      evalId: evalId, studentId: studentId, studentName: studentFullName,
      type: evalType, done: evalDone, total: items.length, overdueCount: evalOverdue,
      nextDueDate: evalNextDue
    });
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
 * Public endpoint: returns all meetings (eval + standalone) for calendar views.
 */
function getAllMeetings() {
  initializeSheetsIfNeeded_();
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
        notes: row.notes || '',
        source: 'standalone'
      });
    }
  }

  return meetings;
}

function getEvalMeetingLabel_(evalType) {
  if (evalType === 'annual-iep') return 'Annual IEP Meeting';
  if (evalType === '3-year-reeval' || evalType === 'reeval') return 'Re-Eval Meeting';
  if (evalType === 'initial-eval' || evalType === 'eval') return 'Initial Eval Meeting';
  return 'Eval Meeting';
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
    member.role = normalizeRole_(member.role);
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

  return { members: members, currentUserRole: currentUserRole };
}

/** Valid roles that a caseload manager can assign to team members. */
var ASSIGNABLE_ROLES = ['service-provider', 'para', 'co-teacher'];

/** Invite a team member by email with a specified role. Only caseload managers can add members. */
function addTeamMember(email, role) {
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

  // Add to CoTeachers sheet with the specified role
  ctSheet.appendRow([email, role, new Date().toISOString()]);

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
function refreshDashboardData() {
  invalidateCache_();
  return getDashboardData();
}

// ───── Evaluation Checklist CRUD ─────

function getEvaluation(studentId) {
  initializeSheetsIfNeeded_();
  var ss = getSS_();
  var sheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!sheet || sheet.getLastRow() <= 1) return null;

  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === studentId) {
      var row = {};
      headers.forEach(function(h, idx) { row[h] = data[i][idx]; });
      try { row.items = JSON.parse(row.itemsJson || '[]'); }
      catch(e) { row.items = []; }
      try { row.files = JSON.parse(row.filesJson || '[]'); }
      catch(e) { row.files = []; }
      return row;
    }
  }
  return null;
}

function createEvaluation(studentId, type) {
  initializeSheetsIfNeeded_();
  if (VALID_EVAL_TYPES.indexOf(type) === -1) {
    return { success: false, error: 'Invalid evaluation type.' };
  }

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
}

function updateEvalMeetingDate(evalId, meetingDate) {
  initializeSheetsIfNeeded_();
  var ss = getSS_();
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
  var ss = getSS_();
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

// Deprecated: use saveEvaluationItems instead
function updateEvaluationItem(evalId, itemId, updates) {
  // Backward compat: old callers pass a boolean for checked
  if (typeof updates === 'boolean') {
    updates = { checked: updates };
  }

  initializeSheetsIfNeeded_();
  var ss = getSS_();
  var sheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!sheet) return { success: false, error: 'Evaluations sheet not found.' };

  var found = findRowById_(sheet, evalId);
  if (!found) return { success: false, error: 'Evaluation not found.' };

  var itemsRaw = sheet.getRange(found.rowIndex, found.colIdx['itemsJson']).getValue();
  var items;
  try { items = JSON.parse(itemsRaw || '[]'); }
  catch(e) { items = []; }

  var updated = false;
  for (var i = 0; i < items.length; i++) {
    if (items[i].id === itemId) {
      if (updates.hasOwnProperty('checked')) {
        items[i].checked = !!updates.checked;
        items[i].completedAt = updates.checked ? new Date().toISOString() : null;
      }
      if (updates.hasOwnProperty('text')) {
        items[i].text = String(updates.text).trim();
      }
      if (updates.hasOwnProperty('dueDate')) {
        items[i].dueDate = updates.dueDate || null;
      }
      updated = true;
      break;
    }
  }

  if (!updated) return { success: false, error: 'Item not found.' };

  batchSetValues_(sheet, found.rowIndex, found.colIdx, {
    itemsJson: JSON.stringify(items),
    updatedAt: new Date().toISOString()
  });

  return { success: true, items: items };
}

function deleteEvaluation(evalId) {
  var ss = getSS_();
  var sheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!sheet) return { success: false };

  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === evalId) {
      sheet.deleteRow(i + 1);
      invalidateEvalCaches_();
      return { success: true };
    }
  }
  return { success: false };
}

// Deprecated: use saveEvaluationItems instead
function addEvaluationItem(evalId, itemData) {
  initializeSheetsIfNeeded_();
  var ss = getSS_();
  var sheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!sheet) return { success: false, error: 'Evaluations sheet not found.' };

  var found = findRowById_(sheet, evalId);
  if (!found) return { success: false, error: 'Evaluation not found.' };

  var itemsRaw = sheet.getRange(found.rowIndex, found.colIdx['itemsJson']).getValue();
  var items;
  try { items = JSON.parse(itemsRaw || '[]'); } catch(e) { items = []; }

  var newItem = {
    id: 'item-custom-' + Utilities.getUuid().substr(0, 8),
    text: String(itemData.text || '').trim(),
    checked: false,
    completedAt: null,
    dueDate: itemData.dueDate || null
  };

  if (!newItem.text) return { success: false, error: 'Task text is required.' };

  items.push(newItem);

  batchSetValues_(sheet, found.rowIndex, found.colIdx, {
    itemsJson: JSON.stringify(items),
    updatedAt: new Date().toISOString()
  });

  return { success: true, items: items, newItem: newItem };
}

// Deprecated: use saveEvaluationItems instead
function deleteEvaluationItem(evalId, itemId) {
  initializeSheetsIfNeeded_();
  var ss = getSS_();
  var sheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!sheet) return { success: false, error: 'Evaluations sheet not found.' };

  var found = findRowById_(sheet, evalId);
  if (!found) return { success: false, error: 'Evaluation not found.' };

  var itemsRaw = sheet.getRange(found.rowIndex, found.colIdx['itemsJson']).getValue();
  var items;
  try { items = JSON.parse(itemsRaw || '[]'); } catch(e) { items = []; }

  var originalLength = items.length;
  items = items.filter(function(it) { return it.id !== itemId; });

  if (items.length === originalLength) return { success: false, error: 'Item not found.' };

  batchSetValues_(sheet, found.rowIndex, found.colIdx, {
    itemsJson: JSON.stringify(items),
    updatedAt: new Date().toISOString()
  });

  return { success: true, items: items };
}

// ───── Drive File Browser & Eval Files ─────

/** Trigger Drive scope — never called, but ensures the scope is added. */
function triggerDriveScope_() { DriveApp.getRootFolder(); }

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
function getQuarterLabel(quarter) {
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

  // Validate student exists
  var students = getStudents();
  var studentExists = false;
  for (var si = 0; si < students.length; si++) {
    if (students[si].id === data.studentId) { studentExists = true; break; }
  }
  if (!studentExists) return { success: false, error: 'Student not found.' };

  var ss = getSS_();
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

/** Get all progress entries for a student in a specific quarter. */
function getProgressEntries(studentId, quarter) {
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
    if (row.studentId === studentId && row.quarter === quarter) {
      results.push(row);
    }
  }
  return results;
}

/** Get all progress entries across all quarters for a student. */
function getAllProgressForStudent(studentId) {
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
    if (row.studentId === studentId) {
      results.push(row);
    }
  }
  return results;
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

    areaMap[area].push({
      id: goal.id,
      text: goal.text,
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
      }
      if (!entry || entry.progressRating === 'No Progress') {
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
      reportingPeriod: getQuarterLabel(quarter),
      totalGoals: totalGoals,
      goalsWithAdequateOrMet: goalsWithAdequateOrMet,
      goalsWithNoProgress: goalsWithNoProgress
    },
    goalGroups: goalGroups,
    grades: grades,
    gpa: gpaResult
  };
}

/** Student-friendly labels for progress ratings. */
var PROGRESS_FRIENDLY_LABELS = {
  'No Progress': "Let's keep working on this",
  'Adequate Progress': "You're making progress!",
  'Objective Met': 'You got it! \u2B50'
};

/** Escape HTML special characters for safe embedding in report. */
function escHtml_(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Generate a printable HTML progress report for one student.
 *  @param {Object} student — student record
 *  @param {string} quarter — e.g. 'Q2'
 *  @param {Array} allEntries — progress entries (all quarters for history)
 *  @param {string} overallSummary — optional teacher-written summary
 *  @returns {string} complete HTML document string */
function generateProgressReportHtml_(student, quarter, allEntries, overallSummary) {
  var data = assembleReportData_(student, quarter, allEntries);
  var s = data.summary;

  // Rating color helper
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
  html += '<title>IEP Progress Report \u2014 ' + escHtml_(s.studentName) + '</title>';
  html += '<style>';

  // Base styles
  html += 'body { font-family: "Roboto", Arial, sans-serif; font-size: 11pt; line-height: 1.5; color: #1C1B1F; max-width: 8.5in; margin: 0 auto; padding: 0.5in; }';
  html += 'h1 { font-size: 18pt; font-weight: 500; margin: 0 0 4px 0; color: #C41E3A; }';
  html += 'h2 { font-size: 14pt; font-weight: 500; margin: 24px 0 12px 0; padding-bottom: 6px; border-bottom: 2px solid #C41E3A; color: #1C1B1F; }';
  html += 'h3 { font-size: 12pt; font-weight: 500; margin: 16px 0 8px 0; color: #1C1B1F; }';
  html += 'p { margin: 4px 0; }';

  // Header
  html += '.report-header { border-bottom: 3px solid #C41E3A; padding-bottom: 12px; margin-bottom: 16px; }';
  html += '.report-subtitle { font-size: 10pt; color: #49454F; }';
  html += '.student-info { display: flex; flex-wrap: wrap; gap: 24px; margin-top: 8px; font-size: 10pt; }';
  html += '.student-info dt { font-weight: 500; color: #49454F; }';
  html += '.student-info dd { margin: 0 0 4px 0; }';

  // Summary card
  html += '.summary-card { background: #F7F2FA; border-radius: 12px; padding: 16px 20px; margin-bottom: 20px; }';
  html += '.summary-stats { display: flex; gap: 24px; flex-wrap: wrap; margin-top: 8px; }';
  html += '.stat-item { text-align: center; }';
  html += '.stat-value { font-size: 20pt; font-weight: 500; }';
  html += '.stat-label { font-size: 9pt; color: #49454F; }';

  // Goal sections
  html += '.goal-area-section { margin-bottom: 20px; break-inside: avoid; }';
  html += '.goal-block { margin-left: 12px; margin-bottom: 12px; }';
  html += '.goal-text { font-style: italic; margin-bottom: 8px; }';
  html += '.objective-row { margin-left: 16px; margin-bottom: 12px; padding: 8px 12px; border-left: 3px solid #D8C2C2; }';
  html += '.rating-badge { display: inline-block; padding: 2px 10px; border-radius: 12px; font-size: 10pt; font-weight: 500; }';
  html += '.friendly-label { font-size: 9pt; font-style: italic; color: #49454F; margin-left: 8px; }';
  html += '.anecdotal-notes { margin-top: 4px; font-size: 10pt; color: #49454F; }';
  html += '.progress-timeline { margin-top: 4px; font-size: 9pt; color: #666; }';
  html += '.timeline-chip { display: inline-block; padding: 1px 6px; border-radius: 8px; font-size: 8pt; margin-right: 4px; }';

  // Grades table
  html += '.grades-table { width: 100%; border-collapse: collapse; margin-top: 8px; font-size: 10pt; }';
  html += '.grades-table th { background: #F5F5F5; text-align: left; padding: 8px 12px; border-bottom: 2px solid #D8C2C2; font-weight: 500; }';
  html += '.grades-table td { padding: 8px 12px; border-bottom: 1px solid #E8E8E8; }';
  html += '.gpa-display { margin-top: 8px; font-size: 11pt; }';
  html += '.gpa-value { font-weight: 500; font-size: 14pt; }';

  // Footer
  html += '.report-footer { margin-top: 32px; padding-top: 12px; border-top: 1px solid #D8C2C2; font-size: 9pt; color: #49454F; }';

  // Print styles
  html += '@media print {';
  html += '  body { padding: 0; margin: 0; }';
  html += '  .goal-area-section { break-inside: avoid; }';
  html += '  .objective-row { break-inside: avoid; }';
  html += '  .summary-card { background: #F7F2FA !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }';
  html += '  .rating-badge { -webkit-print-color-adjust: exact; print-color-adjust: exact; }';
  html += '  .timeline-chip { -webkit-print-color-adjust: exact; print-color-adjust: exact; }';
  html += '}';

  html += '</style></head><body>';

  // ── Header ──
  html += '<div class="report-header">';
  html += '<h1>IEP Progress Report</h1>';
  html += '<p class="report-subtitle">Richfield Public Schools</p>';
  html += '<div class="student-info">';
  html += '<div><dt>Student</dt><dd><strong>' + escHtml_(s.studentName) + '</strong></dd></div>';
  html += '<div><dt>Grade</dt><dd>' + escHtml_(s.gradeLevel) + '</dd></div>';
  html += '<div><dt>Case Manager</dt><dd>' + escHtml_(s.caseManager) + '</dd></div>';
  html += '<div><dt>Reporting Period</dt><dd>' + escHtml_(s.reportingPeriod) + '</dd></div>';
  html += '<div><dt>Date Generated</dt><dd>' + new Date().toLocaleDateString() + '</dd></div>';
  html += '</div></div>';

  // ── Summary Card ──
  html += '<div class="summary-card">';
  html += '<h3 style="margin-top:0;">Progress Summary</h3>';

  if (overallSummary && String(overallSummary).trim()) {
    html += '<p>' + escHtml_(overallSummary) + '</p>';
  }

  html += '<div class="summary-stats">';
  html += '<div class="stat-item"><div class="stat-value">' + s.totalGoals + '</div><div class="stat-label">Total Goals</div></div>';
  html += '<div class="stat-item"><div class="stat-value" style="color:#1B5E20;">' + s.goalsWithAdequateOrMet + '</div><div class="stat-label">On Track</div></div>';
  html += '<div class="stat-item"><div class="stat-value" style="color:#BA1A1A;">' + s.goalsWithNoProgress + '</div><div class="stat-label">Need Attention</div></div>';

  // GPA in summary
  if (data.gpa) {
    html += '<div class="stat-item"><div class="stat-value">' + data.gpa.rounded + '</div><div class="stat-label">Current GPA</div></div>';
  } else {
    html += '<div class="stat-item"><div class="stat-value">N/A</div><div class="stat-label">Current GPA</div></div>';
  }
  html += '</div></div>';

  // ── Goals Section ──
  html += '<h2>Goals &amp; Objectives</h2>';

  if (data.goalGroups.length === 0) {
    html += '<p style="color:#49454F;">No IEP goals have been entered for this student.</p>';
  }

  data.goalGroups.forEach(function(group) {
    html += '<div class="goal-area-section">';
    html += '<h3>' + escHtml_(group.goalArea) + '</h3>';

    group.goals.forEach(function(goal) {
      html += '<div class="goal-block">';
      html += '<p class="goal-text">' + escHtml_(goal.text) + '</p>';

      goal.objectives.forEach(function(obj) {
        html += '<div class="objective-row">';
        html += '<p><strong>' + escHtml_(obj.text) + '</strong></p>';

        // Current rating badge
        var r = obj.currentProgress.rating;
        html += '<span class="rating-badge" style="background:' + ratingBg(r) + ';color:' + ratingColor(r) + ';">' + escHtml_(r) + '</span>';

        // Student-friendly label
        var friendly = PROGRESS_FRIENDLY_LABELS[r] || '';
        if (friendly) {
          html += '<span class="friendly-label">' + escHtml_(friendly) + '</span>';
        }

        // Anecdotal notes
        if (obj.currentProgress.notes) {
          html += '<p class="anecdotal-notes">' + escHtml_(obj.currentProgress.notes) + '</p>';
        }

        // Progress timeline (prior quarters)
        if (obj.progressHistory.length > 0) {
          html += '<div class="progress-timeline">Progress: ';
          obj.progressHistory.forEach(function(h) {
            html += '<span class="timeline-chip" style="background:' + ratingBg(h.rating) + ';color:' + ratingColor(h.rating) + ';">' + h.quarter + ': ' + escHtml_(h.rating) + '</span> ';
          });
          // Current quarter
          html += '<span class="timeline-chip" style="background:' + ratingBg(r) + ';color:' + ratingColor(r) + ';font-weight:500;">' + quarter + ': ' + escHtml_(r) + '</span>';
          html += '</div>';
        }

        html += '</div>'; // objective-row
      });

      html += '</div>'; // goal-block
    });

    html += '</div>'; // goal-area-section
  });

  // ── Grades Section ──
  html += '<h2>Current Grades</h2>';

  if (data.grades.length === 0) {
    html += '<p style="color:#49454F;">Grades not yet available.</p>';
  } else {
    html += '<table class="grades-table">';
    html += '<thead><tr><th>Class</th><th>Grade</th><th>Missing Assignments</th></tr></thead>';
    html += '<tbody>';

    var hasPassFail = false;
    data.grades.forEach(function(g) {
      var isPassFail = !GPA_MAP.hasOwnProperty(g.grade) && g.grade;
      if (isPassFail) hasPassFail = true;
      html += '<tr>';
      html += '<td>' + escHtml_(g.className) + '</td>';
      html += '<td>' + escHtml_(g.grade) + (isPassFail ? ' <em style="font-size:9pt;color:#666;">(Pass/Fail)</em>' : '') + '</td>';
      html += '<td>' + g.missing + '</td>';
      html += '</tr>';
    });
    html += '</tbody></table>';

    // GPA
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

  // ── Footer ──
  html += '<div class="report-footer">';
  html += '<p><strong>Questions?</strong> Contact ' + escHtml_(s.caseManager) + '</p>';
  html += '<p>This progress report is part of the IEP process under IDEA. For questions about your child\'s IEP, please contact the case manager listed above.</p>';
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
  if (!ss.getSheetByName(SHEET_STUDENTS) || !ss.getSheetByName(SHEET_CHECKINS) || !ss.getSheetByName(SHEET_COTEACHERS) || !ss.getSheetByName(SHEET_IEP_MEETINGS)) {
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
