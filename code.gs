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
  'createdAt','updatedAt','iepGoal','caseManagerEmail'
];
var CHECKIN_HEADERS = [
  'id','studentId','weekOf',
  'planningRating','followThroughRating','regulationRating',
  'focusGoalRating','effortRating',
  'whatWentWell','barrier',
  'microGoal','microGoalCategory',
  'teacherNotes','academicDataJson','createdAt'
];
var COTEACHER_HEADERS = ['email', 'role', 'addedAt'];

function initializeSheets() {
  const ss = getSS_();

  let studentsSheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!studentsSheet) {
    studentsSheet = ss.insertSheet(SHEET_STUDENTS);
  }
  ensureHeaders_(studentsSheet, STUDENT_HEADERS);

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

  if (profile.id) {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const colIdx = {};
    headers.forEach(function(h, i) { colIdx[h] = i + 1; });

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === profile.id) {
        sheet.getRange(i+1, colIdx['firstName']).setValue(profile.firstName || '');
        sheet.getRange(i+1, colIdx['lastName']).setValue(profile.lastName || '');
        sheet.getRange(i+1, colIdx['grade']).setValue(profile.grade || '');
        sheet.getRange(i+1, colIdx['period']).setValue(profile.period || '');
        sheet.getRange(i+1, colIdx['focusGoal']).setValue(profile.focusGoal || '');
        sheet.getRange(i+1, colIdx['accommodations']).setValue(profile.accommodations || '');
        sheet.getRange(i+1, colIdx['notes']).setValue(profile.notes || '');
        sheet.getRange(i+1, colIdx['classesJson']).setValue(classesJson);
        sheet.getRange(i+1, colIdx['iepGoal']).setValue(profile.iepGoal || '');
        sheet.getRange(i+1, colIdx['caseManagerEmail']).setValue(profile.caseManagerEmail || '');
        sheet.getRange(i+1, colIdx['updatedAt']).setValue(now);
        invalidateCache_();
        return { success: true, id: profile.id };
      }
    }
  }

  const id = Utilities.getUuid();
  sheet.appendRow([
    id, profile.firstName||'', profile.lastName||'',
    profile.grade||'', profile.period||'',
    profile.focusGoal||'', profile.accommodations||'',
    profile.notes||'', classesJson, now, now,
    profile.iepGoal||'', profile.caseManagerEmail||''
  ]);
  invalidateCache_();
  return { success: true, id: id };
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
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const colIdx = {};
    headers.forEach(function(h, i) { colIdx[h] = i + 1; });

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.id) {
        sheet.getRange(i+1, colIdx['weekOf']).setValue(data.weekOf || '');
        sheet.getRange(i+1, colIdx['planningRating']).setValue(data.planningRating || '');
        sheet.getRange(i+1, colIdx['followThroughRating']).setValue(data.followThroughRating || '');
        sheet.getRange(i+1, colIdx['regulationRating']).setValue(data.regulationRating || '');
        sheet.getRange(i+1, colIdx['focusGoalRating']).setValue(data.focusGoalRating || '');
        sheet.getRange(i+1, colIdx['effortRating']).setValue(data.effortRating || '');
        sheet.getRange(i+1, colIdx['whatWentWell']).setValue(data.whatWentWell || '');
        sheet.getRange(i+1, colIdx['barrier']).setValue(data.barrier || '');
        sheet.getRange(i+1, colIdx['microGoal']).setValue(data.microGoal || '');
        sheet.getRange(i+1, colIdx['microGoalCategory']).setValue(data.microGoalCategory || '');
        sheet.getRange(i+1, colIdx['teacherNotes']).setValue(data.teacherNotes || '');
        sheet.getRange(i+1, colIdx['academicDataJson']).setValue(academicJson);
        invalidateCache_();
        return { success: true, id: data.id };
      }
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
    data.teacherNotes||'', academicJson, now
  ]);
  invalidateCache_();
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
    if (data[i][0] === checkInId) { sheet.deleteRow(i + 1); invalidateCache_(); return { success: true }; }
  }
  return { success: false };
}

function updateCheckInAcademicData(checkInId, academicData) {
  initializeSheetsIfNeeded_();
  const ss = getSS_();
  const sheet = ss.getSheetByName(SHEET_CHECKINS);
  if (!sheet) return { success: false };
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  var colIdx = -1;
  headers.forEach(function(h, i) { if (h === 'academicDataJson') colIdx = i + 1; });
  if (colIdx < 0) return { success: false };

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === checkInId) {
      sheet.getRange(i + 1, colIdx).setValue(JSON.stringify(academicData || []));
      invalidateCache_();
      return { success: true };
    }
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
      iepGoal: s.iepGoal || '',
      caseManagerEmail: s.caseManagerEmail || '',
      classes: s.classes || [],
      totalCheckIns: totalCheckIns,
      latestCheckInId: latest ? latest.id : null,
      latestWeek: latest ? latest.weekOf : null,
      latestMicroGoal: latest ? latest.microGoal : null,
      avgRating: avgRating,
      trend: trend,
      gpa: gpa,
      totalMissing: totalMissing,
      academicData: academicData
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
  } catch(e) {}
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
    // Normalize legacy 'owner' role to 'caseload-manager'
    if (String(member.role).toLowerCase() === 'owner') {
      member.role = 'caseload-manager';
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
      var role = String(data[i][1]).toLowerCase();
      // Normalize legacy 'owner' role to 'caseload-manager'
      if (role === 'owner') role = 'caseload-manager';
      return role;
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
  if (!ss.getSheetByName(SHEET_STUDENTS) || !ss.getSheetByName(SHEET_CHECKINS) || !ss.getSheetByName(SHEET_COTEACHERS)) {
    initializeSheets();
  }
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

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colIdx = {};
  headers.forEach(function(h, i) { colIdx[h] = i + 1; });

  if (!colIdx['caseManagerEmail']) {
    return { success: false, error: 'caseManagerEmail column not found. Please reload.' };
  }

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === studentId) {
      sheet.getRange(i + 1, colIdx['caseManagerEmail']).setValue(caseManagerEmail || '');
      invalidateCache_();
      return { success: true };
    }
  }
  return { success: false, error: 'Student not found.' };
}
