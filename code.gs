/* ============================================================
   EF Weekly Check-In — Google Apps Script Backend
   Storage: Google Sheets (standalone w/ hardcoded ID)
   v2 — adds academic snapshot (classes, grades, missing)
   ============================================================ */

// ───── Constants ─────
const SPREADSHEET_ID = '1VPFDlW87UstmxwQu8PL5t1wS_cLY3A1iS65Xxl2Ol1E';
const SHEET_STUDENTS = 'Students';
const SHEET_CHECKINS = 'CheckIns';

const DEFAULT_STUDENTS = [
  'Antoniah Cureton',
  'Eric Evans',
  'Ian Gutierrez Hernandez',
  'Xander Hanna',
  'Terrance Jones',
  'Roman Lussier',
  'Alex Panayotov',
  'Leo Paradise',
  'Lucas Thielen'
];

const GPA_MAP = {
  'A':4.0, 'A-':3.7,
  'B+':3.3, 'B':3.0, 'B-':2.7,
  'C+':2.3, 'C':2.0, 'C-':1.7,
  'D+':1.3, 'D':1.0, 'D-':0.7,
  'F':0.0
};

/** Get spreadsheet — works standalone or bound */
function getSS_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch(e) {}
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ───── Web App Entry ─────
function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('Richfield Public Schools | EF Check-In')
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
  'createdAt','updatedAt','iepGoal'
];
var CHECKIN_HEADERS = [
  'id','studentId','weekOf',
  'planningRating','followThroughRating','regulationRating',
  'focusGoalRating','effortRating',
  'whatWentWell','barrier',
  'microGoal','microGoalCategory',
  'teacherNotes','academicDataJson','createdAt'
];

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

  if (studentsSheet.getLastRow() <= 1) {
    seedDefaultStudents_(studentsSheet);
  }

  return { success: true };
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

function seedDefaultStudents_(sheet) {
  const now = new Date().toISOString();
  DEFAULT_STUDENTS.forEach(function(fullName) {
    const parts = fullName.split(' ');
    const firstName = parts[0];
    const lastName = parts.slice(1).join(' ');
    const id = Utilities.getUuid();
    sheet.appendRow([id, firstName, lastName, '', '', '', '', '', '[]', now, now, '']);
  });
}

function forceSeedStudents() {
  var ss = getSS_();
  var sheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheet) { initializeSheets(); return; }
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
  }
  seedDefaultStudents_(sheet);
  Logger.log('Seeded ' + DEFAULT_STUDENTS.length + ' students');
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
    profile.iepGoal||''
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
      classes: s.classes || [],
      totalCheckIns: totalCheckIns,
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

// ───── Script Properties Cache ─────
// Sheets = source of truth; Script Properties = fast read cache.
// Pattern: read-through cache, invalidate on write.

var CACHE_PREFIX = 'cache_';

function getCache_(key) {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(CACHE_PREFIX + key);
    if (raw) return JSON.parse(raw);
  } catch(e) {}
  return null;
}

function setCache_(key, data) {
  try {
    PropertiesService.getScriptProperties().setProperty(CACHE_PREFIX + key, JSON.stringify(data));
  } catch(e) { /* exceeds 9KB property limit — skip silently */ }
}

function invalidateCache_() {
  try {
    var props = PropertiesService.getScriptProperties();
    props.deleteProperty(CACHE_PREFIX + 'students');
    props.deleteProperty(CACHE_PREFIX + 'dashboard');
  } catch(e) {}
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
  if (!ss.getSheetByName(SHEET_STUDENTS) || !ss.getSheetByName(SHEET_CHECKINS)) {
    initializeSheets();
  }
}

// ───── Menu ─────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🧠 EF Check-In')
    .addItem('Open Check-In App', 'openWebApp')
    .addItem('Initialize / Reset Sheets', 'initializeSheets')
    .addItem('Force Seed Students', 'forceSeedStudents')
    .addItem('Get Web App URL', 'showWebAppUrl')
    .addToUi();
}

function openWebApp() {
  const html = HtmlService.createHtmlOutput(
    '<p style="font-family:sans-serif;">Opening…</p>' +
    '<script>window.open("' + ScriptApp.getService().getUrl() + '");google.script.host.close();</script>'
  ).setWidth(300).setHeight(80);
  SpreadsheetApp.getUi().showModalDialog(html, 'EF Check-In');
}

function showWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  const html = HtmlService.createHtmlOutput(
    '<p style="font-family:sans-serif;margin-bottom:12px;">Your webapp URL:</p>' +
    '<input type="text" value="' + url + '" style="width:100%;padding:8px;font-size:13px;" onclick="this.select()" readonly>'
  ).setWidth(450).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Web App URL');
}
