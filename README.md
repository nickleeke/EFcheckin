# Caseload Dashboard

## Implementing Multi-User Capabilities

The app currently operates as a single-tenant system — one shared spreadsheet, no authentication, and all data visible to anyone with the web app URL. The following guide describes how to extend it to support multiple teachers, each with their own students and check-in data.

### 1. Enable User Identification

Google Apps Script can identify the signed-in Google user automatically. In `code.gs`, use the `Session` API to get the current user's email:

```javascript
function getCurrentUserEmail_() {
  return Session.getActiveUser().getEmail();
}
```

**Deployment requirement:** For `Session.getActiveUser()` to return a value, the web app must be deployed with **"Execute as: User accessing the web app"** (not "Me"). Update this in the Apps Script editor under **Deploy > Manage deployments > Configuration**.

### 2. Add a Teachers Sheet

Create a new `Teachers` sheet in the spreadsheet to register authorized users and their roles. Add these constants and headers to `code.gs`:

```javascript
const SHEET_TEACHERS = 'Teachers';

var TEACHER_HEADERS = [
  'email', 'displayName', 'role', 'createdAt'
];
```

| Column        | Purpose                                        |
|---------------|------------------------------------------------|
| `email`       | Google account email (primary key)             |
| `displayName` | Teacher's name for display in the UI           |
| `role`        | `admin` (sees all students) or `teacher` (sees own students only) |
| `createdAt`   | Timestamp of when the teacher was registered   |

Initialize the sheet inside `initializeSheets()` alongside the existing Students and CheckIns sheets:

```javascript
let teachersSheet = ss.getSheetByName(SHEET_TEACHERS);
if (!teachersSheet) {
  teachersSheet = ss.insertSheet(SHEET_TEACHERS);
}
ensureHeaders_(teachersSheet, TEACHER_HEADERS);
```

### 3. Map Students to Teachers

Add a `teacherEmail` column to the `Students` sheet to track which teacher owns each student. Update the `STUDENT_HEADERS` array:

```javascript
var STUDENT_HEADERS = [
  'id','firstName','lastName','grade','period',
  'focusGoal','accommodations','notes','classesJson',
  'createdAt','updatedAt','iepGoal','teacherEmail'
];
```

When a teacher creates or edits a student, stamp their email on the record:

```javascript
// Inside saveStudent(), when building a new row:
const email = getCurrentUserEmail_();
sheet.appendRow([
  id, profile.firstName||'', profile.lastName||'',
  profile.grade||'', profile.period||'',
  profile.focusGoal||'', profile.accommodations||'',
  profile.notes||'', classesJson, now, now,
  profile.iepGoal||'', email
]);
```

### 4. Filter Data by Current User

Modify `getStudents()` and `getDashboardData()` to return only the students belonging to the logged-in teacher (or all students for admins):

```javascript
function getStudents() {
  initializeSheetsIfNeeded_();
  const ss = getSS_();
  const sheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const email = getCurrentUserEmail_();
  const isAdmin = checkIsAdmin_(email);

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('teacherEmail');
  const students = [];

  for (let i = 1; i < data.length; i++) {
    // Skip students that don't belong to this teacher
    if (!isAdmin && emailCol >= 0 && data[i][emailCol] !== email) continue;

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

  return students;
}
```

Add a helper to check admin status:

```javascript
function checkIsAdmin_(email) {
  const ss = getSS_();
  const sheet = ss.getSheetByName(SHEET_TEACHERS);
  if (!sheet || sheet.getLastRow() <= 1) return false;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email && data[i][2] === 'admin') return true;
  }
  return false;
}
```

Apply the same per-user filtering pattern to `getDashboardData()` by filtering the students list before building summaries.

### 5. Per-User Caching

The current cache uses global keys (`cache_students`, `cache_dashboard`). With multiple users seeing different data, scope cache keys by email:

```javascript
function getUserCacheKey_(base) {
  return CACHE_PREFIX + getCurrentUserEmail_() + '_' + base;
}
```

Replace the `getCache_('students')` / `setCache_('students', ...)` calls with `getCache_(getUserCacheKey_('students'))` and so on. Update `invalidateCache_()` to clear entries for all users, or switch to `CacheService.getUserCache()` which is scoped per user automatically:

```javascript
function getCache_(key) {
  try {
    var raw = CacheService.getUserCache().get(key);
    if (raw) return JSON.parse(raw);
  } catch(e) {}
  return null;
}

function setCache_(key, data) {
  try {
    CacheService.getUserCache().put(key, JSON.stringify(data), 300);
  } catch(e) {}
}

function invalidateCache_() {
  try {
    var cache = CacheService.getUserCache();
    cache.remove('students');
    cache.remove('dashboard');
  } catch(e) {}
}
```

### 6. Guard Write Operations

Protect `saveStudent()`, `deleteStudent()`, `saveCheckIn()`, and `deleteCheckIn()` so teachers can only modify their own students. Add an ownership check:

```javascript
function assertOwnership_(studentId) {
  const email = getCurrentUserEmail_();
  if (checkIsAdmin_(email)) return;

  const ss = getSS_();
  const sheet = ss.getSheetByName(SHEET_STUDENTS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('teacherEmail');

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === studentId) {
      if (data[i][emailCol] !== email) {
        throw new Error('You do not have permission to modify this student.');
      }
      return;
    }
  }
  throw new Error('Student not found.');
}
```

Call `assertOwnership_(profile.id)` at the top of `saveStudent()` (for updates), `deleteStudent()`, `saveCheckIn()`, and `deleteCheckIn()`.

### 7. Frontend Changes

On the frontend in `JavaScript.html`, display the current user's identity and pass it through where needed:

```javascript
// On app load, fetch and display the current user
google.script.run
  .withSuccessHandler(function(email) {
    document.getElementById('current-user').textContent = email;
  })
  .getCurrentUserEmail_();
```

No changes are needed when calling `getStudents()` or `getDashboardData()` — the backend now filters automatically based on the authenticated session.

### 8. Migration Checklist

When rolling out multi-user support to an existing single-user deployment:

1. **Add the `Teachers` sheet** — Run `initializeSheets()` after updating the code, or manually create the sheet with the headers `email | displayName | role | createdAt`.
2. **Register the first admin** — Add a row to the `Teachers` sheet with the deploying teacher's Google email and `admin` as the role.
3. **Backfill `teacherEmail`** — For existing student records that have no `teacherEmail` value, assign them to the appropriate teacher by filling in the column in the spreadsheet directly.
4. **Redeploy the web app** — In the Apps Script editor, go to **Deploy > Manage deployments**, change "Execute as" to **"User accessing the web app"**, set "Who has access" to **"Anyone with Google account"** (or your organization domain), and create a new deployment version.
5. **Share the spreadsheet** — Give each teacher **Editor** access to the underlying Google Sheet so Apps Script can read/write on their behalf.
6. **Test with a second account** — Log in with a non-admin teacher account and verify they only see their own students.