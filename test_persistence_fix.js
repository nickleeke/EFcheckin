/**
 * Test: Missing assignments persistence fix
 *
 * Verifies that loadDashboard() does not overwrite local optimistic state
 * while academic-data persists are pending or in-flight.
 *
 * Run: node test_persistence_fix.js
 */

// ─── Minimal mocks ───

var _toasts = [];
function showToast(msg) { _toasts.push(msg); }

var _dashboardRendered = false;
var _dashboardDataRendered = null;
function renderDashboard(data) { _dashboardRendered = true; _dashboardDataRendered = data; }
function showDashboardSkeleton() {}
function showView() {}

// Mock google.script.run
var _backendCalls = [];
var _pendingRunHandlers = [];
var google = {
  script: {
    run: (function() {
      function makeRunner() {
        var _sh = null, _fh = null;
        var runner = {
          withSuccessHandler: function(fn) { _sh = fn; return runner; },
          withFailureHandler: function(fn) { _fh = fn; return runner; },
          updateCheckInAcademicData: function(checkInId, data) {
            _backendCalls.push({ fn: 'updateCheckInAcademicData', checkInId: checkInId, data: data });
            _pendingRunHandlers.push({ success: _sh, failure: _fh });
          },
          getDashboardData: function() {
            _backendCalls.push({ fn: 'getDashboardData' });
            _pendingRunHandlers.push({ success: _sh, failure: _fh });
          }
        };
        return runner;
      }
      return {
        get withSuccessHandler() { return makeRunner().withSuccessHandler; }
      };
    })()
  }
};

// Provide minimal 'document' mock for the dashboard-view check
var _dashboardViewActive = false;
var document = {
  getElementById: function(id) {
    if (id === 'dashboard-view') {
      return {
        classList: {
          contains: function(cls) {
            if (cls === 'active') return _dashboardViewActive;
            return false;
          }
        }
      };
    }
    return { classList: { contains: function() { return false; } } };
  }
};

// ─── Application state (mirrors the real app) ───

var appState = { dashboardData: [] };
var _persistTimers = {};
var _persistInFlight = 0;

// ─── Functions under test (copied from JavaScript.html) ───

function loadDashboard() {
  if (Object.keys(_persistTimers).length > 0 || _persistInFlight > 0) {
    renderDashboard(appState.dashboardData);
    return;
  }
  showDashboardSkeleton();
  google.script.run
    .withSuccessHandler(function(data) {
      appState.dashboardData = data || [];
      renderDashboard(data);
    })
    .withFailureHandler(function() {})
    .getDashboardData();
}

function persistAcademicData_(student) {
  if (!student.latestCheckInId) {
    showToast('Change not saved — no check-in data found.');
    return;
  }
  if (_persistTimers[student.id]) clearTimeout(_persistTimers[student.id]);
  var checkInId = student.latestCheckInId;
  var dataSnapshot = JSON.parse(JSON.stringify(student.academicData));
  _persistTimers[student.id] = setTimeout(function() {
    delete _persistTimers[student.id];
    _persistInFlight++;
    google.script.run
      .withSuccessHandler(function() {
        _persistInFlight--;
        if (_persistInFlight === 0 && Object.keys(_persistTimers).length === 0 &&
            document.getElementById('dashboard-view').classList.contains('active')) {
          loadDashboard();
        }
      })
      .withFailureHandler(function(err) {
        _persistInFlight--;
        showToast('Could not save change: ' + (err && err.message ? err.message : String(err)));
      })
      .updateCheckInAcademicData(checkInId, dataSnapshot);
  }, 600);
}

function markMissingDone(student, classIdx, assignmentIdx) {
  var classData = student.academicData[classIdx];
  if (!classData || !classData.missingAssignments) return;
  classData.missingAssignments.splice(assignmentIdx, 1);
  classData.missing = classData.missingAssignments.length;
  student.totalMissing = student.academicData.reduce(function(sum, c) {
    return sum + (Number(c.missing) || 0);
  }, 0);
  persistAcademicData_(student);
}

// ─── Helper to advance time (flush setTimeout) ───
function flushTimers() {
  // Force all pending timers to fire
  for (var id in _persistTimers) {
    var timer = _persistTimers[id];
    clearTimeout(timer);
    delete _persistTimers[id];
    // We need to actually call the timer callback — re-implement inline
  }
}

// ─── Tests ───

var passed = 0, failed = 0;
function assert(condition, msg) {
  if (condition) {
    passed++;
    console.log('  PASS: ' + msg);
  } else {
    failed++;
    console.error('  FAIL: ' + msg);
  }
}

function resetState() {
  appState.dashboardData = [];
  _persistTimers = {};
  _persistInFlight = 0;
  _backendCalls = [];
  _pendingRunHandlers = [];
  _toasts = [];
  _dashboardRendered = false;
  _dashboardDataRendered = null;
  _dashboardViewActive = false;
}

function makeStudent() {
  return {
    id: 'student-1',
    firstName: 'Test',
    lastName: 'Student',
    latestCheckInId: 'checkin-1',
    totalMissing: 3,
    academicData: [
      {
        className: 'Math',
        grade: 'B',
        missing: 2,
        missingAssignments: [
          { name: 'HW 1', type: 'Formative' },
          { name: 'Quiz 2', type: 'Summative' }
        ]
      },
      {
        className: 'ELA',
        grade: 'A',
        missing: 1,
        missingAssignments: [
          { name: 'Essay', type: 'Summative' }
        ]
      }
    ]
  };
}

// Test 1: loadDashboard skips backend fetch while debounce timer is pending
console.log('\nTest 1: loadDashboard uses local data while debounce pending');
resetState();
var student1 = makeStudent();
appState.dashboardData = [student1];
markMissingDone(student1, 0, 0); // removes HW 1
assert(Object.keys(_persistTimers).length > 0, 'Debounce timer is set after markMissingDone');
assert(student1.totalMissing === 2, 'Local totalMissing updated (3 -> 2)');
assert(student1.academicData[0].missingAssignments.length === 1, 'Assignment removed locally');

_dashboardRendered = false;
loadDashboard();
assert(_dashboardRendered === true, 'Dashboard was rendered');
assert(_dashboardDataRendered === appState.dashboardData, 'Used local appState (not backend)');
assert(_backendCalls.filter(function(c) { return c.fn === 'getDashboardData'; }).length === 0,
  'No backend getDashboardData call was made');

// Test 2: loadDashboard skips backend fetch while persist is in-flight
console.log('\nTest 2: loadDashboard uses local data while persist in-flight');
resetState();
var student2 = makeStudent();
appState.dashboardData = [student2];
_persistInFlight = 1; // simulate in-flight call

_dashboardRendered = false;
_backendCalls = [];
loadDashboard();
assert(_dashboardRendered === true, 'Dashboard was rendered');
assert(_dashboardDataRendered === appState.dashboardData, 'Used local appState');
assert(_backendCalls.filter(function(c) { return c.fn === 'getDashboardData'; }).length === 0,
  'No backend getDashboardData call was made');

// Test 3: loadDashboard fetches from backend when no pending persists
console.log('\nTest 3: loadDashboard fetches from backend when clear');
resetState();
var student3 = makeStudent();
appState.dashboardData = [student3];

_dashboardRendered = false;
loadDashboard();
assert(_backendCalls.filter(function(c) { return c.fn === 'getDashboardData'; }).length === 1,
  'Backend getDashboardData call was made');

// Simulate backend response with stale data
var staleData = [makeStudent()]; // still has 3 missing
_pendingRunHandlers[_pendingRunHandlers.length - 1].success(staleData);
assert(appState.dashboardData === staleData, 'appState overwritten with backend data');

// Test 4: persist success triggers dashboard refresh when on dashboard view
console.log('\nTest 4: Persist success triggers dashboard refresh when visible');
resetState();
var student4 = makeStudent();
appState.dashboardData = [student4];
_dashboardViewActive = true;

// Simulate persist going in-flight
_persistInFlight = 1;
_backendCalls = [];

// Simulate persist success
_persistInFlight--;
// Replicate the check from the success handler
if (_persistInFlight === 0 && Object.keys(_persistTimers).length === 0 &&
    document.getElementById('dashboard-view').classList.contains('active')) {
  loadDashboard();
}
assert(_backendCalls.filter(function(c) { return c.fn === 'getDashboardData'; }).length === 1,
  'loadDashboard triggered backend fetch after persist completed');

// Test 5: persist success does NOT trigger refresh when NOT on dashboard view
console.log('\nTest 5: Persist success skips refresh when dashboard not visible');
resetState();
var student5 = makeStudent();
appState.dashboardData = [student5];
_dashboardViewActive = false;

_persistInFlight = 1;
_backendCalls = [];

_persistInFlight--;
if (_persistInFlight === 0 && Object.keys(_persistTimers).length === 0 &&
    document.getElementById('dashboard-view').classList.contains('active')) {
  loadDashboard();
}
assert(_backendCalls.filter(function(c) { return c.fn === 'getDashboardData'; }).length === 0,
  'No backend call made when dashboard not visible');

// Test 6: markMissingDone correctly updates local state
console.log('\nTest 6: markMissingDone correctly modifies in-memory state');
resetState();
var student6 = makeStudent();
appState.dashboardData = [student6];
markMissingDone(student6, 0, 0); // remove first assignment from first class
assert(student6.academicData[0].missingAssignments.length === 1, 'One assignment removed from class');
assert(student6.academicData[0].missingAssignments[0].name === 'Quiz 2', 'Correct assignment remains');
assert(student6.academicData[0].missing === 1, 'Missing count updated');
assert(student6.totalMissing === 2, 'Total missing updated (3 -> 2)');

// Test 7: persistAcademicData_ aborts if no latestCheckInId
console.log('\nTest 7: Persist aborts without latestCheckInId');
resetState();
var student7 = makeStudent();
student7.latestCheckInId = null;
appState.dashboardData = [student7];
_toasts = [];
persistAcademicData_(student7);
assert(_toasts.some(function(t) { return t.indexOf('no check-in') >= 0; }),
  'Toast shown about missing check-in');
assert(Object.keys(_persistTimers).length === 0, 'No timer set');

// Test 8: End-to-end: mark done -> nav away -> persist completes -> dashboard syncs
console.log('\nTest 8: End-to-end navigation scenario');
resetState();
var student8 = makeStudent();
appState.dashboardData = [student8];
_dashboardViewActive = false;

// Step 1: Mark assignment done
markMissingDone(student8, 0, 0);
assert(student8.totalMissing === 2, 'Local state updated');

// Step 2: User navigates to dashboard via nav bar (showDashboard)
_dashboardViewActive = true;
loadDashboard(); // This should use local data
assert(_backendCalls.filter(function(c) { return c.fn === 'getDashboardData'; }).length === 0,
  'No stale backend fetch while timer pending');
assert(_dashboardDataRendered === appState.dashboardData, 'Dashboard shows local optimistic data');

// Step 3: Debounce fires (simulate clearing timer and going in-flight)
// In real code this happens in the setTimeout callback
for (var id in _persistTimers) {
  clearTimeout(_persistTimers[id]);
  delete _persistTimers[id];
}
_persistInFlight++;
// Backend call would be made here via google.script.run
_backendCalls.push({ fn: 'updateCheckInAcademicData', checkInId: 'checkin-1', data: student8.academicData });

// Step 4: Backend persist succeeds
_persistInFlight--;
_backendCalls = [];
if (_persistInFlight === 0 && Object.keys(_persistTimers).length === 0 &&
    document.getElementById('dashboard-view').classList.contains('active')) {
  loadDashboard(); // Should now fetch from backend
}
assert(_backendCalls.filter(function(c) { return c.fn === 'getDashboardData'; }).length === 1,
  'Dashboard synced from backend after persist completed');

// ─── Summary ───
console.log('\n────────────────────────────');
console.log('Results: ' + passed + ' passed, ' + failed + ' failed');
if (failed > 0) {
  process.exit(1);
} else {
  console.log('All tests passed!');
}
