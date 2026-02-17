/* ============================================================
   Tests.gs — IEP Quarterly Progress Reporting
   GAS-compatible function-based test suite.
   Run via Apps Script editor: select function → Run.

   Convention: test_featureName_expectedBehavior()
   Assertion helper: assert_(condition, message)
   ============================================================ */

// ───── Test Helpers ─────

function assert_(condition, message) {
  if (!condition) throw new Error('FAIL: ' + message);
}

function assertEqual_(actual, expected, message) {
  if (actual !== expected) {
    throw new Error('FAIL: ' + message + ' — expected ' + JSON.stringify(expected) + ', got ' + JSON.stringify(actual));
  }
}

function assertNotNull_(value, message) {
  if (value == null) throw new Error('FAIL: ' + message + ' — got null/undefined');
}

function assertNull_(value, message) {
  if (value != null) throw new Error('FAIL: ' + message + ' — expected null, got ' + JSON.stringify(value));
}

function assertThrows_(fn, message) {
  try { fn(); throw new Error('FAIL: ' + message + ' — expected error but none thrown'); }
  catch(e) { if (e.message.indexOf('FAIL:') === 0) throw e; }
}

function assertContains_(str, substring, message) {
  if (String(str).indexOf(substring) === -1) {
    throw new Error('FAIL: ' + message + ' — "' + substring + '" not found in output');
  }
}

function assertNotContains_(str, substring, message) {
  if (String(str).indexOf(substring) !== -1) {
    throw new Error('FAIL: ' + message + ' — "' + substring + '" should not appear in output');
  }
}

/** Create a temporary test spreadsheet, run fn(ss), then delete it. */
function withTestSpreadsheet_(fn) {
  var ss = SpreadsheetApp.create('__TEST_ProgressReport_' + Date.now());
  try {
    fn(ss);
  } finally {
    DriveApp.getFileById(ss.getId()).setTrashed(true);
  }
}

/** Build a mock student object matching the dashboard data shape. */
function buildMockStudent_(overrides) {
  var s = {
    id: 'stu-test-001',
    firstName: 'Alex',
    lastName: 'Johnson',
    grade: '9',
    period: '3',
    focusGoal: 'Improve planning skills',
    accommodations: 'Extended time on tests',
    notes: '',
    iepGoal: 'Alex will improve executive function skills.',
    goalsJson: JSON.stringify([
      {
        id: 'goal-1',
        text: 'Alex will solve multi-step math problems with 80% accuracy.',
        goalArea: 'Math Calculation',
        objectives: [
          { id: 'obj-1a', text: 'Solve 2-step equations with 80% accuracy.' },
          { id: 'obj-1b', text: 'Apply order of operations correctly in 4 out of 5 trials.' }
        ]
      },
      {
        id: 'goal-2',
        text: 'Alex will read and summarize grade-level text with 75% accuracy.',
        goalArea: 'Reading Comprehension',
        objectives: [
          { id: 'obj-2a', text: 'Identify main idea in a passage 3 out of 4 trials.' }
        ]
      }
    ]),
    caseManagerEmail: 'teacher@rpsmn.org',
    contacts: [],
    classes: [
      { name: 'Algebra 1', teacher: 'Ms. Smith' },
      { name: 'English 9', teacher: 'Mr. Brown' }
    ],
    gpa: 3.2,
    totalMissing: 2,
    academicData: [
      { className: 'Algebra 1', grade: 'B+', missing: 1 },
      { className: 'English 9', grade: 'B', missing: 1 },
      { className: 'Phy Ed', grade: 'P', missing: 0 }
    ]
  };
  for (var key in overrides) {
    if (overrides.hasOwnProperty(key)) s[key] = overrides[key];
  }
  return s;
}

// ───── Run All Tests ─────

function runAllProgressReportTests() {
  var tests = [
    // Data Layer — Progress Entry
    'test_progressEntry_savesToCorrectLocation',
    'test_progressEntry_validatesRating',
    'test_progressEntry_validatesQuarter',
    'test_progressEntry_allowsUpdateToExisting',
    'test_progressEntry_requiresAnecdotalNotes',
    'test_progressEntry_associatesWithStudent',
    // Data Layer — Grades & GPA
    'test_grades_retrieveCurrentGrades',
    'test_gpa_calculatesCorrectly',
    'test_gpa_handlesNoGrades',
    // Report Data Assembly
    'test_reportAssembly_gathersAllGoals',
    'test_reportAssembly_gathersObjectivesPerGoal',
    'test_reportAssembly_includesSummaryData',
    'test_reportAssembly_includesGradesSection',
    'test_reportAssembly_handlesPartialData',
    // Printable Report Generation
    'test_printableReport_generatesHTML',
    'test_printableReport_headerSection',
    'test_printableReport_summarySection',
    'test_printableReport_goalsSection',
    'test_printableReport_gradesSection',
    'test_printableReport_footerSection',
    'test_printableReport_printStyles',
    'test_printableReport_studentFriendlyLanguage',
    // Edge Cases
    'test_edge_newStudentNoHistory',
    'test_edge_midYearGoalChange',
    'test_edge_noGradesAvailable',
    'test_edge_longAnecdotalNotes',
    'test_edge_manyObjectives',
    // Quarter Utilities
    'test_getCurrentQuarter_returnsValidQuarter',
    'test_getQuarterLabel_formatsCorrectly'
  ];

  // Build a map of test functions from the global scope
  var testFns = {
    test_progressEntry_savesToCorrectLocation: test_progressEntry_savesToCorrectLocation,
    test_progressEntry_validatesRating: test_progressEntry_validatesRating,
    test_progressEntry_validatesQuarter: test_progressEntry_validatesQuarter,
    test_progressEntry_allowsUpdateToExisting: test_progressEntry_allowsUpdateToExisting,
    test_progressEntry_requiresAnecdotalNotes: test_progressEntry_requiresAnecdotalNotes,
    test_progressEntry_associatesWithStudent: test_progressEntry_associatesWithStudent,
    test_grades_retrieveCurrentGrades: test_grades_retrieveCurrentGrades,
    test_gpa_calculatesCorrectly: test_gpa_calculatesCorrectly,
    test_gpa_handlesNoGrades: test_gpa_handlesNoGrades,
    test_reportAssembly_gathersAllGoals: test_reportAssembly_gathersAllGoals,
    test_reportAssembly_gathersObjectivesPerGoal: test_reportAssembly_gathersObjectivesPerGoal,
    test_reportAssembly_includesSummaryData: test_reportAssembly_includesSummaryData,
    test_reportAssembly_includesGradesSection: test_reportAssembly_includesGradesSection,
    test_reportAssembly_handlesPartialData: test_reportAssembly_handlesPartialData,
    test_printableReport_generatesHTML: test_printableReport_generatesHTML,
    test_printableReport_headerSection: test_printableReport_headerSection,
    test_printableReport_summarySection: test_printableReport_summarySection,
    test_printableReport_goalsSection: test_printableReport_goalsSection,
    test_printableReport_gradesSection: test_printableReport_gradesSection,
    test_printableReport_footerSection: test_printableReport_footerSection,
    test_printableReport_printStyles: test_printableReport_printStyles,
    test_printableReport_studentFriendlyLanguage: test_printableReport_studentFriendlyLanguage,
    test_edge_newStudentNoHistory: test_edge_newStudentNoHistory,
    test_edge_midYearGoalChange: test_edge_midYearGoalChange,
    test_edge_noGradesAvailable: test_edge_noGradesAvailable,
    test_edge_longAnecdotalNotes: test_edge_longAnecdotalNotes,
    test_edge_manyObjectives: test_edge_manyObjectives,
    test_getCurrentQuarter_returnsValidQuarter: test_getCurrentQuarter_returnsValidQuarter,
    test_getQuarterLabel_formatsCorrectly: test_getQuarterLabel_formatsCorrectly
  };

  var passed = 0, failed = 0, errors = [];
  tests.forEach(function(name) {
    try {
      testFns[name]();
      passed++;
      Logger.log('PASS: ' + name);
    } catch(e) {
      failed++;
      errors.push(name + ': ' + e.message);
      Logger.log('FAIL: ' + name + ' — ' + e.message);
    }
  });

  Logger.log('');
  Logger.log('Results: ' + passed + ' passed, ' + failed + ' failed');
  if (errors.length > 0) {
    Logger.log('Failures:');
    errors.forEach(function(e) { Logger.log('  ' + e); });
  }
  return { passed: passed, failed: failed, errors: errors };
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 1. DATA LAYER — Progress Entry
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function test_progressEntry_savesToCorrectLocation() {
  // When a teacher enters progress data for a goal, it saves to the ProgressReporting sheet
  var data = {
    studentId: 'stu-test-001',
    goalId: 'goal-1',
    objectiveId: 'obj-1a',
    quarter: 'Q2',
    progressRating: 'Adequate Progress',
    anecdotalNotes: 'Alex showed improvement in solving two-step equations this quarter.'
  };

  var result = saveProgressEntry(data);
  try {
    assert_(result.success, 'saveProgressEntry should return success');
    assertNotNull_(result.id, 'saveProgressEntry should return an id');

    // Verify it persisted — retrieve and check
    var entries = getProgressEntries(data.studentId, data.quarter);
    var found = entries.filter(function(e) { return e.id === result.id; });
    assertEqual_(found.length, 1, 'Entry should be retrievable after save');
    assertEqual_(found[0].goalId, 'goal-1', 'goalId should match');
    assertEqual_(found[0].objectiveId, 'obj-1a', 'objectiveId should match');
    assertEqual_(found[0].quarter, 'Q2', 'quarter should match');
    assertEqual_(found[0].progressRating, 'Adequate Progress', 'progressRating should match');
    assertNotNull_(found[0].dateEntered, 'dateEntered should be set');
    assertNotNull_(found[0].enteredBy, 'enteredBy should be set');
  } finally {
    if (result && result.id) deleteProgressEntry_(result.id);
  }
}

function test_progressEntry_validatesRating() {
  // progressRating must be one of the three valid values
  var validRatings = ['No Progress', 'Adequate Progress', 'Objective Met'];
  var createdIds = [];

  validRatings.forEach(function(rating) {
    var data = {
      studentId: 'stu-test-001',
      goalId: 'goal-1',
      objectiveId: 'obj-1a',
      quarter: 'Q1',
      progressRating: rating,
      anecdotalNotes: 'This is a valid test note for rating validation.'
    };
    var result = saveProgressEntry(data);
    assert_(result.success, 'Rating "' + rating + '" should be accepted');
    if (result.id) createdIds.push(result.id);
  });
  // Cleanup — all valid saves upsert same combo so only latest ID matters
  createdIds.forEach(function(id) { try { deleteProgressEntry_(id); } catch(e) {} });

  // Invalid rating should be rejected
  var badData = {
    studentId: 'stu-test-001',
    goalId: 'goal-1',
    objectiveId: 'obj-1a',
    quarter: 'Q1',
    progressRating: 'Excellent',
    anecdotalNotes: 'This note is long enough for validation.'
  };
  var badResult = saveProgressEntry(badData);
  assert_(!badResult.success, 'Invalid rating "Excellent" should be rejected');
  assertContains_(badResult.error, 'rating', 'Error should mention rating');
}

function test_progressEntry_validatesQuarter() {
  // quarter must be Q1-Q4
  var validQuarters = ['Q1', 'Q2', 'Q3', 'Q4'];
  var createdIds = [];
  try {
    validQuarters.forEach(function(q) {
      var data = {
        studentId: 'stu-test-001',
        goalId: 'goal-1',
        objectiveId: 'obj-1a',
        quarter: q,
        progressRating: 'Adequate Progress',
        anecdotalNotes: 'Valid quarter test note here for quarter ' + q + '.'
      };
      var result = saveProgressEntry(data);
      assert_(result.success, 'Quarter "' + q + '" should be accepted');
      if (result.id) createdIds.push(result.id);
    });

    // Invalid quarter
    var badResult = saveProgressEntry({
      studentId: 'stu-test-001',
      goalId: 'goal-1',
      objectiveId: 'obj-1a',
      quarter: 'Q5',
      progressRating: 'No Progress',
      anecdotalNotes: 'Invalid quarter test note content here.'
    });
    assert_(!badResult.success, 'Invalid quarter "Q5" should be rejected');
  } finally {
    createdIds.forEach(function(id) { try { deleteProgressEntry_(id); } catch(e) {} });
  }
}

function test_progressEntry_allowsUpdateToExisting() {
  // If progress exists for goal+objective+quarter, update should overwrite
  var data = {
    studentId: 'stu-test-001',
    goalId: 'goal-1',
    objectiveId: 'obj-1a',
    quarter: 'Q1',
    progressRating: 'No Progress',
    anecdotalNotes: 'First entry for this objective, needs more work.'
  };
  var first = saveProgressEntry(data);
  var lastId = first.id;
  try {
    assert_(first.success, 'First save should succeed');

    // Update same combo with new data
    data.progressRating = 'Adequate Progress';
    data.anecdotalNotes = 'Updated: Alex is now showing adequate progress on this objective.';
    var second = saveProgressEntry(data);
    assert_(second.success, 'Update should succeed');
    lastId = second.id;

    // Should still be exactly one entry for this combo
    var entries = getProgressEntries('stu-test-001', 'Q1');
    var matching = entries.filter(function(e) {
      return e.goalId === 'goal-1' && e.objectiveId === 'obj-1a';
    });
    assertEqual_(matching.length, 1, 'Should have exactly one entry after update');
    assertEqual_(matching[0].progressRating, 'Adequate Progress', 'Rating should be updated');
    assertNotNull_(matching[0].lastModified, 'lastModified should be set on update');
  } finally {
    if (lastId) try { deleteProgressEntry_(lastId); } catch(e) {}
  }
}

function test_progressEntry_requiresAnecdotalNotes() {
  // anecdotalNotes cannot be empty or less than 10 characters
  var baseData = {
    studentId: 'stu-test-001',
    goalId: 'goal-1',
    objectiveId: 'obj-1a',
    quarter: 'Q1',
    progressRating: 'No Progress'
  };

  // Empty string
  var r1 = saveProgressEntry(Object.assign({}, baseData, { anecdotalNotes: '' }));
  assert_(!r1.success, 'Empty anecdotalNotes should be rejected');

  // Null
  var r2 = saveProgressEntry(Object.assign({}, baseData, { anecdotalNotes: null }));
  assert_(!r2.success, 'Null anecdotalNotes should be rejected');

  // Too short (under 10 chars)
  var r3 = saveProgressEntry(Object.assign({}, baseData, { anecdotalNotes: 'Short' }));
  assert_(!r3.success, 'Notes under 10 characters should be rejected');
  assertContains_(r3.error, '10', 'Error should mention minimum length');

  // Exactly 10 chars should pass
  var r4 = saveProgressEntry(Object.assign({}, baseData, { anecdotalNotes: '1234567890' }));
  try {
    assert_(r4.success, 'Notes with exactly 10 characters should be accepted');
  } finally {
    if (r4 && r4.id) try { deleteProgressEntry_(r4.id); } catch(e) {}
  }
}

function test_progressEntry_associatesWithStudent() {
  // Progress entry is linked to the correct student and retrievable by studentId + quarter
  var data1 = {
    studentId: 'stu-aaa',
    goalId: 'goal-1',
    objectiveId: 'obj-1a',
    quarter: 'Q2',
    progressRating: 'Objective Met',
    anecdotalNotes: 'Student AAA mastered this objective completely.'
  };
  var data2 = {
    studentId: 'stu-bbb',
    goalId: 'goal-1',
    objectiveId: 'obj-1a',
    quarter: 'Q2',
    progressRating: 'No Progress',
    anecdotalNotes: 'Student BBB needs additional support on this objective.'
  };

  var r1 = saveProgressEntry(data1);
  var r2 = saveProgressEntry(data2);
  try {
    assert_(r1.success && r2.success, 'Both saves should succeed');

    // Query by student A
    var entriesA = getProgressEntries('stu-aaa', 'Q2');
    var entriesB = getProgressEntries('stu-bbb', 'Q2');

    assert_(entriesA.length >= 1, 'Student A should have entries');
    assert_(entriesB.length >= 1, 'Student B should have entries');

    var foundA = entriesA.some(function(e) { return e.studentId === 'stu-aaa'; });
    var foundB = entriesB.some(function(e) { return e.studentId === 'stu-bbb'; });
    assert_(foundA, 'Entries for student A should have correct studentId');
    assert_(foundB, 'Entries for student B should have correct studentId');

    // Cross-check: student A entries should not contain student B data
    var crossCheck = entriesA.filter(function(e) { return e.studentId === 'stu-bbb'; });
    assertEqual_(crossCheck.length, 0, 'Student A query should not return student B entries');
  } finally {
    if (r1 && r1.id) try { deleteProgressEntry_(r1.id); } catch(e) {}
    if (r2 && r2.id) try { deleteProgressEntry_(r2.id); } catch(e) {}
  }
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 2. DATA LAYER — Grades & GPA
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function test_grades_retrieveCurrentGrades() {
  // Returns all current class grades for a given student from academicData
  var student = buildMockStudent_();
  var grades = getGradesForReport_(student);

  assertNotNull_(grades, 'Grades should not be null');
  assert_(Array.isArray(grades), 'Grades should be an array');
  assertEqual_(grades.length, 3, 'Should return all 3 classes');

  // Each record should have className and grade
  grades.forEach(function(g) {
    assertNotNull_(g.className, 'Each grade should have className');
    assertNotNull_(g.grade, 'Each grade should have grade');
  });

  // Should be sorted alphabetically by className
  assert_(grades[0].className <= grades[1].className,
    'Grades should be sorted alphabetically by className');
}

function test_gpa_calculatesCorrectly() {
  // GPA calculates from letter grades using standard 4.0 scale
  var student = buildMockStudent_({
    academicData: [
      { className: 'Math', grade: 'A', missing: 0 },
      { className: 'English', grade: 'B', missing: 0 },
      { className: 'Science', grade: 'C', missing: 0 },
      { className: 'PE', grade: 'P', missing: 0 }  // Pass/fail — excluded
    ]
  });

  var gpaResult = calculateGpaForReport_(student.academicData);

  // A=4.0, B=3.0, C=2.0 → average = 3.0 (PE excluded)
  assertNotNull_(gpaResult, 'GPA result should not be null');
  assertEqual_(gpaResult.raw, 3.0, 'Raw GPA should be 3.0');
  assertEqual_(gpaResult.rounded, '3.00', 'Rounded GPA should be "3.00"');
  assert_(gpaResult.excludedCount >= 1, 'Should report at least 1 excluded class');
}

function test_gpa_handlesNoGrades() {
  // If student has no grades, return null indicator rather than 0.0
  var student = buildMockStudent_({
    academicData: []
  });

  var gpaResult = calculateGpaForReport_(student.academicData);
  assertNull_(gpaResult, 'GPA should be null when no grade data exists');

  // Also test with only pass/fail classes
  var studentPF = buildMockStudent_({
    academicData: [
      { className: 'PE', grade: 'P', missing: 0 },
      { className: 'Advisory', grade: 'P', missing: 0 }
    ]
  });
  var gpaResultPF = calculateGpaForReport_(studentPF.academicData);
  assertNull_(gpaResultPF, 'GPA should be null when all classes are pass/fail');
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 3. REPORT DATA ASSEMBLY
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function test_reportAssembly_gathersAllGoals() {
  // For a given student + quarter, assembles all IEP goals grouped by goalArea
  var student = buildMockStudent_();
  var reportData = assembleReportData_(student, 'Q2', []);

  assertNotNull_(reportData, 'Report data should not be null');
  assertNotNull_(reportData.goalGroups, 'Should have goalGroups');
  assert_(Array.isArray(reportData.goalGroups), 'goalGroups should be an array');

  // Mock student has 2 goals in 2 different goalAreas
  assertEqual_(reportData.goalGroups.length, 2, 'Should have 2 goal area groups');

  // Each group should have goalArea and goals array
  reportData.goalGroups.forEach(function(group) {
    assertNotNull_(group.goalArea, 'Each group should have goalArea');
    assert_(Array.isArray(group.goals), 'Each group should have goals array');
    assert_(group.goals.length > 0, 'Each group should have at least one goal');
    group.goals.forEach(function(g) {
      assertNotNull_(g.text, 'Each goal should have text');
    });
  });
}

function test_reportAssembly_gathersObjectivesPerGoal() {
  // Each goal includes objectives with progress data
  var mockEntries = [
    { goalId: 'goal-1', objectiveId: 'obj-1a', quarter: 'Q1', progressRating: 'No Progress',
      anecdotalNotes: 'Q1: Alex struggled with two-step equations.' },
    { goalId: 'goal-1', objectiveId: 'obj-1a', quarter: 'Q2', progressRating: 'Adequate Progress',
      anecdotalNotes: 'Q2: Improvement shown with two-step equations.' }
  ];

  var student = buildMockStudent_();
  var reportData = assembleReportData_(student, 'Q2', mockEntries);

  // Find the Math Calculation group
  var mathGroup = reportData.goalGroups.filter(function(g) {
    return g.goalArea === 'Math Calculation';
  })[0];
  assertNotNull_(mathGroup, 'Should find Math Calculation group');

  var goal = mathGroup.goals[0];
  assert_(goal.objectives.length >= 2, 'Goal should have its objectives');

  // Check obj-1a has progress data
  var obj1a = goal.objectives.filter(function(o) { return o.id === 'obj-1a'; })[0];
  assertNotNull_(obj1a, 'Should find objective obj-1a');
  assertEqual_(obj1a.currentProgress.rating, 'Adequate Progress', 'Current quarter rating should be set');
  assertEqual_(obj1a.currentProgress.notes, 'Q2: Improvement shown with two-step equations.',
    'Current quarter notes should be set');

  // Check progress history (prior quarters)
  assert_(obj1a.progressHistory.length >= 1, 'Should have prior quarter history');
  assertEqual_(obj1a.progressHistory[0].quarter, 'Q1', 'History should include Q1');
  assertEqual_(obj1a.progressHistory[0].rating, 'No Progress', 'Q1 rating should match');
}

function test_reportAssembly_includesSummaryData() {
  // Summary includes student info, counts, and reporting period
  var student = buildMockStudent_();
  var mockEntries = [
    { goalId: 'goal-1', objectiveId: 'obj-1a', quarter: 'Q2', progressRating: 'Adequate Progress',
      anecdotalNotes: 'Making progress on equations this quarter.' },
    { goalId: 'goal-1', objectiveId: 'obj-1b', quarter: 'Q2', progressRating: 'Objective Met',
      anecdotalNotes: 'Mastered order of operations consistently.' },
    { goalId: 'goal-2', objectiveId: 'obj-2a', quarter: 'Q2', progressRating: 'No Progress',
      anecdotalNotes: 'Still struggling with main idea identification.' }
  ];

  var reportData = assembleReportData_(student, 'Q2', mockEntries);

  // Student info
  assertEqual_(reportData.summary.studentName, 'Alex Johnson', 'Should have full student name');
  assertEqual_(reportData.summary.gradeLevel, '9', 'Should have grade level');
  assertNotNull_(reportData.summary.caseManager, 'Should have case manager');
  assertContains_(reportData.summary.reportingPeriod, 'Q2', 'Reporting period should contain Q2');

  // Counts
  assertEqual_(reportData.summary.totalGoals, 2, 'Should count total goals');
  // Goals where ALL objectives are adequate or met
  assert_(reportData.summary.goalsWithAdequateOrMet >= 0, 'goalsWithAdequateOrMet should be a number');
  assert_(reportData.summary.goalsWithNoProgress >= 0, 'goalsWithNoProgress should be a number');
}

function test_reportAssembly_includesGradesSection() {
  // Report includes current grades and GPA sorted alphabetically
  var student = buildMockStudent_();
  var reportData = assembleReportData_(student, 'Q2', []);

  assertNotNull_(reportData.grades, 'Should have grades section');
  assert_(Array.isArray(reportData.grades), 'Grades should be an array');
  assert_(reportData.grades.length > 0, 'Should have grade entries');

  // Check sorted alphabetically by className
  for (var i = 1; i < reportData.grades.length; i++) {
    assert_(reportData.grades[i].className >= reportData.grades[i - 1].className,
      'Grades should be sorted by className');
  }

  // GPA should be present
  assertNotNull_(reportData.gpa, 'Should have GPA data');
}

function test_reportAssembly_handlesPartialData() {
  // Some goals have progress, others don't — report still generates
  var student = buildMockStudent_();
  // Only provide progress for goal-1, not goal-2
  var partialEntries = [
    { goalId: 'goal-1', objectiveId: 'obj-1a', quarter: 'Q2', progressRating: 'Adequate Progress',
      anecdotalNotes: 'Good improvement on equations this quarter.' }
  ];

  var reportData = assembleReportData_(student, 'Q2', partialEntries);

  assertNotNull_(reportData, 'Report should generate with partial data');
  assertEqual_(reportData.goalGroups.length, 2, 'Should still have both goal groups');

  // goal-2 objectives should show "Not yet reported"
  var readingGroup = reportData.goalGroups.filter(function(g) {
    return g.goalArea === 'Reading Comprehension';
  })[0];
  var obj = readingGroup.goals[0].objectives[0];
  assertEqual_(obj.currentProgress.rating, 'Not yet reported',
    'Missing progress should show "Not yet reported"');
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 4. PRINTABLE REPORT GENERATION
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function test_printableReport_generatesHTML() {
  // Produces valid HTML string with inline CSS
  var student = buildMockStudent_();
  var html = generateProgressReportHtml_(student, 'Q2', [], '');

  assertNotNull_(html, 'HTML output should not be null');
  assert_(typeof html === 'string', 'Output should be a string');
  assertContains_(html, '<html', 'Should contain html tag');
  assertContains_(html, '</html>', 'Should have closing html tag');
  assertContains_(html, '<style', 'Should contain inline styles');
  assertNotContains_(html, '<link', 'Should not reference external stylesheets');
}

function test_printableReport_headerSection() {
  // Header includes school info, student name, case manager, reporting period
  var student = buildMockStudent_();
  var html = generateProgressReportHtml_(student, 'Q2', [], '');

  assertContains_(html, 'IEP Progress Report', 'Should contain report title');
  assertContains_(html, 'Alex Johnson', 'Should contain student name');
  assertContains_(html, 'Q2', 'Should contain reporting period');
  assertContains_(html, 'teacher@rpsmn.org', 'Should contain case manager info');
}

function test_printableReport_summarySection() {
  // At-a-glance summary with goal counts and GPA
  var student = buildMockStudent_();
  var entries = [
    { goalId: 'goal-1', objectiveId: 'obj-1a', quarter: 'Q2', progressRating: 'Adequate Progress',
      anecdotalNotes: 'Good improvement this quarter on equations.' },
    { goalId: 'goal-1', objectiveId: 'obj-1b', quarter: 'Q2', progressRating: 'Objective Met',
      anecdotalNotes: 'Fully mastered order of operations.' }
  ];
  var html = generateProgressReportHtml_(student, 'Q2', entries, 'Alex is working hard this quarter.');

  assertContains_(html, 'Alex is working hard this quarter', 'Should contain teacher summary');
  // Should show GPA
  assertContains_(html, '3.', 'Should contain GPA value');
}

function test_printableReport_goalsSection() {
  // Each goal area is a distinct section with objectives and progress indicators
  var student = buildMockStudent_();
  var entries = [
    { goalId: 'goal-1', objectiveId: 'obj-1a', quarter: 'Q1', progressRating: 'No Progress',
      anecdotalNotes: 'Q1 - needs support on two-step equations.' },
    { goalId: 'goal-1', objectiveId: 'obj-1a', quarter: 'Q2', progressRating: 'Adequate Progress',
      anecdotalNotes: 'Q2 - showing improvement on equations.' }
  ];
  var html = generateProgressReportHtml_(student, 'Q2', entries, '');

  assertContains_(html, 'Math Calculation', 'Should contain goalArea heading');
  assertContains_(html, 'Reading Comprehension', 'Should contain second goalArea');
  assertContains_(html, 'Solve 2-step equations', 'Should contain objective text');
  assertContains_(html, 'Adequate Progress', 'Should contain progress rating');
  assertContains_(html, 'Q1', 'Should show progress timeline');
}

function test_printableReport_gradesSection() {
  // Table with Class, Grade, Missing columns
  var student = buildMockStudent_();
  var html = generateProgressReportHtml_(student, 'Q2', [], '');

  assertContains_(html, 'Algebra 1', 'Should list class name');
  assertContains_(html, 'English 9', 'Should list second class');
  assertContains_(html, 'B+', 'Should show letter grade');
  // GPA display
  assertContains_(html, 'GPA', 'Should label the GPA');
}

function test_printableReport_footerSection() {
  // Case manager contact and IDEA disclaimer
  var student = buildMockStudent_();
  var html = generateProgressReportHtml_(student, 'Q2', [], '');

  assertContains_(html, 'teacher@rpsmn.org', 'Footer should contain case manager email');
  assertContains_(html, 'IDEA', 'Footer should mention IDEA');
  assertContains_(html, 'IEP', 'Footer should mention IEP');
}

function test_printableReport_printStyles() {
  // Report includes print-optimized CSS
  var student = buildMockStudent_();
  var html = generateProgressReportHtml_(student, 'Q2', [], '');

  assertContains_(html, '@media print', 'Should contain print media query');
  assertContains_(html, 'break-inside', 'Should have page break rules');
  // Font sizes — check for readable sizing
  assertContains_(html, '11p', 'Should specify minimum body font size');
}

function test_printableReport_studentFriendlyLanguage() {
  // Progress ratings show student-friendly labels
  var student = buildMockStudent_();
  var entries = [
    { goalId: 'goal-1', objectiveId: 'obj-1a', quarter: 'Q2', progressRating: 'No Progress',
      anecdotalNotes: 'Still working on this skill, more practice needed.' },
    { goalId: 'goal-1', objectiveId: 'obj-1b', quarter: 'Q2', progressRating: 'Adequate Progress',
      anecdotalNotes: 'Shows steady growth toward this objective.' },
    { goalId: 'goal-2', objectiveId: 'obj-2a', quarter: 'Q2', progressRating: 'Objective Met',
      anecdotalNotes: 'Student has fully demonstrated this skill.' }
  ];
  var html = generateProgressReportHtml_(student, 'Q2', entries, '');

  assertContains_(html, "Let's keep working on this", '"No Progress" should show friendly label');
  assertContains_(html, "You're making progress", '"Adequate Progress" should show friendly label');
  assertContains_(html, 'You got it', '"Objective Met" should show friendly label');
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 5. EDGE CASES
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function test_edge_newStudentNoHistory() {
  // New student with no prior quarters — report generates with just current data
  var student = buildMockStudent_();
  var entries = [
    { goalId: 'goal-1', objectiveId: 'obj-1a', quarter: 'Q1', progressRating: 'Adequate Progress',
      anecdotalNotes: 'First quarter for this student, showing promise.' }
  ];
  var reportData = assembleReportData_(student, 'Q1', entries);

  assertNotNull_(reportData, 'Report should generate for new student');
  // Should have no prior quarter history for this objective
  var mathGroup = reportData.goalGroups.filter(function(g) {
    return g.goalArea === 'Math Calculation';
  })[0];
  var obj = mathGroup.goals[0].objectives[0];
  assertEqual_(obj.progressHistory.length, 0, 'New student should have no prior history');
  assertNotNull_(obj.currentProgress, 'Should have current quarter data');
}

function test_edge_midYearGoalChange() {
  // Goal added mid-year — only shows for quarters it existed
  var student = buildMockStudent_({
    goalsJson: JSON.stringify([
      {
        id: 'goal-1',
        text: 'Original goal from start of year.',
        goalArea: 'Math Calculation',
        objectives: [
          { id: 'obj-1a', text: 'Original objective.' }
        ]
      },
      {
        id: 'goal-new',
        text: 'New goal added in Q3.',
        goalArea: 'Writing',
        objectives: [
          { id: 'obj-new-a', text: 'New writing objective.' }
        ]
      }
    ])
  });

  // Progress only exists for Q3 — goal-new didn't exist in Q1/Q2
  var entries = [
    { goalId: 'goal-1', objectiveId: 'obj-1a', quarter: 'Q3', progressRating: 'Adequate Progress',
      anecdotalNotes: 'Continuing to work on math calculation skills.' },
    { goalId: 'goal-new', objectiveId: 'obj-new-a', quarter: 'Q3', progressRating: 'No Progress',
      anecdotalNotes: 'Just starting this new writing goal this quarter.' }
  ];

  var reportData = assembleReportData_(student, 'Q3', entries);

  // Both goals should appear in report for Q3
  assertEqual_(reportData.goalGroups.length, 2, 'Should show both goal areas');

  // New goal should have no prior history
  var writingGroup = reportData.goalGroups.filter(function(g) {
    return g.goalArea === 'Writing';
  })[0];
  var newObj = writingGroup.goals[0].objectives[0];
  assertEqual_(newObj.progressHistory.length, 0, 'New goal should have no prior history');
}

function test_edge_noGradesAvailable() {
  // No grade data — grades section shows fallback, GPA shows N/A
  var student = buildMockStudent_({
    academicData: [],
    gpa: null
  });
  var reportData = assembleReportData_(student, 'Q2', []);

  assertNull_(reportData.gpa, 'GPA should be null with no grades');
  assertEqual_(reportData.grades.length, 0, 'Grades array should be empty');

  // HTML report should still generate
  var html = generateProgressReportHtml_(student, 'Q2', [], '');
  assertNotNull_(html, 'Report should generate without grades');
  assertContains_(html, 'N/A', 'Should show N/A for GPA');
}

function test_edge_longAnecdotalNotes() {
  // Very long notes should not break layout
  var longNote = '';
  for (var i = 0; i < 100; i++) {
    longNote += 'This is a sentence about student progress. ';
  }

  var student = buildMockStudent_();
  var entries = [
    { goalId: 'goal-1', objectiveId: 'obj-1a', quarter: 'Q2', progressRating: 'Adequate Progress',
      anecdotalNotes: longNote }
  ];

  // Should save without error
  var result = saveProgressEntry({
    studentId: 'stu-test-001',
    goalId: 'goal-1',
    objectiveId: 'obj-1a',
    quarter: 'Q2',
    progressRating: 'Adequate Progress',
    anecdotalNotes: longNote
  });
  try {
    assert_(result.success, 'Long notes should save successfully');

    // HTML should generate without error
    var html = generateProgressReportHtml_(student, 'Q2', entries, '');
    assertNotNull_(html, 'Report should generate with long notes');
    assertContains_(html, 'student progress', 'Long notes should appear in report');
  } finally {
    if (result && result.id) try { deleteProgressEntry_(result.id); } catch(e) {}
  }
}

function test_edge_manyObjectives() {
  // Goal with 6 objectives renders correctly
  var objectives = [];
  for (var i = 1; i <= 6; i++) {
    objectives.push({ id: 'obj-many-' + i, text: 'Objective number ' + i + ' for this goal.' });
  }
  var student = buildMockStudent_({
    goalsJson: JSON.stringify([
      {
        id: 'goal-many',
        text: 'Goal with many objectives for testing.',
        goalArea: 'Organization',
        objectives: objectives
      }
    ])
  });

  var entries = objectives.map(function(obj, idx) {
    return {
      goalId: 'goal-many',
      objectiveId: obj.id,
      quarter: 'Q2',
      progressRating: idx % 3 === 0 ? 'No Progress' : idx % 3 === 1 ? 'Adequate Progress' : 'Objective Met',
      anecdotalNotes: 'Progress notes for objective ' + (idx + 1) + ' this quarter.'
    };
  });

  var reportData = assembleReportData_(student, 'Q2', entries);
  var orgGroup = reportData.goalGroups[0];
  assertEqual_(orgGroup.goals[0].objectives.length, 6, 'Should have all 6 objectives');

  var html = generateProgressReportHtml_(student, 'Q2', entries, '');
  assertNotNull_(html, 'Report should generate with many objectives');
  assertContains_(html, 'Objective number 6', 'All objectives should be in report');
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 6. QUARTER UTILITIES
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function test_getCurrentQuarter_returnsValidQuarter() {
  var q = getCurrentQuarter();
  assert_(
    q === 'Q1' || q === 'Q2' || q === 'Q3' || q === 'Q4',
    'getCurrentQuarter should return Q1-Q4, got: ' + q
  );
}

function test_getQuarterLabel_formatsCorrectly() {
  // Quarter labels should be human-readable with season and year range
  var label = getQuarterLabel('Q1');
  assertNotNull_(label, 'Quarter label should not be null');
  assertContains_(label, 'Q1', 'Label should contain quarter identifier');

  var label2 = getQuarterLabel('Q2');
  assertContains_(label2, 'Q2', 'Q2 label should contain Q2');
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 7. DUE PROCESS — IEP Meetings & Completion Map
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function runAllDueProcessTests() {
  var tests = [
    'test_iepMeeting_savesAndRetrieves',
    'test_iepMeeting_validatesRequiredFields',
    'test_iepMeeting_validatesMeetingType',
    'test_iepMeeting_deletesById',
    'test_goalResponsible_persistsViaGoalsJson',
    'test_completionMap_allDoneWhenComplete',
    'test_completionMap_notDoneWhenMissing',
    'test_completionMap_ignoresOtherUsersGoals'
  ];

  var testFns = {
    test_iepMeeting_savesAndRetrieves: test_iepMeeting_savesAndRetrieves,
    test_iepMeeting_validatesRequiredFields: test_iepMeeting_validatesRequiredFields,
    test_iepMeeting_validatesMeetingType: test_iepMeeting_validatesMeetingType,
    test_iepMeeting_deletesById: test_iepMeeting_deletesById,
    test_goalResponsible_persistsViaGoalsJson: test_goalResponsible_persistsViaGoalsJson,
    test_completionMap_allDoneWhenComplete: test_completionMap_allDoneWhenComplete,
    test_completionMap_notDoneWhenMissing: test_completionMap_notDoneWhenMissing,
    test_completionMap_ignoresOtherUsersGoals: test_completionMap_ignoresOtherUsersGoals
  };

  var passed = 0, failed = 0, errors = [];
  tests.forEach(function(name) {
    try {
      testFns[name]();
      passed++;
      Logger.log('PASS: ' + name);
    } catch(e) {
      failed++;
      errors.push(name + ': ' + e.message);
      Logger.log('FAIL: ' + name + ' — ' + e.message);
    }
  });

  Logger.log('');
  Logger.log('Due Process Results: ' + passed + ' passed, ' + failed + ' failed');
  if (errors.length > 0) {
    Logger.log('Failures:');
    errors.forEach(function(e) { Logger.log('  ' + e); });
  }
  return { passed: passed, failed: failed, errors: errors };
}

function test_iepMeeting_savesAndRetrieves() {
  var meetingId = null;
  try {
    var result = saveIEPMeeting({
      studentId: 'stu-test-001',
      meetingDate: '2026-03-15',
      meetingType: 'Annual Review',
      notes: 'Test meeting for IEP review'
    });
    assert_(result.success, 'Meeting should save successfully');
    assertNotNull_(result.id, 'Should return an ID');
    meetingId = result.id;

    var meetings = getIEPMeetings('stu-test-001');
    var found = meetings.filter(function(m) { return m.id === meetingId; });
    assertEqual_(found.length, 1, 'Should find the saved meeting');
    assertEqual_(found[0].meetingType, 'Annual Review', 'Meeting type should match');
    assertEqual_(found[0].notes, 'Test meeting for IEP review', 'Notes should match');
  } finally {
    if (meetingId) try { deleteIEPMeeting(meetingId); } catch(e) {}
  }
}

function test_iepMeeting_validatesRequiredFields() {
  var result1 = saveIEPMeeting({});
  assert_(!result1.success, 'Should fail without required fields');

  var result2 = saveIEPMeeting({ studentId: 'stu-test-001' });
  assert_(!result2.success, 'Should fail without meetingDate');

  var result3 = saveIEPMeeting({ studentId: 'stu-test-001', meetingDate: '2026-03-15' });
  assert_(!result3.success, 'Should fail without meetingType');
}

function test_iepMeeting_validatesMeetingType() {
  var result = saveIEPMeeting({
    studentId: 'stu-test-001',
    meetingDate: '2026-03-15',
    meetingType: 'Invalid Type'
  });
  assert_(!result.success, 'Should reject invalid meeting type');
  assertContains_(result.error, 'Invalid', 'Error should mention invalid type');
}

function test_iepMeeting_deletesById() {
  var meetingId = null;
  try {
    var result = saveIEPMeeting({
      studentId: 'stu-test-001',
      meetingDate: '2026-04-01',
      meetingType: 'Amendment',
      notes: 'To be deleted'
    });
    assert_(result.success, 'Meeting should save');
    meetingId = result.id;

    var delResult = deleteIEPMeeting(meetingId);
    assert_(delResult.success, 'Delete should succeed');
    meetingId = null; // Already deleted

    var meetings = getIEPMeetings('stu-test-001');
    var found = meetings.filter(function(m) { return m.id === result.id; });
    assertEqual_(found.length, 0, 'Meeting should no longer exist after delete');
  } finally {
    if (meetingId) try { deleteIEPMeeting(meetingId); } catch(e) {}
  }
}

function test_goalResponsible_persistsViaGoalsJson() {
  // Verify that responsibleEmail survives round-trip through saveStudentGoals
  var goalsWithResponsible = JSON.stringify([{
    id: 'goal-resp-test',
    text: 'Test goal with responsibility',
    goalArea: 'Math Calculation',
    responsibleEmail: 'teacher@rpsmn.org',
    objectives: [{ id: 'obj-resp-1', text: 'Test objective' }]
  }]);

  var result = saveStudentGoals('stu-test-001', goalsWithResponsible);
  assert_(result.success, 'Should save goals with responsibleEmail');

  var students = getStudents();
  var student = null;
  for (var i = 0; i < students.length; i++) {
    if (students[i].id === 'stu-test-001') { student = students[i]; break; }
  }

  if (student) {
    var goals = JSON.parse(student.goalsJson || '[]');
    assert_(goals.length > 0, 'Should have at least one goal');
    assertEqual_(goals[0].responsibleEmail, 'teacher@rpsmn.org', 'responsibleEmail should persist');
  }
}

function test_completionMap_allDoneWhenComplete() {
  // When all objectives have entries, allDone should be true
  var email = 'teacher@rpsmn.org';
  var students = [buildMockStudent_({
    goalsJson: JSON.stringify([{
      id: 'goal-comp-1',
      text: 'Goal 1',
      goalArea: 'Math',
      responsibleEmail: email,
      objectives: [
        { id: 'obj-c1', text: 'Obj 1' },
        { id: 'obj-c2', text: 'Obj 2' }
      ]
    }])
  })];

  // Create entries for both objectives
  var createdIds = [];
  try {
    var e1 = saveProgressEntry({
      studentId: 'stu-test-001', goalId: 'goal-comp-1', objectiveId: 'obj-c1',
      quarter: 'Q2', progressRating: 'Adequate Progress', anecdotalNotes: 'Good work on this objective test.'
    });
    if (e1.id) createdIds.push(e1.id);

    var e2 = saveProgressEntry({
      studentId: 'stu-test-001', goalId: 'goal-comp-1', objectiveId: 'obj-c2',
      quarter: 'Q2', progressRating: 'Objective Met', anecdotalNotes: 'Excellent work on this objective.'
    });
    if (e2.id) createdIds.push(e2.id);

    var map = buildCompletionMap_(email, students, 'Q2');
    assertNotNull_(map['stu-test-001'], 'Should have entry for test student');
    assert_(map['stu-test-001'].allDone, 'Should be allDone when both objectives have entries');
    assertEqual_(map['stu-test-001'].total, 2, 'Total should be 2');
    assertEqual_(map['stu-test-001'].completed, 2, 'Completed should be 2');
  } finally {
    createdIds.forEach(function(id) { try { deleteProgressEntry_(id); } catch(e) {} });
  }
}

function test_completionMap_notDoneWhenMissing() {
  // When one objective is missing an entry, allDone should be false
  var email = 'teacher@rpsmn.org';
  var students = [buildMockStudent_({
    goalsJson: JSON.stringify([{
      id: 'goal-inc-1',
      text: 'Goal incomplete',
      goalArea: 'Reading',
      responsibleEmail: email,
      objectives: [
        { id: 'obj-i1', text: 'Obj 1' },
        { id: 'obj-i2', text: 'Obj 2' }
      ]
    }])
  })];

  var createdIds = [];
  try {
    // Only create entry for one objective
    var e1 = saveProgressEntry({
      studentId: 'stu-test-001', goalId: 'goal-inc-1', objectiveId: 'obj-i1',
      quarter: 'Q2', progressRating: 'No Progress', anecdotalNotes: 'Needs more support on this skill.'
    });
    if (e1.id) createdIds.push(e1.id);

    var map = buildCompletionMap_(email, students, 'Q2');
    assertNotNull_(map['stu-test-001'], 'Should have entry for test student');
    assert_(!map['stu-test-001'].allDone, 'Should NOT be allDone when one objective missing');
    assertEqual_(map['stu-test-001'].total, 2, 'Total should be 2');
    assertEqual_(map['stu-test-001'].completed, 1, 'Completed should be 1');
  } finally {
    createdIds.forEach(function(id) { try { deleteProgressEntry_(id); } catch(e) {} });
  }
}

function test_completionMap_ignoresOtherUsersGoals() {
  // Goals assigned to a different user should not appear in the map
  var myEmail = 'teacher@rpsmn.org';
  var students = [buildMockStudent_({
    goalsJson: JSON.stringify([
      {
        id: 'goal-mine',
        text: 'My goal',
        goalArea: 'Math',
        responsibleEmail: myEmail,
        objectives: [{ id: 'obj-m1', text: 'My obj' }]
      },
      {
        id: 'goal-theirs',
        text: 'Their goal',
        goalArea: 'Reading',
        responsibleEmail: 'other@rpsmn.org',
        objectives: [{ id: 'obj-t1', text: 'Their obj' }]
      }
    ])
  })];

  var createdIds = [];
  try {
    var e1 = saveProgressEntry({
      studentId: 'stu-test-001', goalId: 'goal-mine', objectiveId: 'obj-m1',
      quarter: 'Q2', progressRating: 'Adequate Progress', anecdotalNotes: 'Making good progress on this.'
    });
    if (e1.id) createdIds.push(e1.id);

    // Only my goal has entry — should be complete for me
    var map = buildCompletionMap_(myEmail, students, 'Q2');
    assertNotNull_(map['stu-test-001'], 'Should have entry for test student');
    assertEqual_(map['stu-test-001'].total, 1, 'Total should only count MY goals (1)');
    assertEqual_(map['stu-test-001'].completed, 1, 'Completed should be 1');
    assert_(map['stu-test-001'].allDone, 'Should be allDone — only my goal counts');
  } finally {
    createdIds.forEach(function(id) { try { deleteProgressEntry_(id); } catch(e) {} });
  }
}
