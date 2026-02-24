# SPED Lead Supervisory Dashboard — Design Document

**Date:** February 24, 2026
**Phase:** Phase 2 (Cross-Caseload Oversight)
**Status:** Approved for Implementation

---

## Overview

This design adds a **SPED Lead role** to the EFcheckin caseload management system. SPED Leads are school heads of special education who need cross-caseload visibility to monitor eval pipeline status, due process compliance, and staff workload across 6-15 case managers. Unlike case managers who own a caseload, SPED Leads are **supervisory observers** with read-only access to all connected caseloads plus full write access to due process features (IEP meetings, progress reporting).

### Key Requirements

- **Bird's eye view:** Aggregate metrics across all case managers (no granular check-in/academic data)
- **Cross-caseload access:** See data from multiple case managers' spreadsheets
- **Drill-down capability:** View individual students' full profiles (read-only)
- **Due process authority:** Edit IEP meeting dates, progress entries, compliance tracking
- **Managed by superuser:** Superuser designates SPED Leads and connects them to case managers

### Architectural Context

The current system uses **per-user data isolation** (FERPA compliant):
- Each case manager has their own Google Sheet stored in UserProperties
- `getSS_()` only opens one spreadsheet at a time (the authenticated user's)
- Team sharing works via Drive editor permissions + CoTeachers sheet invites

The SPED Lead role breaks this pattern by requiring **cross-spreadsheet reads** from multiple case managers' sheets.

---

## Section 1: SPED Lead Role & Data Model

### ScriptProperties Storage

SPED Leads are tracked globally in ScriptProperties (not in any single caseload's CoTeachers sheet):

```javascript
// ScriptProperties keys:
sped_leads = ["lead@rpsmn.org"]  // Array of SPED Lead emails

sped_lead_caseloads_{email} = [
  {"email": "cm1@rpsmn.org", "name": "Jane Smith", "spreadsheetId": "abc123"},
  {"email": "cm2@rpsmn.org", "name": "John Doe", "spreadsheetId": "def456"}
]

sped_lead_last_sync_{email} = "2026-02-24T02:15:00.000Z"  // ISO timestamp

sped_lead_spreadsheet_{email} = "xyz789"  // SPED Lead's dashboard spreadsheet ID
```

### SPED Lead Spreadsheet Schema

Each SPED Lead gets a dedicated spreadsheet with three sheets:

**AggregateMetrics Sheet:**
| caseManagerEmail | caseManagerName | studentCount | activeEvals | overdueEvals | upcomingIEPs | progressCompletionRate | lastSyncStatus | lastSyncDate |
|---|---|---|---|---|---|---|---|---|
| cm1@rpsmn.org | Jane Smith | 12 | 2 | 1 | 3 | 92% | OK | 2026-02-24T02:15:00Z |

**AllStudents Sheet:**
| studentId | firstName | lastName | caseManagerEmail | caseManagerName | grade | evalType | evalStatus | nextIEPDate | gpa |
|---|---|---|---|---|---|---|---|---|---|
| abc123 | Alice | Smith | cm1@rpsmn.org | Jane Smith | 10 | Annual IEP | Active | 2026-03-15 | 3.2 |

**ComplianceTimeline Sheet:**
| date | studentId | studentName | caseManagerEmail | type | meetingType | evalType |
|---|---|---|---|---|---|---|
| 2026-02-26 | abc123 | Alice Smith | cm1@rpsmn.org | IEP | Annual IEP | - |
| 2026-02-27 | def456 | Bob Johnson | cm2@rpsmn.org | Eval | - | Initial Eval |

### Sync Strategy

- **Daily time-driven trigger:** Runs at 2am
- **Manual refresh:** User-initiated via "Refresh Data" button
- **Automatic recovery:** If data is >24 hours old, auto-sync before rendering dashboard
- **Data source:** Reads live from case manager spreadsheets, writes aggregates to SPED Lead sheet
- **Staleness display:** "Last synced at 2:15 am on 02/24" (no color coding)

### Key Design Decision: No Personal Caseload

SPED Leads do NOT get their own student caseload or check-in tracking. Their spreadsheet only contains synced aggregate data. They have no UserProperties entry for a personal caseload spreadsheet.

---

## Section 2: Backend Sync Architecture

### Sync Orchestrator Function

**`syncSpedLeadDashboard(spedLeadEmail)`**

Main workflow:
1. Reads `sped_lead_caseloads_{email}` from ScriptProperties to get list of CM spreadsheet IDs
2. Opens SPED Lead's spreadsheet via `SpreadsheetApp.openById(sped_lead_spreadsheet_{email})`
3. For each case manager:
   - Opens CM's spreadsheet via `SpreadsheetApp.openById(cmSpreadsheetId)`
   - Calls helper functions: `getEvalMetrics_(ss)`, `getDueProcessMetrics_(ss)`, `getStudentSummaries_(ss)`
   - Writes one row to **AggregateMetrics** sheet
   - Appends students to **AllStudents** sheet (tagged with `caseManagerEmail`)
   - Appends IEP/eval dates to **ComplianceTimeline** sheet
   - Uses try-catch per CM to handle partial failures
   - Adds 500ms delay between CMs to avoid quota limits
4. Updates `sped_lead_last_sync_{email}` timestamp
5. Returns `{success: true, syncedCount: 12, failedCount: 3, errors: [...]}`

### Helper Aggregation Functions

**`getEvalMetrics_(ss)`**
- Reads Evaluations sheet from CM's spreadsheet
- Returns: `{activeCount, overdueCount, dueThisWeekCount}`

**`getDueProcessMetrics_(ss, quarter)`**
- Reads ProgressReporting + IEPMeetings sheets
- Returns: `{upcomingIEPs, progressCompletionRate}`

**`getStudentSummaries_(ss)`**
- Reads Students + CheckIns + Evaluations sheets
- Returns flattened array: `[{studentId, firstName, lastName, grade, evalType, nextIEPDate, gpa}, ...]`

### Trigger Management

**`installSpedLeadSyncTrigger(spedLeadEmail)`**
- Creates daily time-based trigger for 2am
- Stores trigger ID in ScriptProperties: `sped_lead_trigger_{email}`

**`onSpedLeadDailySyncTrigger(e)`**
- Trigger handler that reads `sped_leads` array
- Calls `syncSpedLeadDashboard()` for each SPED Lead

**`removeSpedLeadSyncTrigger(spedLeadEmail)`**
- Deletes trigger when SPED Lead is removed

### Cross-Spreadsheet Write Access

When SPED Lead edits due process features (IEP meetings, progress entries), the backend needs to write to the **target case manager's spreadsheet**:

```javascript
function saveIEPMeeting(data) {
  var status = getUserStatus();
  if (status.role === 'sped-lead') {
    // Frontend passes caseManagerSpreadsheetId
    var targetSpreadsheetId = data.caseManagerSpreadsheetId;
    var ss = SpreadsheetApp.openById(targetSpreadsheetId);
    // ... write to IEPMeetings sheet
    // ... invalidate that CM's cache
  } else {
    var ss = getSS_(); // CM writes to their own spreadsheet
  }
  // ... existing logic
}
```

### Quota Considerations

- **Spreadsheet opens:** ~20/minute limit → 500ms delay between CMs → 15 CMs = 7.5 seconds
- **Read operations:** Each CM requires ~5 sheet reads → 75 reads for 15 CMs (well under quota)
- **Trigger quota:** Daily trigger = 1 trigger/day (minimal usage)

---

## Section 3: Frontend SPED Lead Dashboard UI

### New Views

**Index.html additions:**
- `<div id="spedlead-view" class="view">` — Overview dashboard
- `<div id="spedlead-students-view" class="view">` — All students list
- `<div id="spedlead-cm-profile-view" class="view">` — Case Manager profile
- `<div id="spedlead-timeline-modal" class="modal">` — Timeline day detail modal

### Nav Drawer (SPED Lead)

- **Dashboard** → Overview with aggregate metrics, timeline, CM table
- **Students** → Filterable list of all students across caseloads
- **Team** → Admin/superuser only (manage SPED Leads)
- **Settings** → User settings

No access to: Add Student, Evals, Due Process tabs (case manager features)

### Dashboard Overview Layout

**Header:**
```
┌─────────────────────────────────────────────────────────┐
│  SPED Lead Overview        Last synced at 2:15 am on    │
│  15 Case Managers          02/24          [Refresh Data] │
└─────────────────────────────────────────────────────────┘
```

**Timeline Strip (7-day window):**
```
┌─────────────────────────────────────────────────────────┐
│ Upcoming Evals & IEPs                                   │
│ ┌───────┐ ┌───────┐ ┌───────┐ ┌───────┐ ┌───────┐     │
│ │ Mon   │ │ Tue   │ │ Wed   │ │ Thu   │ │ Fri   │ ... │
│ │ 2/24  │ │ 2/25  │ │ 2/26  │ │ 2/27  │ │ 2/28  │     │
│ │       │ │       │ │       │ │       │ │       │     │
│ │ 3 IEP │ │ 1 IEP │ │ 5 IEP │ │ 2 IEP │ │ 4 IEP │     │
│ │ 2 Eval│ │ 1 Eval│ │       │ │ 3 Eval│ │ 1 Eval│     │
│ └───────┘ └───────┘ └───────┘ └───────┘ └───────┘     │
└─────────────────────────────────────────────────────────┘
```

Click a day → Container transform to modal (see Timeline Modal section)

**Aggregate Metric Cards (3 across):**
```
┌──────────────┐ ┌──────────────┐ ┌──────────────┐
│ Eval Pipeline│ │Due Process   │ │Staff Workload│
│              │ │              │ │              │
│ 23 Active    │ │ 8 IEPs       │ │ 156 Students │
│ 4 Overdue    │ │ This Week    │ │ Avg: 10.4    │
│ 12 Due Soon  │ │              │ │ Range: 4-18  │
│              │ │ 85% Progress │ │              │
│              │ │ Complete     │ │              │
└──────────────┘ └──────────────┘ └──────────────┘
```

**Case Manager Breakdown Table (sortable):**
```
┌─────────────┬──────────┬───────┬─────────┬──────────┬─────────┐
│ Case Manager│ Students │ Evals │ Overdue │ IEPs Due │ Progress│
├─────────────┼──────────┼───────┼─────────┼──────────┼─────────┤
│ Jane Smith  │    12    │   2   │    1    │    3     │   92%   │
│ John Doe    │    18    │   4   │    0    │    1     │   100%  │
│ ...         │   ...    │  ...  │   ...   │   ...    │   ...   │
└─────────────┴──────────┴───────┴─────────┴──────────┴─────────┘
```

Click a row → Navigate to Case Manager Profile view

Data source: **AggregateMetrics** and **ComplianceTimeline** sheets

### Students View Layout

**Filter Bar:**
```
┌────────────────────────────────────────────────────────────┐
│ [Case Manager ▼] [Grade ▼] [Search by name...         ]   │
└────────────────────────────────────────────────────────────┘
```

Filters:
- **Case Manager:** Dropdown with "All" + list of connected CMs
- **Grade:** Dropdown with "All" + 9/10/11/12
- **Name search:** Real-time filter by firstName or lastName
- **Multi-field:** All three filters apply simultaneously (AND logic)

**Student List Table (sortable):**
```
┌──────────────┬───────────────┬───────┬──────────┬──────────────┐
│ Name         │ Case Manager  │ Grade │ Eval     │ Next IEP     │
├──────────────┼───────────────┼───────┼──────────┼──────────────┤
│ Alice Smith  │ Jane Smith    │  10   │ Annual   │ Mar 15, 2026 │
│ Bob Johnson  │ John Doe      │  11   │ Re-Eval  │ Apr 3, 2026  │
│ ...          │ ...           │  ...  │ ...      │ ...          │
└──────────────┴───────────────┴───────┴──────────┴──────────────┘
```

Click a row → Open student profile (read-only, from source CM spreadsheet)

Data source: **AllStudents** sheet (filtered client-side)

### Case Manager Profile View Layout

**Header with breadcrumb:**
```
┌────────────────────────────────────────────────────────┐
│ [← Dashboard] > Jane Smith                             │
│ jane.smith@rpsmn.org                    [✉ Email]      │
└────────────────────────────────────────────────────────┘
```

**Timeline Strip (7-day, filtered to this CM):**
```
┌─────────────────────────────────────────────────────────┐
│ Jane Smith's Upcoming Evals & IEPs                      │
│ ┌───────┐ ┌───────┐ ┌───────┐ ...                      │
│ │ Mon   │ │ Tue   │ │ Wed   │                          │
│ │ 1 IEP │ │       │ │ 2 IEP │                          │
│ │ 1 Eval│ │ 1 Eval│ │       │                          │
│ └───────┘ └───────┘ └───────┘                          │
└─────────────────────────────────────────────────────────┘
```

**Metric Cards (4 across):**
```
┌──────────────┐ ┌──────────────┐ ┌──────────────┐ ┌──────────────┐
│ Total        │ │ Active Evals │ │ Overdue Evals│ │ IEPs Due     │
│ Students     │ │              │ │              │ │ This Week    │
│     12       │ │      2       │ │      1       │ │      3       │
└──────────────┘ └──────────────┘ └──────────────┘ └──────────────┘
```

**Student List (no CM column):**
```
┌──────────────┬───────┬──────────┬──────────────┐
│ Name         │ Grade │ Eval     │ Next IEP     │
├──────────────┼───────┼──────────┼──────────────┤
│ Alice Smith  │  10   │ Annual   │ Mar 15, 2026 │
│ Charlie Lee  │   9   │ Initial  │ May 1, 2026  │
│ ...          │  ...  │ ...      │ ...          │
└──────────────┴───────┴──────────┴──────────────┘
```

Data source: **AggregateMetrics** (for metrics) + **AllStudents** filtered by `caseManagerEmail`

### Timeline Day Detail Modal

**Animation:** Container transform (Material Design 3)
- Day card scales/morphs into modal header
- Content fades in below
- Backdrop fades in simultaneously
- Close → reverse animation back to day card

**Modal Structure:**
```
┌─────────────────────────────────────────────────────────┐
│ Wednesday, Feb 26, 2026                        [✕ Close]│
│                                                          │
│ ──── IEP Meetings (5) ────                              │
│ ┌─────────────────────────────────────────────────────┐ │
│ │ Alice Smith (Jane Smith)           Annual IEP       │ │
│ │ Bob Johnson (John Doe)             IEP Amendment    │ │
│ │ ...                                                  │ │
│ └─────────────────────────────────────────────────────┘ │
│                                                          │
│ ──── Evaluations (2) ────                               │
│ ┌─────────────────────────────────────────────────────┐ │
│ │ Charlie Lee (Jane Smith)           Initial Eval     │ │
│ │ Dana Park (Sarah Connor)           3 Year Re-Eval   │ │
│ └─────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────┘
```

**IEP Meeting Types:**
- Annual IEP
- IEP Amendment

**Evaluation Types:**
- Initial Eval
- 3 Year Re-Eval

**Interactions:**
- Click student name → Open student profile (read-only)
- Click IEP meeting row → Edit meeting date (SPED Lead has write access)
- Click eval row → View eval details (read-only)
- Close or click backdrop → Container transform back to day card

Data source: **ComplianceTimeline** sheet filtered by date (and optionally by CM if opened from CM profile)

---

## Section 4: Superuser Admin Interface

### New Admin View Section

Added to existing `admin-view` in Index.html:

**SPED Lead Management Card:**
```
┌─────────────────────────────────────────────────────────┐
│ SPED Lead Management                                    │
│                                                          │
│ Current SPED Leads:                                     │
│ ┌─────────────────────────────────────────────────────┐ │
│ │ lead@rpsmn.org                         [Remove]     │ │
│ └─────────────────────────────────────────────────────┘ │
│                                                          │
│ Add SPED Lead:                                          │
│ [email@rpsmn.org        ] [Add SPED Lead]              │
│                                                          │
│ Configure Access for: [Select SPED Lead ▼]             │
│                                                          │
│ Connected Case Managers (12):                           │
│ ┌─────────────────────────────────────────────────────┐ │
│ │ ☑ Jane Smith (jane.smith@rpsmn.org)                 │ │
│ │ ☑ John Doe (john.doe@rpsmn.org)                     │ │
│ │ ☐ Sarah Connor (sarah.connor@rpsmn.org)             │ │
│ │ ...                                                  │ │
│ └─────────────────────────────────────────────────────┘ │
│                                                          │
│ [Save Connections]  [Provision Spreadsheet]            │
└─────────────────────────────────────────────────────────┘
```

### Workflow

**1. Add SPED Lead:**
- Superuser enters email, clicks "Add SPED Lead"
- Backend: `addSpedLead(email)` → adds to `sped_leads` array in ScriptProperties
- Toast: "SPED Lead added. Configure access and provision spreadsheet below."

**2. Connect Case Managers:**
- Superuser selects SPED Lead from dropdown
- Checks/unchecks case managers (from existing `case_managers` registry)
- Clicks "Save Connections"
- Backend: `updateSpedLeadConnections(spedLeadEmail, caseManagerEmails[])`
  - Updates `sped_lead_caseloads_{email}` in ScriptProperties with CM emails + spreadsheet IDs
  - For each checked CM: `cmSpreadsheet.addEditor(spedLeadEmail)` (grants Drive edit access)
- Toast: "Connected SPED Lead to 12 case managers"

**3. Provision Spreadsheet:**
- Clicks "Provision Spreadsheet"
- Backend: `provisionSpedLeadSpreadsheet(spedLeadEmail)`
  - Creates "SPED Lead Dashboard - {name}" in SPED Lead's Drive
  - Creates 3 sheets: AggregateMetrics, AllStudents, ComplianceTimeline
  - Stores spreadsheet ID in `sped_lead_spreadsheet_{email}` ScriptProperties
  - Runs initial `syncSpedLeadDashboard()` to populate data
  - Calls `installSpedLeadSyncTrigger(spedLeadEmail)` to set up daily 2am sync
- Toast: "Spreadsheet provisioned and synced. SPED Lead is ready."

**4. Remove SPED Lead:**
- Clicks "Remove" next to SPED Lead email
- Confirmation dialog: "This will revoke access and delete the SPED Lead spreadsheet. Continue?"
- Backend: `removeSpedLead(email)`
  - Removes from `sped_leads` array
  - Deletes trigger via `removeSpedLeadSyncTrigger(email)`
  - Optionally deletes spreadsheet (or leaves orphaned for manual cleanup)
  - Removes editor access from all CM spreadsheets

### New Backend Functions

**`addSpedLead(email)`**
- Validates email format
- Checks not already in `sped_leads` array
- Appends to `sped_leads` in ScriptProperties
- Returns `{success: true}`

**`removeSpedLead(email)`**
- Removes from `sped_leads` array
- Deletes `sped_lead_caseloads_{email}`, `sped_lead_spreadsheet_{email}`, `sped_lead_last_sync_{email}`
- Calls `removeSpedLeadSyncTrigger(email)`
- Returns `{success: true}`

**`updateSpedLeadConnections(spedLeadEmail, caseManagerEmails[])`**
- Reads each CM's spreadsheet ID from the global `case_managers` registry (which stores `{email, name, spreadsheetId}`)
- Builds array: `[{email, name, spreadsheetId}, ...]`
- Stores in `sped_lead_caseloads_{email}` ScriptProperties
- For each CM: `SpreadsheetApp.openById(cmSpreadsheetId).addEditor(spedLeadEmail)`
- Returns `{success: true, connectedCount: N}`

**`provisionSpedLeadSpreadsheet(spedLeadEmail)`**
- Creates new spreadsheet: `SpreadsheetApp.create("SPED Lead Dashboard - " + name)`
- Adds 3 sheets with headers
- Stores ID in `sped_lead_spreadsheet_{email}` ScriptProperties
- Runs `syncSpedLeadDashboard(spedLeadEmail)` to populate
- Calls `installSpedLeadSyncTrigger(spedLeadEmail)`
- Returns `{success: true, spreadsheetId, syncResult}`

**`getSpedLeads_()`**
- Reads `sped_leads` from ScriptProperties
- Returns array of emails (or empty array if not found)

**`getSpedLeadCaseloads_(email)`**
- Reads `sped_lead_caseloads_{email}` from ScriptProperties
- Returns array of `{email, name, spreadsheetId}` objects

---

## Section 5: Permission & Access Control

### Permission Resolution

**getUserStatus() for SPED Lead:**

```javascript
function getUserStatus() {
  var userEmail = getCurrentUserEmail_();

  // Check if SPED Lead
  var spedLeads = getSpedLeads_();
  if (spedLeads.indexOf(userEmail) !== -1) {
    return {
      role: 'sped-lead',
      permissions: null,  // null = unrestricted access for due process
      connectedCaseloads: getSpedLeadCaseloads_(userEmail)
    };
  }

  // ... existing CM/co-teacher logic
}
```

### Access Rules

**Read Access:**
- **Unrestricted** read access to all connected case managers' data:
  - Students, Evaluations, IEP Meetings, Progress Reporting
  - Does NOT see: Check-in history, academic data (grades/missing work), teacher feedback

**Write Access:**
- **Due Process only:** Full write access to:
  - `saveIEPMeeting()` — Create/edit IEP meetings
  - `updateIEPMeetingDate()` — Change meeting dates
  - `deleteIEPMeeting()` — Delete meetings
  - `saveProgressEntry()` — Add/edit progress entries
  - `toggleDPReportComplete()` — Mark reports complete
- **Blocked:** All other write operations:
  - `saveStudent()`, `saveCheckIn()`, `saveEvaluation()`, `updateCheckInAcademicData()`, etc.

### Frontend Guards

**Buttons hidden for SPED Lead:**
- "Add Student", "New Check-In", "Edit Student", "Edit Goals", "Edit Contacts"
- Grade select dropdowns (render as plain text)
- Academic data "Mark Done" / "Add Missing Assignment" buttons
- Eval checklist edit buttons

**Buttons shown for SPED Lead:**
- "Edit" on IEP meeting dates (in timeline modal)
- "Add Progress Entry" in Due Process view
- "Mark Report Complete" checkboxes

**Read-only student profiles:**
- All form inputs render as plain text displays
- No edit/save buttons
- Timeline and due process sections show data but no modification controls

### Backend Endpoint Guards

**Write endpoints - blocked:**
```javascript
function saveStudent(data) {
  var status = getUserStatus();
  if (status.role === 'sped-lead') {
    throw new Error('SPED Leads have read-only access to student data');
  }
  // ... existing permission checks
}
```

**Write endpoints - allowed (due process):**
```javascript
function saveIEPMeeting(data) {
  var status = getUserStatus();

  if (status.role === 'sped-lead') {
    // SPED Lead writes to target CM's spreadsheet
    var targetSpreadsheetId = data.caseManagerSpreadsheetId;
    if (!targetSpreadsheetId) {
      throw new Error('caseManagerSpreadsheetId required for SPED Lead writes');
    }

    // Verify SPED Lead has access to this caseload
    var caseloads = getSpedLeadCaseloads_(status.email);
    var hasAccess = caseloads.some(function(cm) {
      return cm.spreadsheetId === targetSpreadsheetId;
    });
    if (!hasAccess) {
      throw new Error('Access denied to this caseload');
    }

    var ss = SpreadsheetApp.openById(targetSpreadsheetId);
    // ... write to IEPMeetings sheet
    // ... invalidate that CM's cache
    return {success: true};
  }

  // Case manager writes to their own spreadsheet
  var ss = getSS_();
  // ... existing logic
}
```

**Read endpoints - enhanced for cross-caseload:**
```javascript
function getStudent(studentId, caseManagerSpreadsheetId) {
  var status = getUserStatus();

  if (status.role === 'sped-lead') {
    // SPED Lead reads from target CM's spreadsheet
    var ss = SpreadsheetApp.openById(caseManagerSpreadsheetId);
    // ... read Students sheet
  } else {
    var ss = getSS_();
    // ... existing logic
  }
}
```

### Key Security Principle

SPED Leads have **supervisory observer** status with compliance authority:
- **See everything** across all connected caseloads (except granular check-in/academic data)
- **Touch nothing** except due process features (IEP meetings, progress entries)
- **Cannot create/edit/delete** students, check-ins, evals, goals, contacts, academic data

---

## Section 6: Error Handling & Edge Cases

### Sync Failures

**Partial sync failure (some CMs succeed, others fail):**

```javascript
function syncSpedLeadDashboard(spedLeadEmail) {
  var caseloads = getSpedLeadCaseloads_(spedLeadEmail);
  var spedLeadSS = SpreadsheetApp.openById(getSpedLeadSpreadsheetId_(spedLeadEmail));

  var syncedCount = 0;
  var failedCount = 0;
  var errors = [];

  caseloads.forEach(function(cm) {
    try {
      var cmSS = SpreadsheetApp.openById(cm.spreadsheetId);
      var metrics = getEvalMetrics_(cmSS);
      // ... write to AggregateMetrics
      syncedCount++;
      Utilities.sleep(500); // Rate limiting
    } catch(e) {
      failedCount++;
      errors.push({cm: cm.email, error: e.message});
      // Write "Failed" status to AggregateMetrics
    }
  });

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
```

Frontend displays: "Synced 12 of 15 caseloads. 3 failed. [View Details]"

**AggregateMetrics lastSyncStatus column:**
- "OK" — Successfully synced
- "Failed" — Access error or spreadsheet not found

### Access Revoked Scenarios

**Case manager removes SPED Lead as editor:**
- Next sync fails with "Permission denied" error
- That CM's row shows status: "Access Lost"
- Superuser admin view shows warning: "SPED Lead 'lead@rpsmn.org' has lost access to Jane Smith's caseload"
- Superuser can re-share via `cmSpreadsheet.addEditor(spedLeadEmail)` or disconnect that CM

**Case manager deletes their spreadsheet:**
- Sync fails with "Spreadsheet not found"
- Backend marks CM as "Archived" in `sped_lead_caseloads_{email}` with flag `{archived: true}`
- Dashboard shows grayed-out row: "Jane Smith (Spreadsheet Deleted)"
- Superuser can remove the archived CM from connections

### Automatic Stale Data Recovery

**If data is >24 hours old:**

```javascript
function getUserStatus() {
  var userEmail = getCurrentUserEmail_();

  if (isSpedLead_(userEmail)) {
    var lastSync = PropertiesService.getScriptProperties()
      .getProperty('sped_lead_last_sync_' + userEmail);

    if (lastSync) {
      var lastSyncDate = new Date(lastSync);
      var hoursSinceSync = (Date.now() - lastSyncDate.getTime()) / 1000 / 60 / 60;

      if (hoursSinceSync > 24) {
        // Emergency sync before returning status
        syncSpedLeadDashboard(userEmail);
      }
    }

    return {
      role: 'sped-lead',
      permissions: null,
      connectedCaseloads: getSpedLeadCaseloads_(userEmail)
    };
  }

  // ... existing logic
}
```

Frontend shows loading state: "Syncing data..." while emergency sync runs.

### Write Conflict Handling

**SPED Lead and Case Manager both edit same IEP meeting:**

SPED Lead edits meeting → Backend writes to CM's spreadsheet → Invalidates CM's cache:

```javascript
function saveIEPMeeting(data) {
  // ... validation, permission checks

  var ss = (status.role === 'sped-lead')
    ? SpreadsheetApp.openById(data.caseManagerSpreadsheetId)
    : getSS_();

  // ... write to IEPMeetings sheet

  // Invalidate cache so CM sees fresh data on next load
  invalidateMeetingCaches_();

  return {success: true};
}
```

CM's next dashboard load fetches fresh data from their spreadsheet.

### Quota Limits

**Too many spreadsheet opens (>20/minute):**
- Sync adds 500ms delay between each CM iteration
- For 15 CMs: 15 × 500ms = 7.5 seconds total sync time
- If SPED Lead manages 20+ CMs, increase delay to 1 second (20 seconds total)
- Future optimization: Batch sync into multiple trigger executions if needed

**Trigger quota:**
- Daily trigger = 1/day per SPED Lead
- Current quota: 20 triggers per user → supports 20 SPED Leads per deployment

### Last Synced Display

**Dashboard header:**
- Plain text: "Last synced at 2:15 am on 02/24"
- No color coding or staleness warnings
- Manual refresh button next to timestamp

---

## Section 7: Testing Strategy

### Unit Tests (Tests.gs)

**Permission resolution tests:**

```javascript
function test_spedlead_getStatusReturnsRole() {
  var testEmail = 'spedlead@test.org';

  // Add to sped_leads
  PropertiesService.getScriptProperties().setProperty('sped_leads', JSON.stringify([testEmail]));

  // Mock Session.getActiveUser()
  // ... (requires test harness)

  var status = getUserStatus();
  assertEqual_(status.role, 'sped-lead');
  assertEqual_(status.permissions, null);

  // Cleanup
  PropertiesService.getScriptProperties().deleteProperty('sped_leads');
}

function test_spedlead_unrestricted_dueProcessAccess() {
  // Verify SPED Lead can call saveIEPMeeting with target spreadsheet ID
  // ... (requires test CM spreadsheet setup)
}

function test_spedlead_blocked_studentEdit() {
  // Verify SPED Lead cannot call saveStudent()
  try {
    saveStudent({firstName: 'Test'});
    assert_(false, 'Should have thrown error');
  } catch(e) {
    assertContains_(e.message, 'read-only access');
  }
}
```

**Sync tests:**

```javascript
function test_spedlead_syncAggregatesMultipleCaseloads() {
  // Setup: Create 3 test CM spreadsheets with mock data
  // Add SPED Lead, provision spreadsheet, connect CMs
  // Run syncSpedLeadDashboard()
  // Verify AggregateMetrics sheet has 3 rows
  // Verify AllStudents sheet has combined student list
  // Cleanup
}

function test_spedlead_syncHandlesFailedCaseload() {
  // Setup: 2 valid CMs, 1 invalid spreadsheet ID
  // Run sync
  // Verify result: {syncedCount: 2, failedCount: 1, errors: [...]}
  // Verify AggregateMetrics shows "Failed" status for invalid CM
}

function test_spedlead_autoSyncWhenStale() {
  // Set sped_lead_last_sync_{email} to 25 hours ago
  // Call getUserStatus()
  // Verify syncSpedLeadDashboard() was called (via timestamp update)
}
```

**Cross-spreadsheet access tests:**

```javascript
function test_spedlead_openByIdMultipleSheets() {
  // Verify SPED Lead can open 3 CM spreadsheets in one execution
  // Assert no quota errors
}

function test_spedlead_writeToTargetCaseloadSheet() {
  // SPED Lead calls saveIEPMeeting with caseManagerSpreadsheetId
  // Verify write goes to correct CM's IEPMeetings sheet
  // Verify CM's own spreadsheet unchanged
}
```

### Integration Tests (Manual)

**Superuser workflow:**
1. Navigate to Admin view
2. Add SPED Lead: `spedlead@test.org`
3. Select SPED Lead from dropdown
4. Check 2-3 case managers
5. Click "Save Connections" → verify toast "Connected to 3 CMs"
6. Click "Provision Spreadsheet" → verify toast "Spreadsheet provisioned and synced"
7. Open SPED Lead's spreadsheet in Drive → verify 3 sheets with data
8. Check ScriptProperties: `sped_leads`, `sped_lead_caseloads_{email}`, `sped_lead_spreadsheet_{email}`

**SPED Lead dashboard workflow:**
1. Login as SPED Lead
2. View dashboard → verify aggregate metrics, CM table, timeline strip
3. Click day card on timeline → verify modal opens with IEPs/Evals, container transform animation
4. Click student in modal → verify profile opens (read-only, from source CM spreadsheet)
5. Click IEP meeting → verify edit form shows
6. Change meeting date, save → verify write to CM's spreadsheet
7. Click CM row in table → verify CM profile view loads with metrics + filtered students
8. Click student from CM profile → verify profile opens
9. Navigate to Students view → verify all students listed
10. Filter by CM, grade, name → verify multi-field filtering works
11. Click "Refresh Data" → verify sync runs, toast shows "Synced X of Y caseloads"

**Sync workflow:**
1. Manually trigger sync via "Refresh Data"
2. Watch toast progress (or loading indicator)
3. Verify completion toast: "Synced 15 of 15 caseloads"
4. Check dashboard updates with new data
5. Verify `sped_lead_last_sync_{email}` timestamp in ScriptProperties
6. Verify "Last synced at XX:XX am/pm on MM/DD" updates in header

**Error scenario: Access revoked**
1. As superuser, revoke SPED Lead's editor access to one CM's spreadsheet
2. Trigger manual sync
3. Verify partial failure toast: "Synced 14 of 15 caseloads. 1 failed."
4. Verify AggregateMetrics shows "Access Lost" for that CM
5. Verify admin view shows warning

**Error scenario: Deleted spreadsheet**
1. Delete a CM's spreadsheet
2. Trigger manual sync
3. Verify partial failure
4. Verify dashboard shows "Jane Smith (Spreadsheet Deleted)" in grayed-out row

**Error scenario: Stale data (>24 hours)**
1. Manually set `sped_lead_last_sync_{email}` to 25 hours ago in ScriptProperties
2. Reload SPED Lead dashboard
3. Verify auto-sync runs (loading state shows)
4. Verify dashboard loads with fresh data

---

## Implementation Notes

### Migration Path

**Existing installations:**
- No database migration needed (new ScriptProperties keys, no schema changes to existing sheets)
- Existing case managers unaffected
- Superuser manually onboards SPED Leads via admin view

### Performance Considerations

**Initial sync time:**
- 15 CMs × (5 sheet reads + 500ms delay) ≈ 7-10 seconds
- Acceptable for daily batch + manual refresh use case

**Dashboard load time:**
- Reads from 1 local spreadsheet (SPED Lead's) = ~1-2 seconds
- Drill-down to student profile reads from source CM spreadsheet = ~2-3 seconds (one-time per student)

**Caching strategy:**
- SPED Lead's synced data is the cache (daily refresh = 24hr TTL)
- No additional UserProperties cache needed
- CM cache invalidation on SPED Lead writes ensures consistency

### Future Enhancements (Out of Scope)

- **Real-time sync:** WebSocket or push notifications for cross-caseload updates
- **SPED Lead-specific reports:** Printable compliance reports, trend analysis
- **Multi-school support:** District-level SPED Leads overseeing multiple schools
- **Custom dashboard widgets:** Configurable metric cards, filtering presets

---

## Approval Status

**Design approved:** February 24, 2026
**Ready for implementation planning:** Yes

Next step: Invoke `writing-plans` skill to create detailed implementation plan with step-by-step tasks.
