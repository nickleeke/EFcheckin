# Caseload Dashboard — CLAUDE.md

## Project Overview

A Google Apps Script web application for Richfield Public Schools educators to manage student caseloads, track executive function (EF) skills, monitor academic progress, and coordinate with co-teachers.

## Tech Stack

- **Backend:** Google Apps Script (`code.gs`)
- **Frontend:** Vanilla JavaScript (`JavaScript.html`, ~3,600 lines)
- **Styling:** Vanilla CSS implementing Material Design 3 (`Stylesheet.html`, ~2,500 lines)
- **Data:** Google Sheets (per-user, auto-provisioned in Drive)
- **Auth:** Google Session API (`Session.getActiveUser()`)
- **Hosting:** Google Apps Script web app deployment

## Architecture

### File Structure

```
code.gs            — Backend: CRUD, auth, caching, team management
Index.html         — HTML shell: nav, views, side panel, toast
JavaScript.html    — Frontend: state, rendering, API calls, forms
Stylesheet.html    — CSS: MD3 design tokens, all component styles
README.md          — Multi-user implementation guide
LICENSE            — MIT License
```

### Data Flow

```
Frontend (JS) → google.script.run → Backend (Apps Script) → Google Sheets
                  (async)             (sync)                  (storage)
```

### Per-User Data Isolation (FERPA)

Each user gets their own Google Sheet stored in UserProperties. The web app must be deployed as "Execute as: User accessing the web app" so `Session.getActiveUser()` and `UserProperties` are scoped per-user.

### Data Schemas

**Students sheet:** id, firstName, lastName, grade, period, focusGoal, accommodations, notes, classesJson, createdAt, updatedAt, iepGoal, goalsJson, caseManagerEmail

**CheckIns sheet:** id, studentId, weekOf, planningRating, followThroughRating, regulationRating, focusGoalRating, effortRating, whatWentWell, barrier, microGoal, microGoalCategory, teacherNotes, academicDataJson, createdAt

**CoTeachers sheet:** email, role, addedAt

**Evaluations sheet:** id, studentId, type, itemsJson, createdAt, updatedAt, filesJson

### Eval Types

| Internal Value | Display Label | Template Used |
|---|---|---|
| `annual-iep` | Annual IEP | Re-eval template |
| `3-year-reeval` | 3 Year Re-Eval | Re-eval template |
| `initial-eval` | Initial Eval | Eval template |
| `eval` (legacy) | Initial Eval | Eval template |
| `reeval` (legacy) | 3 Year Re-Eval | Re-eval template |

Type definitions live in `EVAL_TYPES` array + `EVAL_TYPE_ALIASES` map (frontend) and `VALID_EVAL_TYPES` + `EVAL_INITIAL_TYPES_` (backend). Always use `getEvalTypeLabel(type)` for display — never inline ternaries or hardcoded label strings.

### Roles

- `caseload-manager` — Full access, can add/remove team members
- `co-teacher` — Shared access to caseload data
- `service-provider` — Shared access to caseload data
- `para` — Shared access to caseload data
- Superuser (`nicholas.leeke@rpsmn.org`) — Global admin for case manager assignment

### Caching

Read-through cache in UserProperties with 2-minute TTL. Invalidated on writes. Cache keys: `cache_students`, `cache_dashboard`.

### Key Patterns

- **Autosave:** Debounced (2s) with in-flight guard and dirty flag. Used for check-ins and goals.
- **Optimistic UI:** Local state updated immediately, server call follows.
- **View stack:** Single HTML container with views toggled via `.active` CSS class.
- **Confirmation dialogs:** Created dynamically via `showConfirmDialog()` helper.

## Backend Helpers (code.gs)

- `findRowById_(sheet, id)` — Finds a row by ID column, returns `{rowIndex, colIdx}` or null
- `buildColIdx_(headers)` — Converts header array to `{headerName: 1-based-column-index}` map
- `batchSetValues_(sheet, rowIndex, colIdx, fields)` — Updates multiple cells in one row using a field map
- `normalizeRole_(role)` — Normalizes legacy 'owner' role to 'caseload-manager'

## Frontend Helpers (JavaScript.html)

- `getStudentById(id)` — Finds student in `appState.dashboardData`
- `recalcTotalMissing(student)` — Recalculates totalMissing from academicData
- `modifyMissingAssignment(studentId, classIdx, assignmentIdx, opts)` — Shared logic for mark-done/remove operations
- `updateAutosaveIndicator(elementId, state)` — Unified autosave status indicator
- `showConfirmDialog(opts)` — Reusable confirmation dialog
- `renderErrorState(message)` — Standardized error HTML block
- `renderMissingTable(assignments, studentId, classIdx, context)` — Shared missing assignments table renderer
- `closeConfirmDialog(overlayEl)` — Animated dialog close with timeout fallback
- `animateContentIn(container)` — Cross-fade from skeleton to content
- `createRipple(target, clientX, clientY)` — MD3 ripple effect on buttons (delegated click handler)
- `updateTabIndicator()` — Positions the sliding tab indicator on the active nav tab
- `getEvalTypeLabel(type)` — Resolves eval type value (including legacy aliases) to display label
- `setEvalMenuVisibility_(hasEval)` — Toggles create/view menu items for eval checklists
- `updateEvalBreadcrumb_(typeLabel)` — Updates eval checklist breadcrumb text
- `buildEvalTypeDropdown_(evalId, currentType)` — Builds `<select>` HTML for eval type picker

## UI Design System — Material Design 3

**Reference:** [Material Design 3 Web Components](https://github.com/material-components/material-web/tree/main/docs) — canonical docs for MD3 patterns, tokens, and component specs.

### Seed Color

Cardinal Red `#C41E3A` (Richfield school brand)

### Design Tokens (CSS Custom Properties)

| Token Group | Key Variables |
|---|---|
| **Primary** | `--md-primary: #C41E3A`, `--md-on-primary: #FFF`, `--md-primary-container: #FFDAD6` |
| **Secondary** | `--md-secondary: #775656` (desaturated cardinal) |
| **Tertiary** | `--md-tertiary: #755A2F` (warm gold accent) |
| **Error** | `--md-error: #BA1A1A`, `--md-error-container: #FFDAD6` |
| **Surface** | `--md-surface: #FFFBFF`, levels: lowest → low → container → high → highest |
| **Outline** | `--md-outline: #857373`, `--md-outline-variant: #D8C2C2` |
| **Shape** | xs: 4px, sm: 8px, md: 12px, lg: 16px, xl: 28px, full: 9999px |
| **Elevation** | Levels 1-3 via box-shadow |
| **Motion — Easing** | `--md-easing-standard`, `--md-easing-emphasized-decelerate`, `--md-easing-emphasized-accelerate` |
| **Motion — Duration** | short1–4 (50–200ms), medium1–4 (250–400ms), long1–2 (450–500ms) |

### Typography (Roboto)

| Role | Size/Weight | Usage |
|---|---|---|
| Display Small | 36px/400 | GPA display in side panel |
| Headline Medium | 28px/400 | View titles |
| Headline Small | 24px/400 | Section headers |
| Title Large | 22px/400 | Side panel header, nav brand |
| Title Medium | 16px/500 | Form section titles |
| Title Small | 14px/500 | Table headers, tab labels |
| Label Large | 14px/500 | Buttons, chip text |
| Label Medium | 12px/500 | Breadcrumbs, field labels |
| Body Large | 16px/400 | Form inputs, goal text |
| Body Medium | 14px/400 | Table cells, default text |
| Body Small | 12px/400 | Hints, dates, subtitles |

### MD3 Components Implemented

- **Top App Bar (Small):** Sticky, cardinal red, with logo and Team button
- **Primary Tabs:** Caseload Dashboard / My Caseload / Admin with sliding JS indicator + CSS border-bottom fallback
- **Data Table:** Outlined, sortable headers, hover states, action buttons
- **Outlined Text Fields:** With focus ring transition to primary color
- **Buttons:** Filled (primary), Tonal (secondary), Text/Ghost, Icon, Danger variants — with ripple effect
- **Chips:** Assist-style status badges (green/yellow/red/gray)
- **Cards:** Outlined with expandable detail sections; profile stat cards support tier-colored backgrounds (see GPA Color Tiers below)
- **Side Sheet:** Right-side panel with overlay scrim
- **Dialog:** Centered confirmation with overlay, scale+fade animation, `closeConfirmDialog()` with timeout fallback
- **Snackbar/Toast:** Bottom notification with auto-dismiss, slide+fade animation, debounced via `_toastTimer`
- **Segmented Buttons:** Rating button groups (1-5)
- **Skeleton Loading:** Shimmer animation placeholders per-view, cross-fade to content via `animateContentIn()`
- **Dropdown Menu:** Positioned below trigger with shadow, CSS opacity/transform animation via `.dropdown-open`

### GPA Color Tiers

GPA values are color-coded consistently across three surfaces using MD3 container-style color pairs (background + on-container text). The **Honor Roll** tier uses gold derived from the tertiary palette (`--md-tertiary: #755A2F`).

| Tier | Threshold | Background | Text | Indicator |
|---|---|---|---|---|
| **Honor Roll** | GPA >= 3.5 | `#FFDEAB` (gold) | `#2A1800` | Trophy emoji at GPA >= 3.7 |
| **Good Standing** | GPA >= 2.5 | `#D6F5D6` (green) | `#1B5E20` | — |
| **At Risk** | GPA < 2.5 | `--md-error-container` | `--md-on-error-container` | — |
| **No Data** | null | `--md-surface-container` | `--md-on-surface-variant` | Displays "--" |

**Where applied:**

| Surface | CSS Classes | Notes |
|---|---|---|
| **Dashboard table** | `.chip-gold`, `.chip-green`, `.chip-red` | GPA chip; trophy `&#x1F3C6;` appended inline at >= 3.7 |
| **Side panel** | `.sp-gpa.high`, `.sp-gpa.mid`, `.sp-gpa.low`, `.sp-gpa.none` | Large display (Display Small typography) |
| **Profile stat card** | `.gpa-honor-roll`, `.gpa-good-standing`, `.gpa-at-risk` | Card-level background + border; trophy via `.profile-trophy.visible` at >= 3.7 |

**When adding new color-coded indicators**, follow this pattern:
1. Use MD3 container/on-container color pairs — never raw colors for text without a matching container background
2. Derive tier colors from the existing palette (tertiary-gold for positive, green for satisfactory, error tokens for concern)
3. Keep thresholds consistent across all surfaces (3.5 for honor roll, 2.5 for good standing)
4. Update dynamically: card classes are swapped via `classList.remove()/add()` when data changes without a full re-render

### General Status Chip Colors

Status chips (`.chip`) use a consistent traffic-light scheme across the app for EF ratings, GPA, and missing assignments:

| Class | Background | Text | Usage |
|---|---|---|---|
| `.chip-gold` | `#FFDEAB` | `#2A1800` | Honor Roll: GPA >= 3.5 |
| `.chip-green` | `#D6F5D6` | `#1B5E20` | Good: GPA >= 2.5, EF rating >= 4, missing = 0 |
| `.chip-yellow` | `#FFF3CD` | `#7A5900` | Caution: EF rating 3–3.9, missing 1–3 |
| `.chip-red` | `--md-error-container` | `--md-on-error-container` | Concern: GPA < 2.5, EF rating < 3, missing >= 4 |
| `.chip-gray` | `--md-surface-container-high` | `--md-on-surface-variant` | No data available |

### Responsive Breakpoints

- Mobile: stacked layouts, full-width inputs
- Touch targets: minimum 40px height for buttons/controls
- Side panel: `max-width: 90vw` on small screens

## Animation & Motion Patterns

### View Transitions
Shared Axis X pattern: forward views slide in from right, backward from left. Managed by `showView()` which tracks `_currentViewId` and `VIEW_ORDER` to determine direction.

### Ripple Effect
Delegated click handler creates `<span class="ripple-effect">` inside buttons. Buttons get `overflow: hidden` to contain the ripple, **except** `.rating-btn` (segmented buttons need `overflow: visible` for overlapping borders).

### Animation Safety Rules
- **Never set explicit `opacity: 0`** on elements that rely on animation to become visible (e.g., `.stagger-row`). Use `animation-fill-mode: both` instead — if the animation fails, the element falls back to visible.
- **Always provide CSS fallbacks** for JS-driven visual features (e.g., tab indicator has both sliding JS indicator and `border-bottom` CSS fallback).
- **Always add timeout fallbacks** for `animationend`-dependent cleanup (e.g., `closeConfirmDialog` uses 300ms timeout).
- **Debounce repeated animations** (e.g., `showToast` clears previous timer before starting a new one).
- **Respect `prefers-reduced-motion`**: global `@media` rule reduces all durations to `0.01ms`.

## Code Patterns & Anti-Patterns

### Single Source of Truth for Enumerated Values

**Do:** Define label arrays/maps in one place and derive everything else.

```javascript
// One definition — all labels, dropdowns, and lookups derive from this
var EVAL_TYPES = [
  { value: 'annual-iep', label: 'Annual IEP' },
  ...
];
var EVAL_TYPE_ALIASES = { 'eval': 'initial-eval', 'reeval': '3-year-reeval' };

function getEvalTypeLabel(type) {
  var resolved = EVAL_TYPE_ALIASES[type] || type;
  for (var i = 0; i < EVAL_TYPES.length; i++) {
    if (EVAL_TYPES[i].value === resolved) return EVAL_TYPES[i].label;
  }
  return 'Eval';
}
```

**Don't:** Duplicate labels in separate objects or scatter inline ternaries like `type === 'eval' ? 'Eval' : 'Re-eval'` across multiple files. When a new type is added, every ternary must be found and updated.

### Extract Repeated DOM Manipulation into Helpers

**Do:** Create a helper when the same set of DOM operations appears more than once.

```javascript
// One call replaces 6+ getElementById + style.display lines
function setEvalMenuVisibility_(hasEval) {
  var createIds = ['menu-create-annual-iep', 'menu-create-3-year-reeval', 'menu-create-initial-eval'];
  createIds.forEach(function(id) { ... });
  ...
}
```

**Don't:** Copy-paste blocks of `getElementById()` + `style.display` toggling across success handlers, else branches, and error handlers. Each copy drifts independently when element IDs change.

### Centralize Repeated HTML Fragments

**Do:** Extract HTML builders for fragments used in multiple places (breadcrumbs, dropdowns, badges).

```javascript
function updateEvalBreadcrumb_(typeLabel) { /* one place for the SVG + link HTML */ }
function buildEvalTypeDropdown_(evalId, currentType) { /* one place for <select> construction */ }
```

**Don't:** Inline the same SVG + `innerHTML` assignment in multiple functions. When the markup changes, only one copy gets updated.

### Constants Belong at the Top of the File

**Do:** Declare configuration arrays and validation lists near other constants (e.g., `VALID_EVAL_TYPES` next to `EVALUATION_HEADERS` in `code.gs`).

**Don't:** Declare constants mid-file just above the first function that uses them. They become invisible to anyone scanning the file structure.

### Backend Summary Endpoints Should Return All Relevant Data

**Do:** When building a summary endpoint (like `getEvalTaskSummary`), include all data the frontend might need to decide visibility — e.g., return `activeEvals` alongside aggregate counts.

**Don't:** Return only aggregate metrics (due this week, overdue count) that require items to have due dates set. If the frontend hides the entire section when aggregates are zero, newly created records with no due dates become invisible.

### Use `.view-title-row` for Header + Action Layouts

The existing `.view-title-row` utility class (`display: flex; align-items: center; justify-content: space-between`) is the standard pattern for placing a view title alongside action buttons (e.g., student name + SpEdForms link). Don't create standalone button rows above content — embed the action in the title row to reduce whitespace.

## Development Notes

- Google Apps Script uses V8 runtime (`const`/`let`/`arrow functions` are supported in .gs files)
- HTML files use `<script>` / `<style>` tags and are included via `<?!= include('filename') ?>`
- `google.script.run` is the async bridge between frontend and backend
- Private backend functions use trailing underscore convention (e.g., `getSS_()`)
- XSS prevention via `esc()` function for all user-generated content in HTML
- No external JS/CSS libraries — everything is hand-coded
- **GAS iframe quirk:** `<button>` elements require `appearance: none` to strip native OS chrome. The global reset in `Stylesheet.html` (`button { appearance: none; -webkit-appearance: none; }`) handles this — do not remove it. Without it, buttons render with default browser styling inside the GAS sandbox.

### HtmlService Gotchas

- **HTML entities inside `<script>` blocks:** `HtmlService.createHtmlOutputFromFile().getContent()` may misprocess decimal numeric entities (e.g., `&#127942;`) inside `<script>` tags, corrupting the JavaScript output and breaking the page. Always use **hex entities** (`&#x1F3C6;`) instead of decimal (`&#127942;`) for characters inside script content. Existing hex entities like `&#x1F4AF;` work correctly. This applies to **all** decimal numeric entities regardless of codepoint — even low-range characters like `&#9650;` (▲) must use hex form (`&#x25B2;`). Corruption may not manifest until a seemingly unrelated code change (file size growth, structural refactoring) alters how HtmlService processes the file, making the root cause hard to trace.
- **Null-guard `.style` accesses:** Always null-check `getElementById()` results before accessing `.style` in initialization code (`enterApp`, `showInviteModal`). If `HtmlService` corrupts the template output, DOM elements may not exist.
- **Null-guard `showView()` targets:** Always null-check `getElementById()` in `showView()` before accessing `.classList`. If the target view element is missing, `showView()` removes `.active` from all views first — a subsequent null-access crash leaves every view hidden, producing a blank page with no recovery path.
