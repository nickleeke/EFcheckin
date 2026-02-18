# Caseload Dashboard — CLAUDE.md

## Project Overview

A Google Apps Script web application for Richfield Public Schools educators to manage student caseloads, track executive function (EF) skills, monitor academic progress, and coordinate with co-teachers.

## Tech Stack

- **Backend:** Google Apps Script (`code.gs`, ~3,000 lines)
- **Frontend:** Vanilla JavaScript (`JavaScript.html`, ~9,100 lines)
- **Styling:** Vanilla CSS implementing Material Design 3 (`Stylesheet.html`, ~6,700 lines)
- **Tests:** GAS-compatible function-based test suite (`Tests.gs`, ~1,300 lines)
- **Data:** Google Sheets (per-user, auto-provisioned in Drive)
- **Auth:** Google Session API (`Session.getActiveUser()`)
- **Hosting:** Google Apps Script web app deployment
- No external JS/CSS libraries — everything is hand-coded

### File Structure

```
code.gs            — Backend: CRUD, auth, caching, team management, progress reports
Tests.gs           — GAS-compatible test suite (function-based assertions)
Index.html         — HTML shell: nav, views, side panel, toast
JavaScript.html    — Frontend: state, rendering, API calls, forms
Stylesheet.html    — CSS: MD3 design tokens, all component styles
README.md          — Multi-user implementation guide
```

### Data Flow

```
Frontend (JS) → google.script.run → Backend (Apps Script) → Google Sheets
                  (async)             (sync)                  (storage)
```

## Architecture

### Per-User Data Isolation (FERPA)

Each user gets their own Google Sheet stored in UserProperties. The web app must be deployed as "Execute as: User accessing the web app" so `Session.getActiveUser()` and `UserProperties` are scoped per-user.

### Data Schemas

**Students sheet:** id, firstName, lastName, grade, period, focusGoal, accommodations, notes, classesJson, createdAt, updatedAt, iepGoal, goalsJson, caseManagerEmail

**CheckIns sheet:** id, studentId, weekOf, planningRating, followThroughRating, regulationRating, focusGoalRating, effortRating, whatWentWell, barrier, microGoal, microGoalCategory, teacherNotes, academicDataJson, createdAt, goalMet

**CoTeachers sheet:** email, role, addedAt

**Evaluations sheet:** id, studentId, type, itemsJson, createdAt, updatedAt, filesJson, meetingDate

**ProgressReporting sheet:** id, studentId, goalId, objectiveId, quarter, progressRating, anecdotalNotes, dateEntered, enteredBy, createdAt, lastModified

### Roles

- `caseload-manager` — Full access, can add/remove team members
- `co-teacher` / `service-provider` / `para` — Shared access to caseload data
- Superuser (`nicholas.leeke@rpsmn.org`) — Global admin for case manager assignment

### Eval Types

| Internal Value | Display Label | Template Used |
|---|---|---|
| `annual-iep` | Annual IEP | Re-eval template |
| `3-year-reeval` | 3 Year Re-Eval | Re-eval template |
| `initial-eval` | Initial Eval | Eval template |
| `eval` (legacy) | Initial Eval | Eval template |
| `reeval` (legacy) | 3 Year Re-Eval | Re-eval template |

Definitions: `EVAL_TYPES` array + `EVAL_TYPE_ALIASES` map (frontend), `VALID_EVAL_TYPES` + `EVAL_INITIAL_TYPES_` (backend). Always use `getEvalTypeLabel(type)` for display.

### Caching

**Backend (UserProperties):** Read-through cache with 2-minute TTL. Cache keys: `cache_students`, `cache_dashboard`, `cache_eval_summary`, `cache_progress`, `cache_due_process`. Uses targeted invalidation — each write only clears affected caches (see Backend rules below).

**Frontend (appState):** In-memory cache with 30-second staleness window (`DATA_FRESH_MS`). Timestamps `_lastDashboardFetch`, `_lastEvalSummaryFetch`, `_lastDueProcessFetch` track when data was last fetched. Navigation between views renders from `appState` if data is fresh, avoiding redundant server calls. Post-write handlers render optimistically from `appState` and reset the timestamp to mark data stale for the next navigation.

## Key Patterns

- **Autosave:** Debounced (2s) with in-flight guard and dirty flag. Used for check-ins, goals, and progress entries. Timer must be cleaned up on navigation (see Autosave Safety below).
- **Optimistic UI:** Local state updated immediately, server call follows.
- **View stack:** Single HTML container with views toggled via `.active` CSS class. `showView()` manages Shared Axis X transitions using `_currentViewId` and `VIEW_ORDER`.
- **Confirmation dialogs:** Created dynamically via `showConfirmDialog()` helper.
- **Caseload filter:** `appState.caseloadFilter` ('all' or 'my') controls which students are displayed. All dashboard renderers receive pre-filtered data via `applyFilter_()`. The My Caseload drawer item redirects to `setCaseloadFilter('my')` rather than rendering a separate panel.
- **Dashboard enrichment:** `getDashboardData()` returns `daysSinceCheckIn` and `efHistory` (up to 6 weekly EF averages, oldest first) per student, used for urgency scoring and sparkline rendering.
- **Keyboard shortcuts:** Dashboard supports arrow keys, Enter (side panel), `n` (check-in), `p` (profile), `?` (help). Guards against input/textarea/select focus. `highlightedRowIndex` resets on every `renderDashboard` call.
- **Frontend staleness guard:** `loadDashboard(force)`, `loadEvalSummary(force)`, and `showDueProcess()` skip server calls if data was fetched within `DATA_FRESH_MS` (30s). Pass `force=true` after writes that need guaranteed fresh data (e.g., `doDeleteStudent`). Post-write success handlers should use `renderDashboardFromState(); _lastDashboardFetch = 0;` instead of `loadDashboard()` to avoid skeleton flash and redundant re-fetch.

## Helper Functions

### Backend (code.gs)

| Function | Purpose |
|---|---|
| `findRowById_(sheet, id)` | Find row by ID column, returns `{rowIndex, colIdx}` or null |
| `buildColIdx_(headers)` | Header array to `{name: 1-based-column}` map |
| `batchSetValues_(sheet, rowIndex, colIdx, fields)` | Update multiple cells in one row |
| `normalizeRole_(role)` | Normalize legacy 'owner' to 'caseload-manager' |
| `appendStudentNote(studentId, noteText)` | Append timestamped note, invalidate cache |
| `saveProgressEntry(data)` | Upsert progress for goal+objective+quarter with LockService |
| `getProgressEntries(studentId, quarter)` | Fetch progress entries for one student+quarter |
| `getAllProgressForStudent(studentId)` | Fetch all progress across all quarters |
| `assembleReportData_(student, quarter, allEntries)` | Pure data assembly for report with goalArea grouping |
| `generateProgressReportHtml_(student, quarter, allEntries, summary)` | Full printable HTML report with inline CSS |
| `generateProgressReport(studentId, quarter, summary)` | Public endpoint: single student report |
| `generateBatchReports(quarter, summaries)` | Public endpoint: reports for entire caseload |
| `getCurrentQuarter()` | Returns Q1-Q4 based on current date |
| `invalidateStudentCaches_()` | Clear `cache_students` + `cache_dashboard` (student writes) |
| `invalidateCheckInCaches_()` | Clear `cache_dashboard` (check-in writes) |
| `invalidateEvalCaches_()` | Clear `cache_eval_summary` + `cache_dashboard` + `cache_due_process` (eval writes) |
| `invalidateMeetingCaches_()` | Clear `cache_eval_summary` + `cache_due_process` (meeting writes) |
| `invalidateProgressCaches_()` | Clear `cache_progress` + `cache_due_process` (progress writes) |
| `invalidateCache_()` | Nuclear: clear all 5 caches (deleteStudent, team ops, force-refresh only) |

### Frontend — General Utilities (JavaScript.html)

| Function | Purpose |
|---|---|
| `esc(str)` | XSS-safe HTML escaping via DOM textContent |
| `getStudentById(id)` | Find student in `appState.dashboardData` |
| `recalcTotalMissing(student)` | Recalculate totalMissing from academicData |
| `modifyMissingAssignment(studentId, classIdx, assignmentIdx, opts)` | Shared mark-done/remove logic |
| `showConfirmDialog(opts)` | Reusable confirmation dialog |
| `closeConfirmDialog(overlayEl)` | Animated dialog close with timeout fallback |
| `renderErrorState(message)` | Standardized error HTML block |
| `animateContentIn(container)` | Cross-fade from skeleton to content |
| `createRipple(target, clientX, clientY)` | MD3 ripple effect (delegated click handler) |
| `updateTabIndicator()` | Position sliding tab indicator on active nav tab |

### Frontend — Dashboard Features

| Function | Purpose |
|---|---|
| `computeUrgencyScore(s)` | Integer priority score from daysSinceCheckIn, avgRating, totalMissing, trend |
| `formatRelativeDate(dateStr)` | Returns `{text, tier}` — "3d ago" green / "1 wk ago" yellow / "2 wks ago" red |
| `buildSparklineSvg(efHistory)` | Inline SVG sparkline (60x24) with colored endpoint dot |
| `applyFilter_(data)` | Filter by caseManagerEmail when `caseloadFilter === 'my'` |
| `setCaseloadFilter(filter)` | Toggle chip state, re-render all dashboard sections |
| `renderNeedsAttention(data)` | Urgency-sorted card strip or "all on track" empty state |
| `renderCheckInProgress(data)` | "X of Y checked in this week" with progress bar |
| `renderMissingAggregate(data)` | Total missing metric card with expandable student list |
| `toggleKeyboardHelp()` | Show/hide keyboard shortcut overlay |
| `renderDashboardFromState()` | Render dashboard cards from `appState` without server call |
| `loadDashboard(force)` | Fetch + render dashboard; skips fetch if fresh unless `force=true` |
| `loadEvalSummary(force)` | Fetch + render eval summary; skips fetch if fresh unless `force=true` |

### Frontend — Eval Helpers

| Function | Purpose |
|---|---|
| `getEvalTypeLabel(type)` | Resolve eval type (including legacy aliases) to display label |
| `setEvalMenuVisibility_(hasEval)` | Toggle create/view menu items for eval checklists |
| `updateEvalBreadcrumb_(typeLabel)` | Update eval checklist breadcrumb text |
| `buildEvalTypeDropdown_(evalId, currentType)` | Build `<select>` HTML for eval type picker |
| `renderMissingTable(assignments, studentId, classIdx, context)` | Shared missing assignments table renderer |

## UI Design System — Material Design 3

**Reference:** [Material Design 3 Web Components](https://github.com/material-components/material-web/tree/main/docs)

### Brand Palette (Richfield Public Schools)

| Swatch | Name | HEX | RGB | Usage |
|---|---|---|---|---|
| **Primary** | Richfield Red | `#942022` | 148-32-34 | MD3 seed color, top app bar, primary actions |
| **Primary** | Helmet Gray | `#797a7d` | 121-122-125 | Neutral accents |
| **Primary** | Black | `#000000` | — | Text, icons |
| **Primary** | White | `#FFFFFF` | — | Backgrounds, on-primary text |
| **Secondary** | Navy | `#21376c` | 33-55-108 | — |
| **Secondary** | Gold | `#e8b34b` | 232-179-75 | — |
| **Secondary** | Teal | `#4ea3a8` | 78-163-168 | Accent surfaces |
| **Secondary** | Green | `#79af61` | 121-175-97 | — |
| **Secondary** | Orange | `#cf5d35` | 207-93-53 | — |

### Design Tokens (CSS Custom Properties)

The MD3 token system is derived from the Richfield Red seed color `#942022`, with adjustments for WCAG contrast and MD3 tonal palette generation.

| Token Group | Key Variables |
|---|---|
| **Primary** | `--md-primary: #C41E3A`, `--md-on-primary: #FFF`, `--md-primary-container: #FFDAD6` |
| **Secondary** | `--md-secondary: #775656` (desaturated cardinal) |
| **Tertiary** | `--md-tertiary: #755A2F` (warm gold accent) |
| **Error** | `--md-error: #BA1A1A`, `--md-error-container: #FFDAD6` |
| **Surface** | `--md-surface: #FAFAFA`, levels: lowest / low / container / high / highest |
| **Outline** | `--md-outline: #757575`, `--md-outline-variant: #C8C8C8` |
| **Shape** | xs: 4px, sm: 8px, md: 12px, lg: 16px, xl: 28px, full: 9999px |
| **Elevation** | Levels 1-3 via box-shadow |
| **Motion** | Easing: `standard`, `emphasized-decelerate`, `emphasized-accelerate`. Duration: short1-4 (50-200ms), medium1-4 (250-400ms), long1-2 (450-500ms) |
| **Color Tiers** | `--color-tier-gold-bg/text`, `--color-tier-green-bg/text`, `--color-tier-yellow-bg/text` |

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

### MD3 Components

**Layout & Navigation:**
- **Top App Bar (Small):** Sticky, cardinal red. Contains centered search bar. Mobile (≤840px): hamburger + search bar. Desktop (≥841px): 48px with search bar only (hamburger hidden).
- **Navigation Drawer:** Side nav with collapsible groups, tooltip fallback when collapsed. Bottom section has Shortcuts item (keyboard help). Collapse toggle at bottom (desktop only).
- **Primary Tabs:** With sliding JS indicator + CSS border-bottom fallback
- **Side Sheet:** Right-side panel with overlay scrim

**Inputs & Controls:**
- **Outlined Text Fields:** With focus ring transition to primary color
- **Buttons:** Filled, Tonal, Text/Ghost, Icon, Danger variants — with ripple effect
- **Segmented Buttons:** Rating button groups (1-5) with color-coded selected states (red/yellow/green by value via `[data-value]` attribute selectors), goal-met Yes/Partially/No with semantic colors (green/yellow/red via `[data-goal-met]` selectors)
- **Filter Chips:** Toggle buttons (`.filter-chip` / `.filter-chip-active`) with `aria-pressed`

**Data Display:**
- **Data Table:** Outlined, sortable headers (priority/name/grade), hover states, sparklines, staleness chips
- **Chips:** Status badges (gold/green/yellow/red/gray) — see Color Tiers below
- **Cards:** Outlined with expandable sections; metric cards with `toggleMetricDropdown`; needs-attention cards (left red accent stripe, not full border)
- **Sparklines:** Inline SVG mini-charts (`.ef-sparkline`) in table cells
- **Staleness Chips:** Color-coded relative date (`.staleness-green/yellow/red`)
- **Progress Bar:** Thin linear indicator (`.checkin-progress-fill`) with ARIA progressbar

**Feedback:**
- **Dialog:** Confirmation with overlay, scale+fade, `closeConfirmDialog()` with timeout fallback
- **Snackbar/Toast:** Bottom notification with auto-dismiss, debounced via `_toastTimer`
- **Skeleton Loading:** Diagonal shimmer placeholders, cross-fade via `animateContentIn()`
- **Keyboard Help Dialog:** Accessible via `?` key, nav drawer Shortcuts button, or `toggleKeyboardHelp()`

### Color Tiers

All color-coded indicators use the same tier system across the app (GPA chips, staleness, EF ratings, missing assignments, rating buttons, goal-met buttons). Always use the CSS custom properties — never hardcode hex values.

| Tier | CSS Token (bg) | CSS Token (text) | Hex Values | Thresholds |
|---|---|---|---|---|
| **Gold** | `var(--color-tier-gold-bg)` | `var(--color-tier-gold-text)` | `#FFDEAB` / `#2A1800` | GPA >= 3.5. Trophy at >= 3.7 |
| **Green** | `var(--color-tier-green-bg)` | `var(--color-tier-green-text)` | `#D6F5D6` / `#1B5E20` | GPA 3.0-3.5, EF >= 4, missing = 0, staleness <= 6d |
| **Yellow** | `var(--color-tier-yellow-bg)` | `var(--color-tier-yellow-text)` | `#FFF3CD` / `#7A5900` | GPA 2.5-3.0, EF 3-3.9, missing 1-3, staleness 7-13d |
| **Red** | `var(--md-error-container)` | `var(--md-on-error-container)` | — | GPA < 2.5, EF < 3, missing >= 4, staleness 14d+ |
| **Gray** | `var(--md-surface-container-high)` | `var(--md-on-surface-variant)` | — | No data |

**CSS classes by surface:**

| Surface | Classes |
|---|---|
| Dashboard table chips | `.chip-gold`, `.chip-green`, `.chip-yellow`, `.chip-red`, `.chip-gray` |
| Side panel GPA | `.sp-gpa.high`, `.sp-gpa.mid`, `.sp-gpa.caution`, `.sp-gpa.low`, `.sp-gpa.none` |
| Profile stat cards | `.gpa-honor-roll`, `.gpa-good-standing`, `.gpa-caution`, `.gpa-at-risk` |
| Staleness chips | `.staleness-green`, `.staleness-yellow`, `.staleness-red` |

When adding new color-coded indicators, use `var(--color-tier-*)` tokens. Keep thresholds consistent across surfaces. Update dynamically via `classList.remove()/add()`. Never hardcode tier hex values in Stylesheet.html — they are defined once in `:root`.

### Responsive Breakpoints

- `>= 841px` (desktop): Nav drawer always visible (collapsible). Compact 48px top bar with centered search. Persistent action bar hidden.
- `<= 840px` (mobile): Nav drawer as overlay with scrim. Full 64px top bar with hamburger. Persistent action bar visible below top bar.
- `<= 600px`: Hide progress counter, sparklines; shrink filter chips and attention cards. Top bar shrinks to 56px.
- Touch targets: minimum 40px height for buttons/controls
- Side panel: `max-width: 90vw` on small screens

## Code Rules

### Single Source of Truth

Define constants, label arrays, and lookup maps in one place. Never redeclare local copies with different values.

- **Eval types:** Use `EVAL_TYPES` + `EVAL_TYPE_ALIASES` + `getEvalTypeLabel()`. Never inline ternaries.
- **GPA conversion:** Use the global `GPA_MAP` constant. Never redeclare a local `gradePoints` with different precision.
- **Constants location:** Declare at the top of the file near other constants. Never mid-file above first usage.

### Dashboard Rendering

- **Always filter:** Every `renderDashboard()` call must use `applyFilter_()`: `renderDashboard(applyFilter_(appState.dashboardData))`. This applies to `toggleSort`, `saveCheckIn`, `goBackFromMissing`, `saveStudentForm`, `doDeleteStudent`, `refreshDashboard`.
- **Metric dropdown IDs:** Expandable cards must use `id="metric-card-{key}"` and `id="metric-dropdown-{key}"` for `toggleMetricDropdown(key)` to find them.

### DOM & Accessibility

- **Clickable `<div>` elements** must have `role="button" tabindex="0" onkeydown="if(event.key==='Enter')..."`.
- **Toggle buttons** (filter chips, segmented buttons) must have `aria-pressed`.
- **Progress indicators** must have `role="progressbar"` with `aria-valuenow/min/max`.
- **Extract repeated DOM operations** into helpers. Don't copy-paste `getElementById` + `style.display` blocks.
- **Extract repeated HTML fragments** into builder functions (`buildEvalTypeDropdown_`, `updateEvalBreadcrumb_`).
- **Use `.view-title-row`** for header + action button layouts.

### Ripple Effect

Delegated click handler creates `<span class="ripple-effect">` inside buttons. All button classes must be in the ripple overflow selector: `.btn, .btn-icon, .btn-ghost, .btn-outlined, .action-btn, .rating-btn, .nav-drawer-item, .goal-met-btn, .eval-filter-chip, .eval-nav-btn, .dp-cal-day, .header-search-result-item, .dp-progress-checkbox-row, .dp-progress-flat-card`. Exception: `.rating-btn` needs `overflow: visible` for overlapping borders.

When adding a new button class, add it to this selector in `Stylesheet.html`.

### CSS Rules

- **`.selected:hover`:** Always define a `:hover` rule for `.selected` states — otherwise the base `:hover` overrides the selected background.
- **Animation safety:** Never set explicit `opacity: 0` on elements that rely on animation. Use `animation-fill-mode: both`. Always add timeout fallbacks for `animationend`. Debounce repeated animations.
- **Respect `prefers-reduced-motion`:** Global `@media` rule reduces all durations to `0.01ms`.
- **Motion tokens:** Use `--md-duration-*` and `--md-easing-*` variables. Never hardcode durations.
- **Color tier tokens:** Use `var(--color-tier-*)` for gold/green/yellow tier colors. Never hardcode `#FFDEAB`, `#D6F5D6`, `#FFF3CD`, etc. in Stylesheet.html. Tokens are defined in `:root`. JS inline styles (sparkline dots, progress report CSS strings) may use raw hex since CSS vars aren't available there.
- **No single-side strokes:** Never use `border-left`, `border-top`, etc. as accent indicators. Always use a full `border` around the entire element. For emphasis, use color (e.g., `border: 1px solid var(--md-error)`) rather than a thicker single-side stripe.

### Autosave Safety

When adding new autosave-enabled features (like progress entry):
- **Clean up timers on navigation:** Clear the timer, reset in-flight and dirty flags at the start of the parent view function (e.g., `showStudentProfile()`). Otherwise the timer fires for the wrong student/context.
- **Guard completion handlers:** In async completion callbacks, verify the user is still viewing the same student (`studentId === appState.currentStudentId`) before refreshing state or rescheduling.
- **Cancel on context switch:** Clear the timer when the user switches sub-contexts (e.g., switching quarters within the same student view).

### Concurrent Write Safety

For sheet upsert operations (read-check-write patterns), wrap in `LockService.getUserLock()` with `waitLock()` and `releaseLock()` in a try-finally block. Without this, rapid concurrent saves can create duplicate rows.

### Backend

- **Summary endpoints** should return all data the frontend needs for visibility decisions, not just aggregate counts.
- Private functions use trailing underscore convention (`getSS_()`, `findRowById_()`).
- **Targeted cache invalidation:** After writes, call the narrowest invalidation helper (`invalidateStudentCaches_()`, `invalidateCheckInCaches_()`, `invalidateEvalCaches_()`, `invalidateMeetingCaches_()`, `invalidateProgressCaches_()`). Only use nuclear `invalidateCache_()` for operations that affect all data (deleteStudent, team changes, force-refresh).
- **Validate inputs at public endpoints:** Check `VALID_QUARTERS`, `VALID_PROGRESS_RATINGS`, student existence, and anecdotal note length before writing.

## Development Notes

- Google Apps Script uses V8 runtime (`const`/`let`/`arrow functions` supported in .gs)
- HTML files included via `<?!= include('filename') ?>`
- `google.script.run` is the async bridge between frontend and backend
- XSS prevention via `esc()` function for all user-generated content in HTML
- **GAS iframe quirk:** `<button>` elements require `appearance: none`. The global reset (`button { appearance: none; -webkit-appearance: none; }`) handles this — do not remove it.

### HtmlService Gotchas

### Testing (Tests.gs)

- **No test framework:** Uses plain GAS functions with `assert_()`, `assertEqual_()`, `assertContains_()`, `assertNotNull_()` helpers.
- **Runner pattern:** `runAllProgressReportTests()` uses an explicit function map (not `this[name]()` — `this` inside `forEach` doesn't reference global scope in GAS). Always add new test functions to both the `tests` array and the `testFns` map.
- **Data cleanup:** Tests that write to sheets must use try-finally blocks calling `deleteProgressEntry_(id)` to clean up, even when assertions fail.
- **Naming convention:** `test_category_expectedBehavior` (e.g., `test_gpa_calculatesCorrectly`).

### HtmlService Gotchas

- **Hex entities only in `<script>` blocks:** Always use hex (`&#x1F3C6;`) not decimal (`&#127942;`). Decimal entities can corrupt JavaScript output. This applies to all codepoints, even low-range (`&#x25B2;` not `&#9650;`). Corruption may not manifest until a file size or structure change.
- **Null-guard all DOM access:** Always null-check `getElementById()` before `.style` or `.classList`. If HtmlService corrupts output, elements may not exist. Especially critical in `showView()` — a null crash there hides all views with no recovery.
