# Caseload Dashboard — CLAUDE.md

## Project Overview

A Google Apps Script web application for Richfield Public Schools educators to manage student caseloads, track executive function (EF) skills, monitor academic progress, and coordinate with co-teachers.

## Tech Stack

- **Backend:** Google Apps Script (`code.gs`)
- **Frontend:** Vanilla JavaScript (`JavaScript.html`, ~3,500 lines)
- **Styling:** Vanilla CSS implementing Material Design 3 (`Stylesheet.html`, ~2,200 lines)
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

## UI Design System — Material Design 3

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
- **Cards:** Outlined with expandable detail sections
- **Side Sheet:** Right-side panel with overlay scrim
- **Dialog:** Centered confirmation with overlay, scale+fade animation, `closeConfirmDialog()` with timeout fallback
- **Snackbar/Toast:** Bottom notification with auto-dismiss, slide+fade animation, debounced via `_toastTimer`
- **Segmented Buttons:** Rating button groups (1-5)
- **Skeleton Loading:** Shimmer animation placeholders per-view, cross-fade to content via `animateContentIn()`
- **Dropdown Menu:** Positioned below trigger with shadow, CSS opacity/transform animation via `.dropdown-open`

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

## Development Notes

- Google Apps Script uses V8 runtime (`const`/`let`/`arrow functions` are supported in .gs files)
- HTML files use `<script>` / `<style>` tags and are included via `<?!= include('filename') ?>`
- `google.script.run` is the async bridge between frontend and backend
- Private backend functions use trailing underscore convention (e.g., `getSS_()`)
- XSS prevention via `esc()` function for all user-generated content in HTML
- No external JS/CSS libraries — everything is hand-coded
- **GAS iframe quirk:** `<button>` elements require `appearance: none` to strip native OS chrome. The global reset in `Stylesheet.html` (`button { appearance: none; -webkit-appearance: none; }`) handles this — do not remove it. Without it, buttons render with default browser styling inside the GAS sandbox.
