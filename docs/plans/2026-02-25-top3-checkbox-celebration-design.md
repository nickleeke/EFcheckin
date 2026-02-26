# Top 3 Widget: Custom Checkmarks, Animated Strikethrough & Celebration

**Date:** 2026-02-25
**Status:** Approved

## Problem

The Today's Top 3 widget checkboxes use a plain square checkbox that renders with check marks visible. The checking interaction lacks delight — no animation on the strikethrough, and no reward when all items are completed.

## Design

### 1. Custom Checkmark Icon Buttons (Dotted → Filled)

Replace the square checkbox + hidden `<input>` with a standalone SVG checkmark icon:

- **Unchecked:** Checkmark path with `stroke-dasharray: 4 3` (dotted outline), `fill: none`, stroke in widget color at 60% opacity
- **Checked:** Smooth transition to `stroke-dasharray: 0` (solid), `fill: widgetColor`, full opacity
- Transition: `var(--md-duration-short4)` + `var(--md-easing-standard)`
- Keep hidden `<input type="checkbox">` for state management; SVG is purely visual

### 2. Animated Strikethrough (Left → Right)

Replace `text-decoration: line-through` with a CSS `::after` pseudo-element:

- `::after` positioned at vertical center, 1.5px height, `background: currentColor`
- Default: `transform: scaleX(0); transform-origin: left`
- `.completed::after`: `transform: scaleX(1)`
- Duration: `var(--md-duration-medium2)` + `var(--md-easing-emphasized-decelerate)`
- Text opacity fades to 0.5 simultaneously

### 3. Lottie Celebration + Glass Blur

**Trigger:** All items with text are checked (`checkAllCompleted_(widget)` returns true).

**Z-index layering (user-specified):**
1. Card content (base)
2. Glass blur — `backdrop-filter: blur(8px)` + `rgba(255,255,255,0.3)` overlay
3. Lottie animation — `<dotlottie-wc>` element centered over card

**Lifecycle:**
- Blur + Lottie fade in: `opacity 0→1`, `medium2` + `emphasized-decelerate`
- Auto-dismiss after ~3.5s
- Fade out: `opacity 1→0`, `medium1` + `emphasized-accelerate`
- DOM cleanup after fade-out transition completes

**Script:** `@lottiefiles/dotlottie-wc` loaded via `<script type="module">` in Index.html.
**Lottie URL:** `https://lottie.host/cf69825e-33ef-4ab2-a5d0-3fdd65b94db0/PJLIvrZACE.lottie`

### Card CSS Adjustments

- Add `position: relative` to `.top3-widget-card` (for absolute overlay positioning)
- May need to bump `max-height` if the dotted checkmarks take more vertical space

## Files Modified

- **Stylesheet.html** — `.top3-checkmark` restyle, `.top3-text::after` animation, `.widget-celebration-blur`, `.widget-celebration-lottie`
- **JavaScript.html** — `renderTop3Collapsed_()` SVG replacement, `toggleTop3Item()` celebration check, new `checkAllCompleted_()` and `showWidgetCelebration_()` functions
- **Index.html** — `<script type="module">` for dotlottie-wc CDN

## Respects Reduced Motion

All animations gate on `_prefersReducedMotion`. When true: instant state changes, no strikethrough animation, no celebration overlay (just a toast "All done!" instead).
