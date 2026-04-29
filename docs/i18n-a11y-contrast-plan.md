# Renderer UI: i18n, Accessibility & Contrast Fixes

## Context

The new renderer-based unified window (V3) replaces the old Mithril-based injected sidebar. While the new UI uses better semantic HTML (`<button>`, `<textarea>`, `<label>`), it regressed in three areas compared to the old UI:

1. **i18n** — ~15 hardcoded English strings that the old UI localized via `Localization.getLocalizedString()`
2. **Accessibility** — Lost ARIA state attributes, keyboard navigation, focus outlines, and aria-live regions
3. **Contrast** — Error text color `#ff6b6b` fails WCAG AA; region button border fails non-text contrast

The old UI's patterns are the blueprint — most string keys already exist in `strings.json`, and the ARIA patterns are well-documented in the old components.

### How i18n works in this extension

Strings are fetched from `https://www.onenote.com/strings?ids=WebClipper.&locale={locale}` at startup by `extensionBase.ts`, stored in `localStorage.locStrings`. The renderer reads them via `loc(key, fallback)`. The `strings.json` file in the repo is the **English fallback only** — actual translations for 60 locales come from the server. New keys added to `strings.json` will only have English until the server is updated, so **reuse existing keys wherever possible**.

### How locale is detected

`extensionBase.ts` line 133: `navigator.language || navigator.userLanguage` → stored in `localStorage.locale`. A user override is also supported via `localStorage.displayLocaleOverride`. The old UI only set `<html lang="en">` statically — never dynamically. RTL was handled by loading separate CSS files (`clipper-rtl.css`), not via `lang`/`dir` attributes.

## Files Modified

| File | Changes |
|------|---------|
| `src/renderer.html` | `lang` attr, ARIA roles/attributes, `for` associations, aria-live region |
| `src/scripts/renderer.ts` | Wire `loc()` to all remaining strings, dynamic `lang` attr, ARIA state management, keyboard nav, focus management, aria-live announcements |
| `src/styles/renderer.less` | Focus outlines, error color fix, region button contrast, sr-only class, high-contrast media query |

---

## Phase 1: i18n (mirror old `Localization.getLocalizedString()` via existing `loc()`)

### 1a. Wire sign-in panel to `loc()` in renderer.ts

Sign-in panel HTML strings (`signin-description`, `signin-msa-btn`, `signin-orgid-btn`, `signin-progress`) were never replaced by JS. Now wired to existing keys:

- `WebClipper.Label.SignInDescription` → sign-in description
- `WebClipper.Action.SigninMsa` → MSA button
- `WebClipper.Action.SigninOrgId` → OrgId button
- "Signing in..." kept as hardcoded English (no existing key, brief transient state)

### 1b. Wire field labels to `loc()`

- Note label → `WebClipper.Label.Annotation` ("Note")
- Save to label → `WebClipper.Label.ClipLocation` ("Location")
- Source label → kept as hardcoded "Source" (no old UI equivalent, new element)
- Title label → kept as hardcoded "Title" (old UI had no visible label, used placeholder only)

### 1c. Wire remaining hardcoded strings to `loc()`

| String | Key | Key status |
|--------|-----|-----------|
| `"Capture complete"` | — | **REMOVED** — dead code, never referenced |
| `"No notebooks available"` | `WebClipper.SectionPicker.NoNotebooksFound` | EXISTS (60 locales) |
| `"Error loading notebooks"` | `WebClipper.SectionPicker.NotebookLoadFailureMessage` | EXISTS (60 locales) |
| `"Loading article..."` | `WebClipper.Preview.LoadingMessage` | EXISTS (60 locales) |
| `"Article content not available..."` | `WebClipper.Preview.NoContentFound` | EXISTS (60 locales) |
| `"Unknown error"` | — | kept as-is (technical fallback) |
| `"Sign-in failed..."` | `WebClipper.Error.SignInUnsuccessful` | EXISTS (60 locales) |

### 1d. New keys in strings.json

The original V3 i18n pass added zero new keys (everything mapped to existing server-translated keys). The Fluent 2 redesign that followed introduced several new keys (English fallback used until the translation pipeline picks them up):

| Key | Fallback (English) | Used in |
|-----|--------------------|---------|
| `WebClipper.Label.WhatToCapture` | "What do you want to capture?" | Caption above mode buttons |
| `WebClipper.Action.Discard` | "Discard" | Region thumbnail remove pill |
| `WebClipper.Label.ClipSuccessTitle` | "Saved to your notebook" | Success banner heading |
| `WebClipper.Label.ClipSuccessDescription` | "You can access and continue working on it anytime from your notebook." | Success banner body |
| `WebClipper.Label.ClipErrorTitle` | "Couldn't save to your notebook" | Error banner heading |
| `WebClipper.Label.ClipErrorDescription` | "Something went wrong while saving your clip. Try saving your clip again" | Error banner body |

---

## Phase 2: Contrast Fixes

> **Note**: These fixes were originally calibrated for the dark purple theme. The Fluent 2 white theme that followed brought its own Fluent-tokenized colors, so the specific hex values below are largely superseded. The principles still apply: error text and region button borders meet WCAG AA, focus rings are visible.

### 2a. Error text color (legacy purple theme)

`.signin-error` color: `#ff6b6b` → `#ff9999` (~5.3:1 on `#56197c` purple bg, passed WCAG AA).

**Current** (Fluent 2 white theme): error/danger uses `#a80000` (red) on white surfaces — ~7.7:1, comfortably AA.

### 2b. Region add-button border

Border: `#bbb` → `#999` (~3.4:1 on light gray bg, passed SC 1.4.11).

**Current** (Fluent 2): region "Add another region" button uses Fluent outline pattern — neutral 1px dashed border + Fluent Add icon, hover state turns brand purple.

### 2c. Focus outlines

**Original** (dark purple theme): `outline: solid 1px #f8f8f8 !important; outline-offset: 1px` — light ring on dark sidebar.

**Current** (Fluent 2 white theme): split by element type to avoid doubled borders next to the element's own border:
- **Bordered elements** (`button`, `textarea`, `[tabindex]`, `.signin-btn`): `outline: none; box-shadow: inset 0 0 0 1px purple; border-color: purple` — single 2px purple edge merging the 1px border + 1px inset.
- **Text links** (`<a>`): `outline: 2px solid purple; outline-offset: 2px` — outer ring with gap.
- **Section list items** (no border): `outline: 2px solid purple; outline-offset: -2px` — inset to avoid layout shift in scrollable list.
- **Clip button** (filled primary): Fluent 2 "halo" pattern — `box-shadow: inset 0 0 0 2px #fff, 0 0 0 2px #242424` so the indicator contrasts against both the purple fill (white inner ring) and the white sidebar bg (dark outer ring).
- **High contrast** (`@media (forced-colors: active)`): `outline: solid 2px Highlight !important; outline-offset: 2px !important`

---

## Phase 3: Accessibility — ARIA & Keyboard

### 3a. `<html lang>` attribute (WCAG 3.1.1 Level A)

Static `lang="en"` in HTML + dynamic override in JS reading `localStorage.locale` (or `displayLocaleOverride`). Converts `_` to `-` for BCP 47 (e.g., `zh_CN` → `zh-CN`).

### 3b. Mode buttons — ARIA state

- Container: `role="toolbar"` with localized `aria-label`
- Buttons: standard `<button>` elements with `aria-pressed` (toggle button pattern)
- Reverted from `role="radio"` / radiogroup — buttons are more natural for mode selection

### 3c. Mode buttons — arrow key navigation

Arrow Up/Down/Left/Right + Home/End navigation between mode buttons. Mirrors old `enableAriaInvoke()` from `componentBase.ts`.

### 3d. Section picker — ARIA combobox

- Trigger: `role="combobox"`, `aria-haspopup="listbox"`, `aria-expanded`
- List items: `role="option"`, `aria-selected`
- Escape to close, arrow keys to navigate

### 3e. Label `for` associations

- `<label for="title-field">` and `<label for="note-field">`
- `aria-labelledby` on source-url and section-selected (non-input elements)

### 3f. aria-live regions

- `<div id="aria-status" class="sr-only" aria-live="polite" aria-atomic="true">`
- `announceToScreenReader()` helper for: capture start/complete, mode change, sign-in error
- Save success/error use `role="status"` / `role="alert"` on the banner element itself (auto-announce, no manual aria-live needed — was double-announcing previously)

### 3g. Sign-in dialog

- Sign-in overlay: `role="dialog"` + `aria-modal="true"` + `aria-label="Sign in"` so AT treats it as a real modal
- Z-index 9999 + `isolation: isolate` + explicit `width: 100vw; height: 100vh` to defend against macOS full-window edge cases
- Focus management: focus first sign-in button on overlay show; focus first mode button on overlay hide

### 3h. Collapsible section headings (Fluent redesign)

- Section headings made keyboard-accessible: `role="button"`, `tabindex="-1"`, Enter/Space toggles collapse
- `aria-expanded` reflects state, `data-depth` enables depth-aware sibling visibility toggle

### 3i. Copy diagnostics button

`aria-label="Copy diagnostic information"` (hardcoded English — technical label). Restored after the Fluent inline-alert redesign because keyboard users can't easily select text from a `<pre>` for clipboard copy.

### 3i. Sign-out disabled state

Replace `pointer-events: none` + opacity with `aria-disabled="true"` + `tabindex="-1"` (accessible to keyboard/screen readers).

---

## RTL Support Assessment (Deferred — Separate Effort)

**How old UI handled RTL:**
- `localeSpecificTasks.ts` calls `Rtl.isRtl(locale)` (checks `ar, fa, he, sd, ug, ur`)
- Loads `clipper-rtl.css` / `sectionPicker-rtl.css` instead of LTR versions
- `styledFrameFactory.ts` flips iframe position (`left: 0` instead of `right: 0`)

**Locale override:** No UI exists for switching locale. `localStorage.displayLocaleOverride` is a developer/testing mechanism only (set via console).

**What RTL would need for the renderer (estimated ~50-80 LESS lines + testing):**
1. Detect RTL locale and set `dir="rtl"` on `<html>`
2. Convert `renderer.less` to CSS logical properties (`margin-inline-start/end`, etc.)
3. Flip: sidebar position, section picker arrow, user-info alignment, feedback icon margin
4. Test with at least one RTL locale (ar or he)

**Why defer:** RTL affects layout fundamentals. Best done as a dedicated pass.

---

## Implementation Order

1. **Phase 2a** — Error contrast fix (LESS)
2. **Phase 1a–1c** — i18n wiring (renderer.ts only, reuse existing keys)
3. **Phase 2b–2c** — Remaining contrast + focus outlines (LESS)
4. **Phase 3a** — `<html lang>` + dynamic locale (HTML + TS)
5. **Phase 3e** — Label `for` associations (HTML)
6. **Phase 3b–3c** — Mode button ARIA + arrow keys (TS)
7. **Phase 3d** — Section picker ARIA (HTML + TS)
8. **Phase 3f** — aria-live regions (HTML + LESS + TS)
9. **Phase 3g–3i** — Focus management, copy button label, signout disabled state (TS)

---

## Verification

1. **Build**: `npm run build` — check for TS compilation errors
2. **Edge target**: Verify `renderer.js` and `renderer.css` in target output
3. **Manual testing in Edge** (verified):
   - Sign-in panel: localized text appears
   - Mode buttons: arrow key navigation, `aria-checked` updates in devtools
   - Section picker: `aria-expanded` toggles, Escape closes, arrow keys navigate items
   - Save error: `#ff9999` error text readable
   - Tab order: mode buttons → title → note → section → Clip → Cancel → feedback → sign out (wraps)
   - Focus outlines: `2px solid #f8f8f8` visible on all sidebar controls
   - Focus management: Cancel focused during capture → Full Page button after capture → sign-in button on overlay
4. **No functional regressions**: Capture, save, region, article, bookmark modes all work
5. **Screen reader testing** (verified with NVDA + Edge):
   - ARIA roles, states, and aria-live announcements are implemented and working
   - Accessibility tree verified correct via `edge://accessibility` — all nodes present with proper roles
   - NVDA reads sidebar controls, mode buttons, section picker, and aria-live announcements
   - Blur handler was removed to avoid conflicting with screen reader focus management
   - Keydown handler allows navigation keys (arrows, Tab, Escape, Home/End, PageUp/PageDown) and modifier combos to pass through for screen reader compatibility
   - **Note**: Previous testing on a stale devbox showed Edge not activating accessibility API flags for extension popup windows. A devbox reboot resolved this — NVDA works correctly with the renderer window
