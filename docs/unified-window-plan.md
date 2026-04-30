# Plan: Unified Clipper Window

## Context

The OneNote Web Clipper previously used two separate UI contexts: an injected iframe sidebar on the original page, and a separate renderer popup window for full-page screenshot capture. This created a jarring experience.

**Goal:** Merge the clipper UI into the renderer window so everything lives in one place.

## V1 — Implemented

### Layout
```
Window width: 1280 (content) + 321 (sidebar) = ~1601px
┌──────────────────────────────────┬──────────────┐
│                                  │              │
│  Content iframe (1280px)         │  Sidebar     │
│  pointer-events: none            │  (321px)     │
│                                  │              │
│  During capture: live render     │  Logo        │
│  (scrolling page)                │  Progress    │
│                                  │  "3 of 5"    │
│  After capture: scrollable       │  Cancel btn  │
│  preview image                   │              │
│                                  │  Save btn    │
│                                  │              │
└──────────────────────────────────┴──────────────┘
```

### What V1 Includes
- Branded sidebar with OneNote logo (`onenote_logo_clipper.png`)
- (Originally OneNote purple theme; superseded by Fluent 2 white theme — see "Fluent 2 redesign" below)
- Localized progress text ("Capturing {0} of {1}...") via `fullPageStrings` in session storage
- Sidebar pixel cropping: `captureVisibleTab` captures full window, canvas only draws left content area
- Preview phase: iframe hides, scrollable image appears with Save/Close buttons
- Cancel closes window (port disconnect triggers cleanup)
- All i18n keys from existing `strings.json` + two new keys (`IncrementalProgress`, `Saving`)

### Files Modified (V1)
| File | Change |
|------|--------|
| `src/renderer.html` | Flexbox layout: content iframe + sidebar div, references `renderer.css` and logo |
| `src/styles/renderer.less` | NEW — LESS styles using OneNote brand colors |
| `src/scripts/renderer.ts` | Sidebar progress, capture cropping, preview phase, i18n string application |
| `src/scripts/contentCapture/fullPageScreenshotHelper.ts` | Passes `fullPageStrings` i18n map via session storage |
| `src/scripts/extensions/webExtensionBase/webExtensionWorker.ts` | Window width = content + sidebar, passes `totalViewports` |
| `src/strings.json` | New keys: `ScreenShot.IncrementalProgress`, `ScreenShot.Saving` |
| `gulpfile.js` | Compiles `renderer.less` alongside `clipper.less` |

---

## V2/V3 — Implemented: Full Clipper UI in Unified Window

### Design Principle
- **V2**: Post-auth UI (modes, save, region, section picker) in the renderer window
- **V3**: Self-contained sign-in — sign-in also moved into the renderer; no clipperInject.ts injection at all
- No Mithril.js — plain HTML/TS sidebar with port messages to the worker
- **Self-contained extraction** — Readability and bookmark metadata extracted directly in renderer from content-frame DOM, no round-trip to worker/clipper.tsx

### Flow (Current)

**No clipperInject.ts injection. No Mithril sidebar. Everything in the renderer window.**

```
┌─────────────────────────────────────────────────────────────────────┐
│ 1. USER CLICKS EXTENSION BUTTON                                     │
│                                                                     │
│    Worker.closeAllFramesAndInvokeClipper()                          │
│      ├─ activeRendererWindowId set? → focus existing window, return │
│      └─ openRendererWindow()                                        │
│           ├─ clipperData.getValue("isUserLoggedIn")                 │
│           └─ launchRenderer(signedIn: boolean)                      │
└─────────────────────────────────────────────────────────────────────┘
                            │
              ┌─────────────┴─────────────┐
              ▼                           ▼
┌──────────────────────┐    ┌──────────────────────────────────┐
│ NOT SIGNED IN        │    │ SIGNED IN                        │
│                      │    │                                  │
│ Opens renderer window│    │ Opens renderer window            │
│ (no content capture) │    │ + injects contentCaptureInject.js│
│                      │    │   into original tab (parallel)   │
│ Renderer detects no  │    │                                  │
│ userInformation in   │    │ contentCaptureInject cleans DOM, │
│ localStorage         │    │ sends HTML, title, URL           │
│   → shows sign-in    │    │   → sendMessage to worker        │
│     overlay          │    │   → worker stores in             │
│                      │    │     chrome.storage.session        │
└──────┬───────────────┘    └──────────────┬───────────────────┘
       │                                   │
       ▼                                   ▼
┌──────────────────────┐    ┌──────────────────────────────────┐
│ 2. SIGN-IN FLOW      │    │ 3. CAPTURE FLOW                  │
│                      │    │                                  │
│ User clicks MSA or   │    │ Renderer port "ready"            │
│ OrgId button         │    │   → worker sends "loadContent"   │
│   → port: signIn     │    │     (after content captured)     │
│   → worker opens     │    │                                  │
│     OAuth popup via   │    │ Renderer loads HTML into iframe, │
│     windows.create    │    │ processes CSS, sends "dimensions"│
│   → webRequest detects│    │                                  │
│     redirect          │    │ Worker scroll-capture loop:      │
│   → updateUserInfoData│    │   scroll → captureVisibleTab     │
│   → port: signInResult│    │   → drawCapture → drawComplete   │
│                      │    │   → repeat until atBottom         │
│ Renderer hides overlay│    │   → finalize                     │
│ shows sidebar + iframe│    │                                  │
│ injects content       │    │ Renderer stitches JPEG on canvas,│
│ capture script        │    │ stores in session storage,       │
│   → proceeds to ──────┼───►│ enables mode buttons + sign-out  │
│     capture flow      │    │                                  │
└──────────────────────┘    └──────────────┬───────────────────┘
                                           │
                                           ▼
┌─────────────────────────────────────────────────────────────────────┐
│ 4. MODE SELECTION                                                   │
│                                                                     │
│ Full Page: preview-container shows stitched JPEG (cached)           │
│ Article:  Readability on content-frame DOM clone → preview-frame    │
│ Bookmark: og:image/description from DOM → card in preview-frame     │
│ Region:   overlay injected on original tab → crosshair selection    │
│           → captureVisibleTab → canvas crop → multi-region support  │
└──────────────────────────────┬──────────────────────────────────────┘
                               │
                               ▼
┌─────────────────────────────────────────────────────────────────────┐
│ 5. SAVE (CLIP)                                                      │
│                                                                     │
│ lockSidebar() — all controls + sign-out disabled                    │
│ Renderer → port: { action: "save", title, annotation, url, mode,   │
│                    sectionId, contentHtml }                         │
│ Worker reads accessToken from localStorage via offscreen            │
│ Worker builds multipart form (ONML + image MIME parts)              │
│ Worker POSTs to https://www.onenote.com/api/v1.0/.../pages         │
│ Worker → port: { action: "saveResult", success, pageUrl/error }     │
│                                                                     │
│ Success: Clip → "View in OneNote" link, Close stays                 │
│ Error: error message + expandable diagnostics (correlationId, date) │
└──────────────────────────────┬──────────────────────────────────────┘
                               │
                               ▼
┌─────────────────────────────────────────────────────────────────────┐
│ 6. SIGN-OUT                                                         │
│                                                                     │
│ Renderer → port: { action: "signOut", authType }                    │
│ Worker: doSignOutAction (hits sign-out URL)                         │
│ Worker: clears userInformation, notebooks, curSection, isUserLoggedIn│
│ Worker → port: { action: "signOutComplete" }                        │
│ Renderer: full UI reset (clears DOM, caches, notebook dropdown)     │
│ Renderer: shows sign-in overlay (stays in same window)              │
│                                                                     │
│ User can sign in again as same or different user without closing    │
└─────────────────────────────────────────────────────────────────────┘

WINDOW LIFECYCLE:
  • Duplicate prevention: activeRendererWindowId check
  • Tab navigation: tabs.onUpdated closes renderer when source tab URL changes
  • Focus retention: blur handler re-focuses (disabled during sign-in + region)
  • Resize lock: resize handler + chrome.windows API reverts maximize
  • UI lock: sign-out + controls disabled during capture and save
  • Cleanup: port disconnect removes window + session storage keys

MESSAGE PROTOCOL (port — all renderer sends via safeSend() wrapper):
  Renderer → Worker: ready, dimensions, scrollResult, drawComplete,
                     finalizeComplete, save, startRegion, signIn, signOut,
                     telemetry, openFeedback
  Worker → Renderer: loadContent, scroll, initCanvas, drawCapture,
                     finalize, saveResult, regionCaptured, regionCancelled,
                     signInResult, signOutComplete

MESSAGE PROTOCOL (chrome.runtime.sendMessage — JSON strings):
  contentCaptureInject → Worker: contentCaptureComplete
  regionOverlay → Worker: regionSelected, regionCancelled
```

### Layout
```
Window width: 1280 (content) + 321 (sidebar) = ~1601px
┌──────────────────────────────────┬──────────────────────┐
│                                  │ Logo  OneNote Clipper │
│  content-frame (capture)         │                      │
│  OR preview-frame (article/bkmk) │ [Full Page] selected │
│  OR preview-container (JPEG)     │ [Article]            │
│  OR region thumbnails + add btn  │ [Bookmark]           │
│                                  │ [Region]             │
│  pointer-events: none (capture)  │ ─────────────────    │
│  scrollbar visible (preview)     │ TITLE  [textarea]    │
│                                  │ NOTE   [textarea]    │
│                                  │ SOURCE url           │
│                                  │ SAVE TO [dropdown]   │
│                                  │                      │
│                                  │ [Cancel] [Clip]      │
│                                  │ progress/status/error│
│                                  │                      │
│                                  │ user@email | Sign out│
└──────────────────────────────────┴──────────────────────┘
```

### Three Content Panels (Left Side)
- **content-frame** (`<iframe>`): Loads page HTML for capture. DOM also used as source for Readability and bookmark extraction. Visible during capture, hidden after.
- **preview-frame** (`<iframe>`): Article and bookmark HTML rendered via `document.write`. Hidden during full-page/region mode.
- **preview-container** (`<div>`): Scrollable JPEG preview image after capture (full-page), OR region thumbnails with ×remove buttons and "+Add another region" button. Scrollbar visible. Hidden during article/bookmark modes.

### Sidebar Implementation
Plain HTML + TypeScript — no Mithril dependency:
- **Mode buttons**: 5 buttons with SVG icons in order **Full Page → Region → Bookmark → Article → PDF**. PDF hidden by default, shown only when content type is PDF. Full Page pre-selected. All modes disabled during capture, enabled after.
- **Title**: `<textarea>` pre-filled from page title (session storage), editable
- **Note**: `<textarea>` placeholder "Add a note..."
- **Source URL**: read-only display from session storage
- **Section picker**: `<select>` dropdown populated from `localStorage.notebooks` cache + auto-refreshed from OneNote API. Flattens notebook > section hierarchy including section groups. Selection persisted to `localStorage.curSection`. Token expiration check uses relative `accessTokenExpiration` with `lastUpdated` offset.
- **Button row**: Cancel + Clip side by side (flex row). Clip disabled during capture and save.
- **Status panel**: Below buttons. Shows capture progress during screenshot, "Saving..." during API call, error + expandable diagnostics on failure.
- **User info footer**: Pinned to bottom of sidebar (outside sidebar-body). Shows user email/name + "Sign out" link. Hidden if no user info available. Disabled during save lock.

### Region Capture
Standalone `regionOverlay.ts` injected directly into original tab via `scripting.executeScript`. No Mithril, no clipper.tsx reactivation.
- **Overlay**: Full-viewport div with canvas-drawn crosshair (cursor:none, drawn for both mouse and keyboard). Dark overlay with hole-punch selection.
- **Instruction bar**: Shadow DOM pill with i18n instruction text + "Back (Esc)" button. CSS-isolated from page styles. Hides during drag, reappears on too-small selection. i18n strings passed from renderer → worker → `window.__regionStrings` injection before overlay script.
- **Mouse selection**: Drag draws rectangle. Min 5px. Esc/Back cancels (stays in region mode, shows thumbnails or "Add another region").
- **Keyboard selection**: Arrow keys with velocity acceleration (1→5), two-phase Enter (start → complete), ARIA live announcements. Seamless keyboard→mouse handoff (preserves start point).
- **Message format**: JSON string via `chrome.runtime.sendMessage` (required by offscreen.ts message handler)
- **Capture**: Worker captures original tab as JPEG 95% via `captureVisibleTab`, sends full image + coords via port
- **Crop**: Renderer crops using canvas with DPR handling, converts to JPEG 95%
- **Multi-region**: `regionImages[]` page-variable array accumulates captures (no session storage). Thumbnails at natural size (no pixelation, per designer spec) with **Discard pill above each image** (translucent black 5% bg, trash icon + label). 24px gap between captures. Cached across mode switches within session.
- **Save**: Worker accumulates per-port `saveImage` chunks (one per region) and builds ONML with one `<img>` per region.
- **Focus**: Renderer blur handler skips re-focus when `currentMode === "region"`. Worker focuses original tab's window via `tabs.get` + `windows.update`.

### Sign-Out Flow
Message-based, stays in renderer window (V3):
1. Renderer: `safeSend({ action: "signOut", authType })`
2. Worker: `doSignOutAction(authType)` hits sign-out URL, `clipperData.removeKey()` for userInformation/curSection/notebooks/isUserLoggedIn
3. Worker → port: `{ action: "signOutComplete" }`
4. Renderer: full UI reset (clears DOM, caches, notebook dropdown)
5. Renderer: shows sign-in overlay (stays in same window, user can sign in again)

### Section Refresh
Auto-fetch on renderer open (runs in background during capture):
1. Reads access token from `localStorage.userInformation` (TimeStampedData format)
2. Checks token expiration: `(lastUpdated + accessTokenExpiration * 1000 - 180000) < Date.now()`
3. Fetches `https://www.onenote.com/api/v1.0/me/notes/notebooks?$expand=sections,sectionGroups($expand=sections,sectionGroups($expand=sections,sectionGroups($expand=sections,sectionGroups)))` — 4 levels deep, matches legacy `onenotepicker` `maxExpandedSections = 4`
4. Compares with cached — only updates dropdown if data changed
5. Preserves selected section if still exists in fresh data
6. Silently keeps cached data on failure

**Hierarchical picker UI**: Notebook and section-group headings rendered as `role="button"` collapsible toggles with `arrow_down`/`arrow_right` icons. Depth-aware collapse logic walks `nextElementSibling` and breaks at headings of same/shallower depth (`data-depth` attribute). Section items rendered as `role="option"` with depth-based padding-left.

### Save Flow (Worker Side)
Save uses chunked port streaming — no session storage for image data:
1. Renderer sends `{ action: "save", saveImageCount: N, saveAttachment: bool, ...metadata }` (title, annotation, url, mode, sectionId, contentHtml, etc.)
2. Renderer streams `{ action: "saveImage", index, dataUrl }` messages (one per image)
3. For PDF mode: optional `{ action: "saveAttachment", dataUrl }` message with binary PDF bytes
4. Worker accumulates chunks in per-port variables (`pendingSave`, `saveImages[]`, `saveAttachmentData`); `executeSave()` runs when all expected chunks received
5. Worker reads access token from `clipperData.getValue("userInformation")` → `data.accessToken`
6. Worker builds OneNote page ONML:
   - Full Page: single image MIME part
   - Region: multiple image MIME parts (one per region)
   - Article/Bookmark: HTML from `contentHtml` message field
   - PDF (non-distributed): all selected pages as MIME parts + optional `<object>` attachment
   - PDF (distributed): sequential POSTs, one OneNote page per PDF page
7. Generates UTC offset timestamp matching `OneNotePage.formUtcOffsetString`
8. POSTs multipart form to `https://www.onenote.com/api/v1.0/me/notes/sections/{sectionId}/pages`
9. On success: parses `links.oneNoteWebUrl.href` from response, sends `saveResult` with `pageUrl`
10. On error: captures error message, status, X-CorrelationId, X-UserSessionId, date; sends as `saveResult`

**Per-port isolation**: Each renderer port has its own accumulator, so multiple renderer windows can save concurrently without interfering. Bypasses the 10MB `chrome.storage.session` quota that was a bottleneck for large captures.

### i18n
Renderer reads `localStorage.locStrings` directly (shared extension origin). Falls back to hardcoded English. No session storage dependency for strings (legacy `fullPageStrings` passthrough still exists in fullPageScreenshotHelper but unused).

**New i18n keys** added by the Fluent 2 redesign (English fallback used until added to translation pipeline):
- `WebClipper.Label.WhatToCapture` — "What do you want to capture?" caption above mode buttons
- `WebClipper.Action.Discard` — "Discard" label on region thumbnail remove pill
- `WebClipper.Label.ClipSuccessTitle` — "Saved to your notebook"
- `WebClipper.Label.ClipSuccessDescription` — "You can access and continue working on it anytime from your notebook."
- `WebClipper.Label.ClipErrorTitle` — "Couldn't save to your notebook"
- `WebClipper.Label.ClipErrorDescription` — "Something went wrong while saving your clip. Try saving your clip again"

### Window Lifecycle
- **Duplicate prevention**: `closeAllFramesAndInvokeClipper` override checks `activeRendererWindowId`
- **Tab navigation**: `tabs.onUpdated` listener closes renderer when source tab URL changes
- **Focus retention**: `blur` handler re-focuses renderer (disabled after save and during region capture)
- **Resize**: locked during capture (2px macOS tolerance, `resizing` guard flag), unlocked after via `showPreviewFrame()`. Post-capture min 1000x600 enforced via `chrome.windows.update`. Worker creation: content width capped at min(browserWidth, 1280), total clamped to max(browserWidth, 1000), height clamped to browserHeight − 32px.
- **Service worker keepalive**: 25s ping from renderer prevents MV3 SW suspension
- **Inactivity auto-close**: 5-min timer, reset on user input
- **Save timeout**: 30s client-side timeout (SW `setTimeout` unreliable); scales with PDF page count (30s + 5s/page)
- **Cleanup**: port disconnect (Cancel click or window close) triggers worker cleanup

### Files Modified (V2/V3)
| File | Change |
|------|--------|
| `src/renderer.html` | Two iframes + sidebar with mode buttons, metadata fields, button row, status panel, user-info footer |
| `src/styles/renderer.less` | Full sidebar styles, user-info footer, region thumbnails/buttons, preview scrollbar |
| `src/scripts/renderer.ts` | Mode switching, Readability, bookmark extraction, section picker + auto-refresh, region capture + multi-region, sign-out, save flow, UI lock, i18n |
| `src/scripts/extensions/webExtensionBase/webExtensionWorker.ts` | Save handler (fullpage/region/article/bookmark), sign-out handler, region overlay injection + message listener, duplicate window guard, tab nav listener |
| `src/scripts/extensions/regionOverlay.ts` | NEW — Standalone crosshair overlay for region selection on original tab |
| `src/scripts/contentCapture/fullPageScreenshotHelper.ts` | Passes pageTitle, legacy i18n strings |
| `src/scripts/clipperUI/clipper.tsx` | hideUi for signed-in users, post-sign-in transition, showSignInPanel handler, non-signed-in guard |
| `src/scripts/extensions/clipperInject.ts` | showUi handler (mirror of hideUi) |
| `src/scripts/constants.ts` | New keys: showUi, showSignInPanel, startRegionCapture, regionCaptureComplete, regionCaptureCancelled |
| `gulpfile.js` | Compiles + deploys regionOverlay.js to Chrome and Edge targets |

---

## Verification Checklist

### V1 (Done)
- [x] Window opens at correct width (content + sidebar)
- [x] Sidebar shows OneNote branding and progress
- [x] Captures exclude sidebar pixels
- [x] Preview appears after capture
- [x] Cancel/Close works
- [x] All strings localized
- [x] LESS styles compiled and deployed

### V2/V3 (Done)
- [x] Mode buttons switch left panel content (Full Page, Region, Bookmark, Article, PDF)
- [x] Article mode: Readability extracts content directly from content-frame DOM
- [x] Bookmark mode: metadata card with og:image/description from DOM (thumbnail base64-converted via canvas — OneNote API can't fetch external URLs)
- [x] Title editable, Note field, Source URL display
- [x] Custom section picker (ul/li dropdown) with scrollbar + auto-refresh from API
- [x] Clip saves to OneNote via direct fetch to API (multipart form)
- [x] Window stays open after capture for mode switching
- [x] Save includes annotation, "Clipped from" citation, timestamp
- [x] Post-save "View in OneNote" button with page URL from API response
- [x] Error display with expandable diagnostic info (correlation ID, date, status)
- [x] i18n from localStorage (locStrings) — no session storage dependency
- [x] Duplicate window prevention (focus existing on re-click)
- [x] Tab navigation detection closes renderer window
- [x] Modal-like focus (blur handler re-focuses renderer, disabled during sign-in + region)
- [x] UI lock during capture and save (all inputs + sign-out disabled)
- [x] Close + Clip side-by-side button row
- [x] Region capture: canvas-drawn crosshair overlay, overflow:hidden on root
- [x] Multi-region: thumbnails, ×remove, +add another, cached across mode switches
- [x] Region save: multiple images as separate MIME parts in ONML
- [x] Self-contained sign-in: MSA/OrgId overlay, OAuth popup, no clipperInject injection
- [x] Sign-out from renderer → shows sign-in overlay (stays in same window)
- [x] Sign-out cleanup: clears notebooks, section cache, stale capture data
- [x] Content capture via standalone contentCaptureInject.ts (master DomUtils pipeline + enhancements, no imports, no iframe injection)
- [x] Position neutralization with `!important` prevents sticky element duplication in stitched captures
- [x] Stylesheet caching removed — renderer fetches CSS directly via `<link>` tags
- [x] `[hidden]{display:none!important}` CSS override enforces HTML hidden attribute
- [x] `safeSend()` wrapper on all port.postMessage calls (handles disconnected port from devtools)
- [x] User info footer pinned to sidebar bottom
- [x] Preview container + sidebar-body scrollbars visible
- [x] Full-page preview restored from cached data URL when switching modes
- [x] Anti-maximize: chrome.windows API reverts maximized state
- [x] Article/bookmark preview: links disabled (pointer-events: none), consistent scrollbar
- [x] Feedback link in footer: smiley icon + i18n label, hidden for MSA users, opens MS Feedback Portal popup with context params (LogCategory, originalUrl, clipperId, usid, version, type). Per-session USID (cccccccc- prefix) sent as X-UserSessionId header in save API calls for log correlation.
- [x] Error diagnostics copy button: clipboard icon next to "More information", copies error text via navigator.clipboard.writeText
- [x] Article preview header bar: highlighter toggle, serif/sans-serif font toggle, font size +/- (2px increments, 8–72px). Purple toolbar above preview-frame, visible only in article mode. TextHighlighter library injected into iframe for drag-to-highlight with yellow (#fefe56) spans. Delete buttons (red circle ×, top-left) on each highlight group. Font/size applied to saved OneNote content via wrapping `<div style>`. Highlight state preserved across mode switches via `articleWorkingHtml`.
- [x] Save timeout: 30-second client-side timeout in renderer (service worker setTimeout unreliable — SW goes inactive). Shows "Request timed out (30s)" with retry-enabled Clip button.
- [x] Version bump to 3.11.0, gulp-uglify v2→v3 for ES6+ minification, `<meta charset="utf-8">` in renderer.html
- [x] Full-page image width fix: renderer sends actual CSS width (contentPixelWidth/DPR) to worker for ONML `<img width>`, fixing aspect ratio distortion
- [x] Capture progress: removed "Clipping Page" heading (reserved for save phase), progress bar hidden until first viewport capture
- [x] Mode buttons reverted from role="radio" to role="toolbar" + aria-pressed (more natural toggle button pattern)
- [x] NVDA screen reader verified working with Edge after devbox reboot
- [x] Consistent × button positioning: region thumbnails and highlight delete buttons both use top-left corner
- [x] Success banner: separate from Clip button — "✓ Clip Successful!" banner with purple "View in OneNote" button (opens page + closes window). Clip button stays as Clip for re-clipping. Banner clears on mode switch, re-clip, or section change.
- [x] Hierarchical section picker: notebook headings (with notebook.png icon) and section group headings (with section_group.png icon) as non-clickable groupings. Sections indented under their parent with section.png icon. Icons inverted to white for dark background contrast. Keyboard nav skips headings.
- [x] Section selection decoupled from button state: changing section only clears success banner, does not touch Clip button disabled/text state (was causing premature Clip enable during capture).
- [x] Lint conformance: fixed all tslint errors in renderer.ts, regionOverlay.ts, webExtensionWorker.ts, domUtils.ts (null→undefined, var→let, quotemark, else placement, shadowed variable). Only pre-existing third-party definition file errors remain.
- [x] Service worker keepalive: renderer sends keepalive ping every 25s via port to prevent MV3 service worker suspension while popup is open.
- [x] Inactivity auto-close: renderer closes after 5 minutes without mouse, keyboard, scroll, or focus activity. Timer resets on any interaction.

---

## Known Limitations

### Capture Quality
1. **Viewport width vs responsive breakpoints** — Content-frame renders at 1280px. Most sites (including MS Learn at 1088px breakpoint) render correctly, but sites with breakpoints above 1280px may lose sidebars or switch to mobile layout. CSS `zoom` doesn't affect media queries. Widening the cap is possible but creates wider images. A future approach could capture at the original tab viewport width and scale on save.
2. ~~Sticky element duplication~~ — Resolved. Position neutralization (`sticky → relative`, `fixed → absolute`) now uses `!important` via `setProperty()` in both `contentCaptureInject.ts` and `renderer.ts`, beating CSS utility classes like `.position-sticky{position:sticky!important}`. Hidden duplicate elements (e.g., MS Learn collapsible TOC) handled by `inlineHiddenElements` + `[hidden]{display:none!important}` CSS override.
3. ~~Bottom void on grid-layout sites~~ — Resolved. Canvas is trimmed to `stitchYOffset` (actual pixels drawn) during finalize, removing any trailing blank space.
4. **Right-edge clipping** — Pages with 0 margins may get content cut off at the right edge.
5. **Canvas height cap** — Maximum stitched canvas height is 16384px (browser limitation).
6. **Video/streaming embeds** — YouTube and other iframe-based video embeds show broken players in the content iframe (cross-origin restrictions, no JS execution). Server-side Puppeteer had the same limitation (showed "Unable to execute Javascript"). Left as-is to avoid accidentally stripping legitimate iframe content.
7. **Client-side rendered / Shadow DOM sites** — Sites that deliver a blank JavaScript shell and render all content client-side via web components (e.g., MSN.com rebuilt on Microsoft FAST framework) produce empty captures. `document.cloneNode(true)` cannot copy shadow root content per the DOM spec, so custom element shells clone as empty boxes. This is a **pre-existing limitation** — server-side Puppeteer also ran with `--disable-javascript` and would see the same blank shell. Affects both full-page screenshot and article extraction (Readability finds no content). Potential future mitigations: `Element.getHTML({ serializableShadowRoots: true })` API, manual shadow root traversal, or `document.body.innerText` fallback for article mode.

### ~~Missing Features~~ Implemented
6. ~~**Self-contained sign-in**~~ — Implemented. Worker opens renderer directly on button click (no clipperInject.ts injection). Renderer checks `localStorage.userInformation` — shows MSA/OrgId sign-in overlay when not signed in. Sign-in via port messages, OAuth popup via `chrome.windows.create`, `auth.updateUserInfoData` on redirect. Content capture via standalone `contentCaptureInject.ts` (injected by worker via `scripting.executeScript`). Sign-out stays in renderer (shows sign-in overlay, clears storage). Custom section picker with scrollable `<ul>/<li>` dropdown. UI locked during capture to prevent race conditions. Region overlay selection border drawn on canvas (no separate div). Old injected sidebar (clipperInject.ts → clipper.tsx → Mithril) is now dead code.

### UI Polish (Deferred)
7. **Toolbar layout redesign** — Current sidebar takes 322px width. A top-bar + bottom-bar layout would eliminate the sidebar width tax, making the window narrower and giving content full width.

### Technical Debt
8. ~~fullPageStrings in session storage~~ — Resolved. Removed legacy i18n passthrough from fullPageScreenshotHelper.ts. Renderer reads from localStorage directly.
9. ~~fullPageScreenshotHelper promise~~ — Resolved. Promise now resolves on `finalizeComplete` with `{ success: true, format: "jpeg", cssWidth }`. Session storage cleanup runs correctly; `fullPageFinalImage` kept for save flow, cleaned up on window close.
10. ~~Save URL from session storage~~ — Resolved. URL passed via port message (`message.url`) instead of session storage. All save parameters come from port messages.
10a. ~~Image data in session storage~~ — Resolved. Images now flow via chunked port streaming (`saveImage` chunks per port). Session storage holds only HTML metadata (`fullPage*` keys for content HTML/title/URL/contentType).
11. **Readability bundled in renderer** — Dynamic import pattern added but browserify converts to synchronous require (no code-splitting). True lazy-loading requires bundler upgrade (webpack, rollup, esbuild).
12. **ES target compatibility** — Project targets old ES version. `String.startsWith()` not available, must use `indexOf`. Silent build failures if modern APIs are used.
13. **chrome.runtime.sendMessage format** — offscreen.ts does `JSON.parse(message)` on ALL incoming messages. New scripts sending via `chrome.runtime.sendMessage` must use `JSON.stringify(...)` not plain objects.

### Telemetry
14. **Renderer telemetry** — Renderer sends funnel events via port `{ action: "telemetry", data: LogDataPackage }`. Worker routes to `this.logger` via `Log.parseAndLogDataPackage()`. Uses imported enums: `Funnel.Label`, `LogMethods`, `Session.EndTrigger`.
15. **Console output requires flag** — `LogManager.createExtLogger()` returns a `StubSessionLogger` (no-op) unless `enable_console_logging` is `"true"` in localStorage. To enable: open any extension page console (renderer, offscreen) and run `localStorage.setItem("enable_console_logging", "true")`, then reload extension. Logs appear in the **service worker console** (edge://extensions → Inspect service worker), not in the renderer or page console.
16. **No external telemetry endpoint** — All logging goes to `console.log()` via `ConsoleLoggerDecorator → WebConsole`. No HTTP POST, no Application Insights, no Aria SDK. The decorator pattern allows plugging in a real backend later.

---

## PDF Mode (Implemented)

PDF detected by `contentCaptureInject.ts` (URL `.pdf` suffix, `<embed type="application/pdf">`, or `window.PDFJS`); contentType reported as `"pdf"` via the loadContent message. Renderer enters `pdfMode` and uses `PDFJS.getDocument()` from `pdf.combined.js` (pdfjs-dist v1.7.290) loaded as `<script>` in renderer.html.

- **Mode buttons in PDF mode**: PDF + Region + Bookmark visible (Full Page / Article hidden — not applicable to PDF content).
- **Preview**: Pages rendered to canvas at scale=2 → data URL → `<img>` lazy-loaded into `preview-container`. Initial 3 pages rendered, ±1 on scroll. Page number overlay top-left of each image. Unselected pages get `opacity: 0.3`.
- **Page selection**: Radio group (All pages / Page range with validation). Range parser mirrors legacy `StringUtils.parsePageRange` — supports `1-5, 7, 9-12` syntax. Range input is `readonly` when in "All pages" mode (visible muted style); clicking the input auto-switches to "Page range" mode.
- **Attach PDF**: Checkbox enabled if file ≤24.9MB (`Constants.Settings.maximumMimeSizeLimit`). Auto-disabled with warning when too large.
- **Distribute pages**: Checkbox creates one OneNote page per PDF page (sequential POSTs). First page gets annotation + citation + optional attachment. Title uses actual page numbers (e.g., "Doc.pdf: Page 3").
- **Save**: Chunked port streaming — selected pages rendered, sent via `saveImage` chunks; optional `saveAttachment` with binary PDF bytes (base64-encoded data URL).
- **Bookmark fallback**: PDF pages have no og: tags in their viewer DOM, so `generatePdfBookmarkHtml(title, url)` builds a bookmark card from title + URL.
- **Telemetry**: `ClipPdfOptions` (PdfAllPagesClipped, PdfAttachmentClipped, PdfIsLocalFile, PdfIsBatched, PdfFileSelectedPageCount, PdfFileTotalPageCount), `PdfByteMetadata` (ByteLength, BytesPerPdfPage), `ClipCommonOptions` (ClipMode=Pdf).

---

## Fluent 2 Redesign (Implemented)

Per designer Figma spec (`wK4ryPaULoiSDMt2xVc1fH`). Replaces the original dark purple sidebar with a Fluent 2 white theme using OneNote brand accents.

### Tokens & palette
- **Neutral**: `colorNeutralBackground1` (#fff surface), `colorNeutralForeground1` (#242424 primary text), `colorNeutralForeground2` (#424242 labels), `colorNeutralForeground3` (#707070 tertiary), `colorNeutralStroke1` (#d1d1d1 input/button border), `colorNeutralStroke2` (#e0e0e0 divider).
- **OneNote brand ramp** (from Fluent 2 library `Colors/Brand/OneNote/*`): `oneNoteBrand50` (#5B1382 pressed), `oneNoteBrand60` (#6C179A hover), `oneNoteBrand80` (#7719AA rest). Mapped to `colorCompoundBrandBackground/Hover/Pressed`. Used on Clip button, focus rings, selected mode states.
- **Spacing**: All paddings aligned to multiples of 4 per Fluent spacing scale (XXL=24, L=16, M=12, S=8, XS=4).

### Component patterns
- **Mode buttons**: Fluent outline (white bg, 1px neutral border, icon + left-aligned text). Selected = `colorSubtleBackgroundHover` bg + brand-tinted border + brand text + brand-tinted icon (via SVG `currentColor` + CSS filter on `.selected`).
- **Inputs / textarea / dropdown**: Fluent outline (1px stroke). Focus = 1px purple inset shadow + border-color purple (single 2px purple edge, no doubled border).
- **Clip button**: Filled primary brand. Focus uses Fluent 2 "halo" pattern (`box-shadow: inset 0 0 0 2px #fff, 0 0 0 2px #242424`) — white inner ring against purple fill, dark outer ring against white sidebar bg.
- **Cancel button**: Fluent outline secondary.
- **Section picker**: White Fluent dropdown with chevron SVG. Section list popup with shadow. Headings as collapsible toggles with `arrow_down`/`arrow_right` icons.
- **Sign-in panel**: White Fluent theme matching sidebar. `aria-modal="true"` on dialog. Brand primary + outline secondary buttons. Focused element gets `colorCompoundBrandBackground` ring.
- **Article toolbar**: White Fluent (was purple). Highlighter, font toggle, font size buttons use neutral `currentColor` icons; brand-tinted active state.
- **Region thumbnails**: Natural size (no pixelation per designer spec). 24px gap between captures. **Discard pill above each image** (translucent black 5% bg, trash icon + "Discard" label).
- **Success/error banners**: Fluent inline alerts (icon + title + description + outline secondary CTA). Success: `colorStatusSuccessForeground1` (#0e700e) checkmark. Error: `colorStatusDangerForeground1` (#c50f1f) circle. No filled green/red banners.
- **Preview rounded frame**: 20px border-radius (matches Figma `rounded-[20px]`) overlay drawn after capture via `.preview-ready` class. Removed before subsequent captures so the rounded margin doesn't bake into screenshots.

### Typography
- Header title: Segoe UI Semibold, 20px / line 26px (Subtitle 1)
- Field labels: Segoe UI Semibold, 12px / line 16px (Caption 1 Strong)
- Mode button text: Segoe UI Semibold, 14px / line 20px (Body 1 Strong)
- Clip / Cancel: Segoe UI Semibold, 16px / line 22px (Subtitle 2)
- Body / placeholder text: Segoe UI Regular, 14px / line 20px (Body 1)
- Footer email/signout: Segoe UI Regular, 12px / line 16px (Caption 1)

### Header & footer
- Header: full clipper logo (N + scissors, `onenote_logo_clipper.png`) + "OneNote Web Clipper" title.
- Footer: avatar/email/signout on **left**, feedback link (or thumb up/down for OrgId) on **right**.

## Reviewer Feedback Round 3 (Implemented)

### Keyboard / focus
- **Clip button focus indicator after success**: explicit `saveBtn.focus()` after the success banner shows. While save is in flight the button is `disabled`, which drops focus and the `:focus-visible` promotion. A fresh programmatic focus brings both back.
- **Region Esc with no captures snaps back**: cancelling a fresh region selection returns to the page's default mode (Full Page on web, **PDF Document on PDF pages**) and leaves keyboard focus on the Region mode button so the user can re-enter region mode immediately.
- **Region overlay Back button reachable**: shadow root mode flipped from `closed` → `open`, `:focus-visible` style added, document-level keydown handler now traps Tab inside the overlay (cycles between root and back-btn), and lets native Enter/Space activate the button when it has focus (would otherwise be eaten by the region-selection start handler).
- **Section list headings reachable via arrow keys**: shared `focusAdjacentSectionRow` helper walks all visible rows (headings + items, skipping `display:none` collapsed children). Headings activate via Enter/Space (toggle); items via Enter/Space (select).
- **Preview iframe wrap**: `<iframe id="preview-frame">` is now wrapped in `<div id="preview-frame-wrap" tabindex="0">`. The iframe itself is `tabindex=-1`, so Tab can never traverse into article links/buttons (legacy parity with `editorPreviewComponentBase.makeChildLinksNonTabbable`). Wrapper handles arrow / Page / Home / End keys via JS to scroll `previewFrame.contentWindow`.
- **Sign-in cancellation silent reset**: worker sends `{success:false, cancelled:true}` (no error string) when MSAL throws on user-cancel; renderer just resets the sign-in panel buttons rather than showing a banner. Matches legacy clipper behavior.

### Visual / Figma alignment
- **Sidebar width 320 + 1px border = 321px total** (Figma 242:4365). Was 322 with `box-sizing: border-box` (= 321 content). Off by 1px.
- **Sidebar header padding** `20px 24px 12px` → `24px 24px 12px` (Figma `pt-24`).
- **Sidebar body padding-top** `4px` → `0` (Figma flush against header divider).
- **Sidebar group padding** `12px` → `8px` above/below dividers (Figma `py-[8px]`).
- **Field label margins** `12/4` → `16/8` (16px between fields, 8px label-to-input gap; Figma 242:4396 `gap-[8px]` within block, `gap-[16px]` between blocks).
- **`#meta-group` flex gap** removed — relies on label margins now.
- **Mode buttons use inline SVG with `fill="currentColor"`** instead of `<img src=".svg">`. New Fluent UI System Icons (Page Fit, Tab Add, Bookmark, Form). Inline SVG inherits the button's text color, so it tints correctly in normal mode (black/purple) and forced-colors mode (ButtonText/HighlightText) without filter tricks.
- **Source URL link icon and feedback smiley → inline SVG** with `fill="currentColor"`. URL text moved to `<span id="source-url-text">` so the sibling icon survives `textContent` writes.
- **PDF mode icon (`pdf.png`)** kept as PNG (no Fluent SVG variant). On the white sidebar the original white-pixel PNG was invisible; `filter: brightness(0)` flips white → black for normal mode, and the legacy purple hue-rotate filter handles selected state.

### Forced-colors hygiene
- **Single sharp halo** instead of stacked blurred drop-shadows: `drop-shadow(0 0 0 CanvasText)` (was 2× `drop-shadow(0 0 1px CanvasText)` — accumulated blur softened icons).
- **Mode-button inline SVGs dropped from filter selector** entirely (currentColor handles theming natively); only `<img>`-based icons (PDF mode, section list, article header) still get the halo.

### Preview frame & PDF mode
- **Focus rings switched from `outline` with `outline-offset: -2px` to outer `box-shadow`** on `#preview-frame-wrap` and `#preview-container`. Inset outlines were getting painted *under* PDF page images (which have `position: relative` + `box-shadow` triggering compositor-layer promotion); outer shadow lives outside the box where children can't reach it.
- **`#preview-container` adds `overflow-x: hidden`** when `.preview-ready` so absolutely-positioned children (PDF page-number badges) get clipped to the rounded corner.
- **`#preview-frame-wrap` has `overflow: hidden`** so the iframe inside takes the wrap's rounded corners.
- **Code blocks in article preview wrap** (`white-space: pre-wrap; overflow-wrap: anywhere; !important`) instead of horizontal-scroll. OneNote doesn't preserve scrollbars in saved pages, so showing one in preview was misleading.
