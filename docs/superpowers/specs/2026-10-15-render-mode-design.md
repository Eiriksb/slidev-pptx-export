# Render mode: export what Slidev actually draws

Date: 2026-10-15 · Status: design approved in conversation, awaiting spec review

## Goal

An exported `.pptx` that looks like what Slidev shows — including a deck's own CSS, custom HTML
and Vue components — while staying editable (real text, shapes, tables and charts), without the
author mirroring their theme into `pptx.config.mjs` by hand. Elements PowerPoint cannot express
are allowed to become pictures; everything else is native.

Audience: any Slidev deck, proven on the Kartverket/innotek template deck
(`../innotek-SlideDevBase`), which is the main test case.

## Decisions made while designing

| Question | Decision |
|---|---|
| Who it is for | Any deck, proven on the innotek deck |
| How the browser is used | Opt-in `--render` flag; today's static export stays the default and the fallback |
| Unsupported elements | Native where possible, otherwise a picture of just that element |
| Approach | Render the deck in headless Chromium with its own Slidev and convert measured elements (not static CSS parsing, not a static/browser hybrid) |

## Background

Today the exporter only reads `slides.md`. Custom pieces reach the `.pptx` because
`pptx.config.mjs` restates them: brand colors, `cssVarColors`, `eyebrowClass`, a chart resolver
per Vue component, and a PNG copy of the SVG logo. Anything not restated becomes flattened text
or a placeholder box (the cover's gradient glow, image borders/shadows, any new component).

A throwaway probe against the innotek deck (Slidev 52.x dev server, Chromium) confirmed the two
assumptions this design rests on:

- Loading `/<n>?print` renders one slide on a 980×552 px canvas. Walking the slide's DOM with
  `getBoundingClientRect` and `getComputedStyle` gives every element's exact rectangle plus font
  family/weight/size, color, background, border, radius, letter spacing and text-transform
  (e.g. `.kv-card` → 8px radius, `#EEF4F0` fill, 1px `#D9E3DB` border; `.kv-eyebrow` → 800 12px
  Plus Jakarta Sans, uppercase, 1.68px tracking, `#1A833B`).
- The Chart.js chart behind `ExampleChart.vue` is reachable from its canvas through the owning
  Vue component (`__vueParentComponent` chain → exposed `chart`), giving type, labels, dataset
  values and colors, axis min/max, legend position and the tick callback (`v + "%"`). This relies
  on Vue's development build, which is what a Slidev dev server serves.

## Architecture

```
slides.md ──► src/render/server.mjs   starts the deck's own Slidev dev server
                     │
                     ▼
              src/render/capture.mjs  Playwright Chromium, one page load per slide
                     │  runs src/render/extract.js in the page
                     ▼
              Scene (JSON per slide)  measured nodes in canvas px
                     │
                     ▼
              src/render/scene.mjs    Scene → SlideSpec (inches, theme colors, placeholders)
                     │
                     ▼
              src/ooxml/*             existing writer, with the extensions listed below
```

`convert({ render: true })` / CLI `--render` takes this path; otherwise `convert()` behaves exactly
as today.

### `src/render/server.mjs`

- Resolves `@slidev/cli` from the deck directory (not from this package). Missing → error
  telling the user to run `npm install` in the deck folder.
- Spawns the CLI's bin with the deck's entry file, `--port <free port>`, `--bind 127.0.0.1`,
  stdio piped (no TTY prompts). Ready when an HTTP GET of `/` returns 200; startup timeout 60 s.
- Keeps the last 50 lines of server output for error messages. `stop()` kills the process tree;
  `convert()` always calls it in `finally`.

### `src/render/capture.mjs`

- Resolves Playwright at runtime: `playwright-chromium`, then `playwright`, looked up from the
  deck directory first, then from this package. Declared as optional peer dependencies
  (`peerDependenciesMeta`), so installing `slidev-pptx-export` never downloads a browser.
  Missing → error: install `playwright-chromium` (the package `slidev export` uses) and run
  `npx playwright install chromium`.
- Reads the canvas size from the page (`.slidev-slide-content` width/height; 980×552 by
  default). Viewport = canvas size, `deviceScaleFactor: 2` (for picture fallbacks).
- For each slide `n` (count from the page's slide list): load `/<n>?print`, wait for
  `document.fonts.ready`, network idle, and any Chart.js charts to finish their first render;
  per-slide timeout `render.timeout` (default 15 s). Captures the final click state (all
  `v-click` fragments shown). Phase 1's first task verifies the exact query/route for that and
  where the page exposes each slide's speaker notes.
- Runs `extract.js` in the page, then takes element screenshots for every node the extractor
  marked `raster`.
- Compatibility guard: if `.slidev-page`/`.slidev-slide-content` are not found, fail with
  "unsupported Slidev version" and the detected `@slidev/cli` version.

### Scene format (output of `extract.js`)

All geometry in canvas px relative to the slide's top-left. Nodes are in paint order (DOM order,
parents before children); z-order in the `.pptx` follows it.

```
Scene = {
  index, canvas: { w, h }, layout: "<slidev layout name>", notes: string | null,
  background: Fill | { raster: true },
  nodes: Node[]
}
Fill = { type: "solid", color, alpha }
     | { type: "linear", angle, stops: [{ pos, color, alpha }] }
     | { type: "radial", cx, cy, stops: [...] }
Node =
  { kind: "box", rect, fill?, border?: { width, color, alpha } | sides, radius, shadow?, opacity, rotate? }
  { kind: "text", rect, role: "title" | "body", align, lineHeight, paras: [{ runs: Run[] }], singleLine }
  { kind: "list", rect, ordered, start, markerColor, items: [{ level, indent, runs: Run[] }] }
  { kind: "image", rect, src, format: "png"|"jpg"|"gif"|"svg", fit, crop?, border?, shadow?, alt }
  { kind: "table", rect, colW: [], rowH: [], cells: [[{ runs, fill?, borders?, span? }]] }
  { kind: "code", rect, fill?, border?, radius, lines: [Run[]] }
  { kind: "chart", rect, chart: ChartSpec }
  { kind: "raster", rect, reason, png? }
Run = { text, font, size, bold, italic, color, alpha, spacing, underline, strike, highlight? }
```

### `src/render/scene.mjs`

Pure function `sceneToSpec(scene, style) → SlideSpec` (no browser; unit-tested with JSON
fixtures). Conversions:

- Scale `k = slideWidthIn / canvas.w` (13.333 / 980). Font size pt = px × k × 72
  (17.6 px ≈ 17.2 pt). Slide size stays `config.slide` (16:9, 13.333 × 7.5 in).
- Colors: `#RRGGBB` (+alpha). A color equal to a theme color maps to its theme slot via the
  existing `toColor`, so theme recoloring keeps working for brand colors.
- Layout: the deck's first slide and slides whose Slidev layout is `cover`/`intro` → *Title
  Slide*; all others → *Title Only*. The first `h1` (role `title`) fills the title placeholder
  with `titleBox` = its measured rect; there is no body placeholder in render mode.
- Notes from the scene (Slidev's own slide data in the page); if the page doesn't expose them,
  from our parser by slide index.

## Mapping rules

| Rendered element | `.pptx` result |
|---|---|
| Text block (`p`, `h1`–`h6`, `div` with own text, `blockquote`, …; table cells are handled by the table row below) | One text box at the content box (border-box minus padding), zero insets, top anchor. Inline descendants become runs (font = first family in the computed stack, weight ≥ 600 → bold, italic, size, color+alpha, letter spacing, underline/strike). `text-transform` is applied to the text itself. `<br>` → line break. `text-align` and computed `line-height` (as exact line spacing) are copied. A block rendered on one line gets `wrap="none"`; multi-line blocks wrap at the measured width. |
| Inline `code` | A run in its mono font with its background as run highlight. |
| First `h1` | Title placeholder at its measured rect. |
| `ul` / `ol` | One text box; each `li` a paragraph with its nesting level, native `•` bullet or auto-number (with `start`), bullet color from `::marker`, indent from the measured text start. Non-text content inside an `li` is emitted as its own nodes. |
| Element with background / border / box-shadow | A shape behind its contents: `rect`, or `roundRect` with `adj = radius / min(w, h) × 100000` (max 50000). Solid fill with alpha; uniform border → outline; single-side borders → thin bar shapes; `box-shadow` → `outerShdw` (blur, distance/direction from offsets, color+alpha). `opacity` multiplies the alpha of everything inside. A pure rotation → `xfrm rot`. |
| `img` PNG/JPG/GIF | Native picture, cropped (`srcRect`) to match `object-fit`, with border and shadow. |
| `img` SVG | PNG from the browser at 2× plus the SVG embedded via the `asvg:svgBlip` extension (vector in PowerPoint 2016+; PNG elsewhere). Phase 1 uses the PNG alone. |
| CSS `background-image: url()` | A picture behind the element's contents, cropped per `background-size`. |
| `table` | Native table: measured column widths and row heights, cell runs, each cell's computed fill and per-side borders written directly (built-in "No Style, No Grid" table style), `colspan`/`rowspan` → `gridSpan`/`rowSpan` + merges. |
| Shiki code block (`pre`) | Panel shape from the `pre`'s computed style + a text box: one paragraph per line, mono runs in token colors, leading spaces preserved. |
| Chart.js canvas | Native chart. Found via the owning Vue component's exposed `chart`. Bar charts map to the writer's column/bar charts (`indexAxis: "y"` → bar, `stacked`), with labels, values, per-series colors, axis min/max and legend position. Number format from calling the tick callback on sample values: output `prefix + value + suffix` → `"prefix"0"suffix"`, else General. Other chart types → picture until the writer supports them. A `chartResolvers` entry for the component wins over the captured chart. |
| Slide background (`.slidev-layout`) | Solid → slide background fill. Linear/radial gradient layers that reduce to one gradient (e.g. a radial glow over a solid color → a radial gradient from the blended color to the base color) → `gradFill` background. Anything else → a screenshot of the background (contents hidden) as the background picture. |
| Not expressible | A picture of that element (2× element screenshot, including what is behind it) at its rect; its descendants are not emitted. Applies to: canvas without a readable chart, inline `svg`, `iframe`/`video`/`audio`, `filter`/`backdrop-filter`/`mask`/`clip-path`/`mix-blend-mode`, skew/3D transforms, `background-clip: text`. Each prints `Notice: slide N: <element> exported as a picture (<reason>)`. |
| Same element on every slide | An image or text node with the same content and the same rect (±0.5 px) on all slides → drawn once on the slide master, removed from the slides. |
| Skipped | `display: none`, `visibility: hidden`, zero size, `opacity: 0`, outside the canvas, Slidev's own UI. |

## Writer extensions (`src/ooxml/*`)

- Per-slide background (`p:bg`): solid, gradient, or picture.
- Gradient fills (`gradFill` with alpha stops; linear `lin ang`, radial `path="circle"` +
  `fillToRect`) for shapes and backgrounds.
- Parameterized shadow (`outerShdw` blur/dist/dir/color/alpha) instead of the fixed preset.
- Shape rotation (`xfrm rot`).
- Pictures: `srcRect` crop, outline, shadow; SVG media (`image/svg+xml`) with the
  `asvg:svgBlip` extension next to a PNG blip.
- Text: run highlight, underline, strike; exact line spacing (`lnSpc/spcPts`); bullet color,
  indent and start per paragraph in free text boxes.
- Tables: per-cell fill and per-side borders, merges, and an explicit "No Style, No Grid"
  table style for render mode.
- Slide master: an arbitrary list of shapes (not just the logo).

Existing behavior for the static path must not change (all current tests keep passing).

## Configuration

- CLI `--render`; API `convert({ ..., render: true })`.
- `pptx.config.mjs` → `render: { timeout }` (per-slide wait, default 15000 ms). Nothing else is
  required.
- `colors`/`fonts` still build the theme; exact-match colors become theme slots. Without a
  config the default theme is used and colors are literal.
- `chartResolvers` override captured charts for their component.
- `logo`, `eyebrowClass`, `cssVarColors` are ignored in render mode (the page already shows
  their effect); a single notice says so when any is set.
- Dependencies: `playwright-chromium` and `playwright` as optional peer dependencies. No new
  runtime dependency is installed for static-mode users.

## Errors

| Condition | Behavior |
|---|---|
| Deck has no `@slidev/cli` | Stop: "render mode runs the deck's own Slidev — run `npm install` in <deck dir>" |
| Playwright missing | Stop: install `playwright-chromium`, then `npx playwright install chromium` |
| Dev server not ready in 60 s / exits | Stop, printing the last 50 lines of its output |
| Unsupported Slidev DOM | Stop: "unsupported Slidev version <version>" |
| One slide fails or times out | That slide uses the static layout (today's export) with `Warning: slide N could not be rendered (<reason>) - exported from markdown instead`; export continues |
| Any outcome | Server and browser are always shut down |

Determinism: the same machine, browser build and fonts produce byte-identical output (no
timestamps; nodes in DOM order). Across machines, output may differ where font rendering does.

## Phases

Each phase is usable on its own and is a separate section of the implementation plan.

1. **Core.** Server, capture, extractor and `sceneToSpec` for text, lists, boxes/cards,
   raster images, solid backgrounds, title placeholders, notes; picture fallback for everything
   else (so SVG logo, chart, cover glow and tables are pictures for now); writer: backgrounds,
   shadow, rotation, picture crop/outline/shadow, run highlight/underline/strike, line
   spacing, bullet color; CLI flag, config, errors, per-slide static fallback.
2. **Content.** Native tables (per-cell formatting, merges), Shiki code blocks, Chart.js bar
   charts (tick format, stacked, horizontal), `chartResolvers` precedence.
3. **Polish.** Native gradients (shapes and backgrounds), SVG via `svgBlip`, master
   de-duplication of repeated elements, and the visual check.

## Testing

1. `npm test` (no browser; runs anywhere): hand-written Scene JSON fixtures for every mapping
   rule → `sceneToSpec` assertions; writer unit tests for each extension; every generated
   package passes `lintPackage`.
2. `npm run test:render` (needs Playwright; skipped with a notice otherwise):
   - Extractor tests: `extract.js` runs in Chromium against small static HTML fixtures (cards,
     lists, inline styles, a Chart.js chart, an inline SVG) — no Slidev needed.
   - End-to-end: `--render` export of `../innotek-SlideDevBase` when checked out (as
     `npm run render` does today), linted.
3. Visual check (phase 3, a `--render` option of `npm run render`): per slide, Slidev's own
   screenshot vs LibreOffice's render of the exported slide, both reduced to 160×90 grayscale;
   similarity = 1 − mean absolute difference. Prints a score per slide; fails below 0.90 (the
   starting threshold, tuned once on the innotek deck).

## Success criteria (innotek deck, `--render`)

- 7 slides; all text is editable text.
- `kv-card` panels are rounded shapes whose fill and outline resolve to `#EEF4F0` and `#D9E3DB`
  (as theme colors where they match the config's colors).
- The eyebrow is uppercase, green (`accent1` via theme matching), letter-spaced.
- `ExampleChart` is a native, editable chart with no resolver configured (phase 2).
- The SVG logo is on the slide master once, as SVG with PNG fallback (phase 3).
- The cover glow is a native radial gradient (phase 3).
- No picture fallbacks needed for this deck (after phase 3).
- Every slide passes the visual check; every package passes `lintPackage`.

## Out of scope

Animations and transitions; hyperlinks (link text stays styled text); font embedding (viewers
need the fonts, as today); interactive component state beyond first render; chart types other
than bar/column natively; Slidev drawings.

## Risks

- **Font metrics differ** between Chromium and PowerPoint, so text can wrap differently. Mitigated
  by single-line boxes not wrapping and by the visual check; residual risk accepted.
- **Dev-mode reliance** for charts (`__vueParentComponent`). If a Vue/Slidev update removes it,
  charts fall back to pictures (still correct visually).
- **Slidev DOM changes** across versions. Guarded by the compatibility check; tested with
  `@slidev/cli` 52.x.
- **Slide size**: Slidev's 980×552 canvas is 16:9 to within 0.2%; content is scaled by width.
- **Speed**: one dev-server start plus one page load per slide (roughly 10–30 s per deck).
