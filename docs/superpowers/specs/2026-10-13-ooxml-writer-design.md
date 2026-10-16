# Own OOXML writer (replace pptxgenjs)

**Date:** 2026-10-13 · **Status:** approved design, awaiting implementation plan · **Target version:** 0.2.0

## Why

We studied how [DeckCraft](https://github.com/storytold/deckcraft) (a Rust clean-room PowerPoint
reimplementation) writes `.pptx`, and exported the `innotek-SlideDevBase` deck with the current
tool as a baseline. The two findings that drive this design:

1. **Our package and XML correctness is at the mercy of pptxgenjs.** Beyond the phantom
   `slideMaster` content-type overrides we already patch (`src/fix-pptx.mjs`), pptxgenjs 4.0.1's
   `genXmlTextBody` writes a `<a:pPr>` per *run*, so any bullet item with inline `**bold**`
   produces several `<a:pPr>` in one `<a:p>`. That is schema-invalid (one `pPr`, first), drops the
   bullet in LibreOffice, and is a likely PowerPoint Online repair trigger. It cannot be avoided
   through pptxgenjs options.
2. **Our output is not PowerPoint-native.** Every title and paragraph is a free text box, colors
   and fonts are stamped as literals on every run, and the theme is the stock Office one. So
   outline view, Accessibility Checker slide titles, Reset Slide, Designer, layout switching and
   Design ▸ Variants recoloring don't work.

DeckCraft avoids both by owning the writer: a part registry that makes content types and
relationships impossible to desync, a real master with typed layouts and placeholders, a theme
that shapes reference via `schemeClr` / `+mj-lt` / `+mn-lt`, schema-ordered text XML, built-in or
theme-based table styles, notes slides, and deterministic IDs. We adopt those ideas. We do **not**
copy its charts (no embedded workbook, so Edit Data doesn't work).

## Goals

- Replace pptxgenjs with an in-repo OOXML writer (`src/ooxml/`).
- Exported decks open without a repair prompt and with layout intact in **PowerPoint desktop,
  PowerPoint Online, Google Slides, Keynote and LibreOffice Impress**.
- Titles live in real title placeholders; text-only bodies live in real body placeholders
  (hybrid rule below); theme colors and fonts come from the config.
- Speaker notes from Slidev's `<!-- … -->` convention.
- `layout: center` (and `end`), `layout: section` and cover slides map to matching layouts.
- Byte-identical output for identical input.
- A test suite (there is none today), including a package lint that catches repair-prompt causes.

## Non-goals (sub-project 2, "Slidev coverage")

`two-cols` / `::right::` slots; nested lists and links in markdown; Shiki syntax highlighting;
transitions and `v-click` animations; `---` inside code fences being split as a slide separator;
stat-card alignment and inline-color bugs; `image-left`/`image-right` layouts; SVG images. The new
writer supports levels, hyperlinks and per-run colors, so these only need parser and layout work.

## Compatibility

Unchanged: `convert(options)` signature and return value, the CLI, the `pptx.config.mjs` shape
(including resolver kinds `statcards`, `barchart-panels`, `barchart`), and the `parse.mjs` block
model. New optional config keys: `lang` (default `"en-US"`) and `timestamps` (default `false`;
`true` writes the real export time into `docProps/core.xml`, breaking byte-identical output).
`innotek-SlideDevBase` must export with no config change.

Dependencies: drop `pptxgenjs`; keep `jszip`, `js-yaml`; add `image-size` (runtime, already in
the tree via pptxgenjs) and `@xmldom/xmldom` (dev only). `src/fix-pptx.mjs` is deleted.

## Architecture

```
slides.md ─► parse.mjs ─► layout.mjs ─► SlideSpec[] ─► src/ooxml/ ─► deck.pptx
```

- `parse.mjs`: unchanged, except it now also extracts speaker notes (see below).
- `layout.mjs`: replaces `render.mjs`. Keeps today's height estimation, card-row, stat-card and
  chart-panel geometry, but emits plain `SlideSpec` objects instead of calling pptxgenjs.
- `src/ooxml/`: the writer. Knows nothing about Slidev or markdown.

### SlideSpec (the contract between layout and writer)

All geometry in inches; the writer converts to EMU (914400/inch).

```js
SlideSpec = {
  layout: "title" | "content" | "titleOnly" | "centered" | "section" | "blank",
  placeholders: {
    title?: Para[],
    subtitle?: Para[],          // "title" layout only
    body?: Para[],              // text-only bodies (hybrid rule)
    // Optional slide-level position overrides, used by "centered" (and titleBox by render mode):
    titleBox?: Box, bodyBox?: Box,
    bodyFontScale?: number,     // 0..1, written as <a:normAutofit fontScale>
    // Render mode only (added 2026-10-15): the title placeholder at its measured DOM position.
    titleBody?: { inset?: number, anchor?: "t", autofit?: "none", wrap?: false }, // its bodyPr
    titleZ?: number,            // the title is drawn after this many shapes (DOM order), not first
  },
  shapes: Shape[],              // drawn after placeholders, in order
  notes?: string,
  background?: { color, alpha? } | { gradient } | { image: { data, ext } }, // render mode: p:bg
}
Box   = { x, y, w, h }
Shape =
  | { kind: "text", ...Box, geometry: "rect"|"roundRect"|"ellipse", radius?: number,
      fill?: Color, line?: { color: Color, width: number, dash?: "dash" } | null,
      shadow?: boolean, paras: Para[], anchor?: "t"|"ctr"|"b", inset?: number,
      charSpacing?: number }
  | { kind: "table", ...Box, colW?: number[], rows: Para[][][], header: boolean,
      variant?: "default"|"deflist" }
  | { kind: "image", ...Box, path: string, alt: string }
  | { kind: "chart", ...Box, chart: ChartSpec }
Para = { level?: 0..2, bullet?: "char"|"number"|"none", number?: number,
         align?: "l"|"ctr"|"r", spaceBefore?: number /* pt */, runs: Run[] }
Run  = { text, b?, i?, font?: "heading"|"body"|"mono", size?: number /* pt */,
         color?: Color, link?: string, caps?: boolean }
Color = "dk1"|"lt1"|"dk2"|"lt2"|"accent1".."accent6" | "#RRGGBB"
ChartSpec = { type: "bar"|"column", categories: string[],
              series: [{ name, color: Color, values: number[] }],
              min?, max?, valueAxis: boolean, legend: boolean, gapWidth, overlap?,
              dataLabels: { color: Color, size: number }, background: Color,
              catLabelColor: Color, catLabelSize: number }
```

`layout.mjs` resolves config color keys to `Color` via `toColor()`: a value equal to one of the
theme-mapped config colors becomes its slot name; anything else becomes `#hex`.

**Render-mode extensions (2026-10-15).** Render mode (`src/render/scene.mjs`, see
`2026-10-15-render-mode-design.md`) builds the same SlideSpecs with a few more optional fields,
all ignored when absent so the static export's output is unchanged: the slide `background` and
the title's `titleBox` / `titleBody` / `titleZ` above; on text shapes `fill` as a gradient
(`{ type: "linear"|"radial", angle | cx, cy, stops: [{ pos, color, alpha? }] }`), `fillAlpha`,
`line.alpha`, a parameterized `shadow` (`{ blur, dist, dir, color, alpha? }` instead of `true`),
`rot` and `wrap: false`; on paragraphs `lineSpacing` (pt), `bulletColor`, `marL` / `indent`;
on runs `typeface`, `alpha`, `spacing`, `u`, `strike` and `highlight`; on images `data` / `ext`
(bytes instead of a path), `crop` (`srcRect`), `line`, `shadow`, `rot` and `svg` (an
`asvg:svgBlip` next to the PNG); `variant: "plain"` tables whose cells carry their own fill,
per-side borders, margins, anchor and `gridSpan` / `rowSpan` / merge markers; chart
`reverseCategories`, `numFmt`, `gridColor` and series `alpha`; and `deckOpts.masterShapes`,
shapes drawn once on the slide master.

### Writer modules (`src/ooxml/`)

| Module | Responsibility |
|---|---|
| `package.mjs` | `Package` class. `addXml(partName, contentType, xml)`, `addMedia(partName, buffer)` (Default by extension), `rel(fromPart, toPart, type, {external})` → `rId` (per-part counter, identical rels reused). Builds `[Content_Types].xml` and every `_rels/*.rels` from what was actually added. Writes `[Content_Types].xml` first; fixed zip timestamps (1980-01-01); media stored, XML deflated. Relationship targets computed as relative paths from part names. |
| `xml.mjs` | `esc(text)` strips chars invalid in XML 1.0 (incl. lone surrogates, U+FFFE/FFFF) and escapes `& < > " '`; `emu(inches)`; `pt100(pt)`; `guid(seed)` deterministic FNV-1a based GUID. |
| `theme.mjs` | `theme1.xml`: 12-slot color scheme, font scheme (latin/ea/cs for major & minor), format scheme with the 3-entry lists the schema requires. |
| `master.mjs` | `slideMaster1.xml` (clrMap, `txStyles`, logo picture) and the 7 layouts. |
| `text.mjs` | `txBody(paras, bodyOpts)` with schema-ordered `bodyPr` / `pPr` / `rPr`; exactly one `pPr` per `a:p`, first; empty bodies get one `<a:p/>`. |
| `slide.mjs` | Slide XML: spTree header, placeholders, shapes, pictures, `graphicFrame`s; shape-id allocation (from 2, unique per part); `p:clrMapOvr`. |
| `table.mjs` | `a:tbl` + `tableStyles.xml` deck table styles. |
| `chart.mjs` | `chartN.xml` (`c:chartSpace`, caches, `c:externalData` + `c:autoUpdate val="0"`). |
| `xlsx.mjs` | Minimal embedded workbook for one chart (content types, rels, workbook, one sheet with inline strings, minimal styles). |
| `notes.mjs` | `notesMaster1.xml` (+ its own theme part) and `notesSlideN.xml`, written only if any slide has notes. |
| `presentation.mjs` | `presentation.xml` (sldIdLst from 256, master/layout ids from 2147483648 sharing one counter, `sldSz`, `notesSz`, `defaultTextStyle`), `presProps.xml`, `viewProps.xml`, `docProps/app.xml`, `docProps/core.xml`. |
| `index.mjs` | `writePptx(slideSpecs, deckOpts) → Buffer`. |

## Theme

| Slot | Source | Slot | Source |
|---|---|---|---|
| `dk1` | `colors.text` | `accent1` | `colors.accent` |
| `lt1` | `FFFFFF` | `accent2` | `colors.accent2` |
| `dk2` | `colors.accentDark` | `accent3` | `colors.muted` |
| `lt2` | `colors.panel` | `accent4` | `colors.border` |
| `hlink` | `colors.accent` | `accent5` | `9CA3AF` |
| `folHlink` | `colors.accentDark` | `accent6` | `F59E0B` |

Major font = `fonts.heading`, minor font = `fonts.body`. Runs reference `+mj-lt` (title, eyebrow)
or `+mn-lt` (everything else); `font: "mono"` writes `fonts.mono` literally. Fonts are written on
runs explicitly (as theme references) rather than relying only on inheritance, because Google
Slides and Keynote honor run-level theme font references more reliably than `defaultTextStyle`.

## Master and layouts

Master: white `bg1` background; `p:clrMap` standard mapping (bg1=lt1, tx1=dk1, bg2=lt2, tx2=dk2,
accents, hlinks). `txStyles`:

- `titleStyle`: `+mj-lt`, 26 pt, bold, `tx1`, anchor bottom.
- `bodyStyle` levels 1–3: `+mn-lt`, 13 / 12 / 11 pt, `tx1`, `spcBef` 6 pt, `buChar "•"` with
  `buClr accent1`, `marL`/`indent` of 0.25"/0.5"/0.75" with -0.2" hanging indent.
- `otherStyle`: `+mn-lt`, `tx1`.

The logo (if configured and present) is a `p:pic` on the master at today's position. Layouts
inherit it (`showMasterSp` default).

Slide geometry: width/height from `config.slide`, margins from `config.slide.marginX`. Layouts
always reserve the eyebrow band (`y = 0.45`, 0.3" tall), like a real template would, so no
slide-level overrides are needed for it: title placeholder at `y = 0.80`, 0.6" tall (today's
position on eyebrow slides); body placeholder from `y = 1.60` to `height - 0.35`. Slides without
an eyebrow therefore get their title 0.35" lower than today, consistent across the deck.

| Layout name | `type` | Placeholders | Used when |
|---|---|---|---|
| Title Slide | `title` | `ctrTitle`, `subTitle idx=1` | first slide; `layout: cover` or `intro`; `class` contains `cover` |
| Title and Content | `obj` | `title`, body `idx=1` | default layout, text-only body |
| Title Only | `titleOnly` | `title` | default layout, rich body |
| Centered | `cust` | `title`, body `idx=1`, centered band | `layout: center` or `end` |
| Section Header | `secHead` | `title`, body `idx=1` | `layout: section` |
| Two Content | `twoObj` | `title`, body `idx=1`, body `idx=2` | not mapped yet (sub-project 2) |
| Blank | `blank` | none | slide with no title and not cover |

Layout placeholders carry geometry, anchor and `<a:normAutofit/>` (body) so slides can leave
`<p:spPr/>` empty and inherit. Placeholders always get `<a:spLocks noGrp="1"/>`.

**Eyebrow.** Free text box (`txBox="1"`) in the reserved eyebrow band, `+mj-lt`, 9 pt, bold,
caps, `accent1`, `spc` 150.

**Hybrid body rule.** A body is *text-only* when every block is `paragraph`, `bullets` or
`heading`. Text-only bodies go into the body placeholder:

- `paragraph` → `bullet: "none"` (written as `marL=0 indent=0 <a:buNone/>`).
- `bullets` → one Para per item, `bullet: "char"` (inherits master bullet) or `"number"`
  (`<a:buAutoNum type="arabicPeriod"/>`, `startAt` only when not 1).
- `heading` → bold run, size body+5 (`##`) or body+2 (`###`), `dk2`, `bullet: "none"`.
- `opacity` blocks → runs colored with the white-blended hex, as today.
- Runs in placeholders carry no `size` (except headings) so they inherit the master's 13 pt;
  dense slides shrink through `bodyFontScale`, not through per-run sizes.

A `blank` slide (no title, not cover) draws its whole body as shapes.

Everything else (`table`, `code`, `image`, `chart`, `cardrow`, `deflist`, `styledtext`, `quote`)
makes the body *rich*: layout `titleOnly` (or `centered`), body drawn as shapes with today's
geometry.

**Overflow.** For placeholder bodies, `layout.mjs` estimates height with the existing estimator;
if it exceeds the placeholder, `bodyFontScale = max(0.55, available / estimated)` is written as
`<a:normAutofit fontScale="…" lnSpcReduction="10000"/>` (only when < 1). For rich bodies the
existing scale-down logic stays.

**Centered.** `layout.mjs` computes the total stack height (eyebrow + title + body) and writes
slide-level `titleBox` / `bodyBox` overrides (and offsets shapes) so the stack is vertically
centered. `class: text-center` sets `align: "ctr"` on title and body paragraphs. Reset Slide
returns to the Centered layout's default geometry.

## Content rules

**Text XML.** `bodyPr` attributes only when set (`wrap="square"`, insets, `anchor`); children in
schema order (`noAutofit`|`spAutoFit`|`normAutofit`). `pPr` children: `lnSpc`, `spcBef`, `spcAft`,
`buClr`, `buSzPct`, `buFont`, bullet element. `rPr` attributes `lang`, `sz`, `b`, `i`, `cap`,
`spc`, `dirty="0"`; children `solidFill`, `latin`, `ea`, `cs`, `hlinkClick` (external rel). `\n`
inside run text becomes `<a:br/>`. Free text boxes get `txBox="1"` and explicit `marL`/`indent`/
`buChar` for bullets since they don't inherit `bodyStyle`. Every run gets `lang` (`config.lang`).

**Shapes.** `roundRect` radius → `avLst` `adj` = `radius / min(w, h) * 100000` (clamped
0–50000). Card shadow → `a:effectLst/a:outerShdw` (blurRad 9pt, dist 3pt, dir 90°, `dk2` at 16%
alpha). Line `none` → `<a:ln><a:noFill/></a:ln>`.

**Tables.** `p:graphicFrame` with `<a:graphicFrameLocks noGrp="1"/>`; `a:tblPr firstRow="1"`
(header) and `<a:tableStyleId>` of a deck style defined in `tableStyles.xml`: header row `lt2` fill
with bold `dk2` text, all borders `accent4` 1 pt, body cells no fill. A second `deflist` style:
`lt2` fill all cells, `accent4` 0.5 pt borders. Style GUIDs are deterministic. Rows padded to the
grid width. Cell margins as `marL/marR/marT/marB` on `tcPr`. Cells carry no direct fills or
borders (so Table Design and theme recoloring work). **Risk:** Keynote and Google Slides may
ignore custom table styles; if the manual spot check shows unstyled tables, add direct cell
borders/fills as a fallback.

**Charts.** `ppt/charts/chartN.xml` + `ppt/embeddings/Microsoft_Excel_Worksheet{N}.xlsx`, related
from the chart part (`package` relationship) and referenced by `<c:externalData r:id>`.
`c:barChart` with `barDir`, `grouping clustered`, `varyColors 0`, one `c:ser` per series
(`c:tx` strRef `Sheet1!$B$1`…, `c:cat` strRef `Sheet1!$A$2:$A$n` with `strCache`, `c:val` numRef
with `numCache`, `formatCode General`), `c:dLbls` (outEnd, color/size from spec), `gapWidth`,
`overlap`, axis ids fixed per chart, `c:catAx` / `c:valAx` with `delete` for a hidden value axis,
no gridlines, no axis lines, `c:legend` bottom when `legend`, chart and plot area fill from
`background`. Series fill uses `schemeClr` when the color maps to a theme slot. Non-finite values
are omitted from caches. The slide references the chart via `p:graphicFrame` →
`c:chart r:id`.

**Images.** `p:pic` with `cNvPr descr` = alt text (also `name`), `a:blip r:embed`,
`a:stretch/a:fillRect`, size fitted with "contain" inside the box using `image-size`. Media named
`ppt/media/imageN.<ext>`; identical files are stored once. PNG/JPEG/GIF supported; SVG, missing
files and remote URLs render today's labeled placeholder box.

**Notes.** `parse.mjs`: after removing code fences, the slide's trailing HTML comment(s)
(`<!-- … -->` at the end of the slide content, Slidev's notes convention) become `notes`, with
leading/trailing whitespace trimmed and internal blank lines preserved as paragraph breaks.
Comments elsewhere in the slide are discarded as today. Notes slide: `sldImg` placeholder + body
`idx=1` with the text; rels to the notes master and the slide.

**IDs and determinism.** Shape ids from 2 per part; `sldId` from 256; master + layout ids share a
counter from 2147483648; GUIDs via `guid(seed)`; `docProps/core.xml` dates fixed unless
`config.timestamps === true`. Same input ⇒ identical bytes.

## Testing

Runner: `node --test` (`npm test`). Dev dependency: `@xmldom/xmldom`.

1. **Unit:** `esc()` (control characters, lone surrogates, U+FFFE), `Package` (overrides match
   parts, rels resolve, rId reuse), `toColor()` mapping, `text.mjs` (one `pPr` first, schema
   order), `xlsx.mjs` (zip contains the expected parts).
2. **Package lint** (`test/lint.mjs`, run against every generated deck in the suite):
   - every XML part parses;
   - every `Override` names an existing part and every part has a content type;
   - every relationship target exists; every `r:id` / `r:embed` / `r:link` used in a part is
     defined in that part's rels;
   - `cNvPr` ids unique per part; `sldId` ≥ 256; master/layout ids ≥ 2147483648 and unique;
   - every `a:p` has at most one `a:pPr` and it is the first child; every `txBody` has ≥ 1 `a:p`.
3. **Fixture deck** `test/fixtures/kitchen-sink/` (slides.md + config + chart JSON + a PNG):
   every block type, cover / default text-only / default rich / center / section / no-title
   slides, notes, a chart. Assertions: layout name and placeholder types per slide, notes text,
   chart part has an embedded xlsx, title text lives in a `type="title"` / `ctrTitle`
   placeholder, two exports are byte-identical.
4. **LibreOffice render** (`npm run render`, `scripts/render.mjs`): exports the fixture and, when
   present, `../innotek-SlideDevBase`, converts to PDF and PNG contact sheets for visual review,
   and asserts page count and expected strings via `pdftotext`. Skipped when `soffice` is missing.
   No pixel-diff assertions.
5. **Manual spot check** (owner): open the fixture and innotek exports in PowerPoint desktop,
   PowerPoint Online, Google Slides and Keynote; check no repair prompt, layouts, theme colors in
   Design ▸ Variants, table style, Edit Data on the chart, notes, outline view titles.

## Risks

- **Custom table styles** in Keynote / Google Slides (fallback described above).
- **`normAutofit fontScale`**: PowerPoint recomputes on edit; Google Slides may ignore it, leaving
  overflow on very dense text-only slides. Mitigation: the estimator is conservative, and the
  scale also caps at 0.55 like today.
- **Embedded xlsx validity**: Excel is strict. Mitigation: minimal workbook modeled on what
  pptxgenjs emits today (known to open), checked by unit test and by the manual Edit Data check.
- **Visual drift** from today's output: positioned shapes keep today's geometry; placeholder
  bodies may wrap slightly differently; titles on slides without an eyebrow sit 0.35" lower.
  Reviewed via the LibreOffice contact sheets before/after.
