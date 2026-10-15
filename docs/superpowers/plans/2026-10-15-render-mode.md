# Render Mode Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** `slidev-pptx-export slides.md --render` runs the deck with its own Slidev in headless Chromium and converts what is actually drawn — including custom CSS, HTML and Vue components — into native, editable PowerPoint content, with pictures only for what PowerPoint can't express.

**Architecture:** `src/render/server.mjs` starts the deck's own Slidev dev server; `src/render/capture.mjs` drives Playwright Chromium through every slide and runs `src/render/extract.js` (a classic in-page script) to produce a JSON *Scene* per slide; `src/render/scene.mjs` turns a Scene into the SlideSpecs the existing OOXML writer (`src/ooxml/*`) already draws, which gains a few features (backgrounds, shadows, gradients, per-cell tables, SVG pictures, master shapes). `src/render/index.mjs` orchestrates and falls back to the static export per slide.

**Tech Stack:** Node ≥ 18 ESM, `node --test`, Playwright (`playwright-chromium`, optional peer dependency), the deck's own `@slidev/cli` (tested with 52.x), existing `jszip`/`image-size`.

**Spec:** `docs/superpowers/specs/2026-10-15-render-mode-design.md`

## Global Constraints

- Without `--render`, behavior and output are unchanged: every existing test keeps passing, including the byte-identical export test.
- No new runtime dependency. Playwright is an *optional peer dependency* (`playwright-chromium`, `playwright`), resolved at runtime from the deck folder first, then this package; it is a devDependency here only for `npm run test:render`.
- `npm test` must not need a browser. Browser tests live in `test/render/` and run with `npm run test:render`; they skip (never fail) when Playwright/Chromium can't launch or `../innotek-SlideDevBase` isn't checked out.
- Units: canvas px → inches with `k = slideWidthIn / canvas.w` (13.333 / 980); points = px × k × 72. Colors go through `toColor(hex, theme)`, so a color equal to a theme color becomes that theme slot.
- Exact message formats:
  - `Notice: slide N: <label> exported as a picture (<reason>).`
  - `Notice: slide N: image <src> didn't load - skipped.`
  - `Notice: render mode ignores <keys> from pptx.config.mjs - the rendered page already shows them.`
  - `Warning: slide N could not be rendered (<reason>) - exported from markdown instead.`
  - `Notice: slide master: <label> exported as a picture (<reason>).` (Task 12)
  - N is the slide's position in the exported `.pptx`.
- `src/render/extract.js` is a classic browser script (no `import`/`export`, no build step), loaded with `page.addScriptTag({ content })`; it defines `window.__pptxExtract(pageEl, canvas)`.
- Every commit message ends with a `Co-Authored-By: Claude <model> <noreply@anthropic.com>` line naming the model that wrote the commit.
- Tool quirk: the Write/Edit tools turn backslash-u escape sequences (`\u` followed by four hex digits) into literal characters. Nothing in this plan needs one; don't introduce any.

### Facts verified against Slidev 52.20 (from the design probes)

- `/<n>?print` renders slide `n` alone, with every `v-click` fragment shown (final click state). A viewport equal to the canvas size gives `--slidev-slide-scale: 1`.
- `window.__slidev__.nav.total` is the slide count (hidden slides excluded); `nav.slides[i]` has `no`, `meta.slide.note` (speaker notes, raw markdown) and `meta.slide.sourceIndex` (the slide's index in the markdown, hidden slides included).
- `window.__slidev__.configs.canvasWidth` (980) and `.aspectRatio` (16/9) give the canvas size; the slide element is `.slidev-page-<n>`, and its layout root carries `slidev-layout <name>` classes.
- A Chart.js chart drawn by vue-chartjs is reachable from its canvas: walk `canvas.__vueParentComponent` (then `.parent`) to the component whose `exposed.chart` (a shallowRef: use `.value`) has `config`, `data`, `options`.
- Slidev exits immediately (non-zero, message on stderr) when the deck's theme isn't installed.

### Spec gaps resolved here

- **Non-16:9 decks:** the spec keeps `config.slide` (16:9). A 4:3 deck would then be cut off, so the slide height follows the canvas aspect when they differ by more than 1% (Task 6, Review Focus).
- **Phase 1 pictures:** until phases 2–3 land, tables, code blocks, canvases (charts), inline SVG, SVG images and gradients are pictures. Each later task removes one of those from the raster list.
- **`chartResolvers` in render mode:** only resolvers returning `kind: "barchart"` override the captured chart. The other kinds describe HTML the page already renders.

## Review Focus

Inputs the spec implies but no feature test would otherwise exercise, most likely first. Each has a test in the task named.

1. **Slidev fails to start** (deck theme not installed, syntax error in a layout): the export stops and shows Slidev's own error lines — Task 5.
2. **`@slidev/cli` hoisted into a parent folder** (npm/pnpm workspaces, monorepos): it is still found — Task 5.
3. **An image that doesn't load** (404, blocked host): it is skipped with a notice and the rest of the slide exports — Tasks 4 and 6.
4. **Non-16:9 decks** (`aspectRatio: 4/3`): the slide height follows the canvas, so nothing is cut off — Task 6.
5. **A custom layout without the `slidev-layout` class**: the slide is still extracted, with a white background — Task 4.

---

## File Structure

```
src/
  render/
    server.mjs     CREATE  find the deck's @slidev/cli, start its dev server on a free port, stop it
    browser.mjs    CREATE  resolve Playwright's chromium at runtime (deck first, then this package)
    extract.js     CREATE  in-page DOM -> Scene extractor (classic script, window.__pptxExtract)
    capture.mjs    CREATE  per-slide navigation and waits, extraction, image bytes, screenshots
    scene.mjs      CREATE  sceneToSpec(): Scene -> SlideSpec (pure; unit-tested)
    index.mjs      CREATE  renderSpecs(): orchestration, static fallback, notices, master de-dup
    compare.mjs    CREATE  similarity of two PNGs (phase 3 visual check)
  ooxml/
    text.mjs       MODIFY  typeface, underline, strike, highlight, run alpha, line spacing, bullet color, margins
    shapes.mjs     MODIFY  fill/line alpha, shadow params, rotation, picture crop/outline/shadow/SVG, backgrounds
    xml.mjs        MODIFY  gradient fills
    index.mjs      MODIFY  media from buffers, slide backgrounds, SVG pictures, master shapes
    table.mjs      MODIFY  "plain" variant with per-cell fill/borders/margins/anchor/merges
    chart.mjs      MODIFY  number format, legend position, stacking, gridlines, transparent background
    master.mjs     MODIFY  extra master shapes
    package.mjs    MODIFY  SVG media type
  layout.mjs       MODIFY  layoutEach(); groupedBarChart() exported
  index.mjs        MODIFY  convert({ render })
bin/cli.mjs        MODIFY  --render
scripts/render.mjs MODIFY  --render option: render-mode export + visual check
test/
  text.test.mjs, shapes.test.mjs, writer.test.mjs, table.test.mjs, chart.test.mjs   MODIFY
  render-scene.test.mjs    CREATE  sceneToSpec rules
  render-server.test.mjs   CREATE  server/browser resolution and startup errors (fake CLI, no browser)
  render-index.test.mjs    CREATE  renderSpecs with injected capture
  render/helpers.mjs       CREATE  usable-chromium check, extractHtml()
  render/extract.test.mjs  CREATE  extractor against static HTML
  render/deck.test.mjs     CREATE  end-to-end on ../innotek-SlideDevBase
  render/compare.test.mjs  CREATE  PNG similarity
package.json       MODIFY  test:render script, optional peer deps, playwright-chromium devDependency, version
README.md          MODIFY  render mode
```

---

# Phase 1 — Core

### Task 1: Writer — text runs and paragraphs for measured text

**Files:**
- Modify: `src/ooxml/text.mjs`
- Modify: `src/ooxml/shapes.mjs` (pass `wrap` through), `src/ooxml/index.mjs` (title placeholder body properties)
- Test: `test/text.test.mjs`, `test/shapes.test.mjs`, `test/writer.test.mjs`

**Interfaces:**
- Consumes: `solidFill(color, alpha)`, `colorXml(color, alpha)`, `emu`, `pt100` from `src/ooxml/xml.mjs`.
- Produces (fields later tasks put on SlideSpec text):
  - Run: `typeface` (any font name; wins over `font`), `u` (underline), `strike`, `alpha` (0..1, with `color`), `highlight` (color).
  - Paragraph: `lineSpacing` (pt, exact), `bulletColor` (color), `marL`/`indent` (inches, free text boxes only).
  - Text shape: `wrap: false` → `wrap="none"`. `bodyPrXml` writes `wrap="none"` for placeholders too when asked.
  - SlideSpec `placeholders.titleBody`: body properties for the title placeholder (`{ wrap, inset, anchor }`). Render mode measures titles tightly, so the master's insets, wrapping and bottom anchor must not apply.

- [ ] **Step 1: Write the failing tests**

In `test/text.test.mjs`, change the import line to:

```js
import { txBodyXml, bodyPrXml, rPrXml, pPrXml } from "../src/ooxml/text.mjs";
```

and append:

```js
test("runs can name any typeface, and carry underline, strike, alpha and a highlight in schema order", () => {
  const xml = rPrXml({ text: "x", typeface: "Plus Jakarta Sans", color: "#112233", alpha: 0.5, highlight: "#F5F5F5", u: true, strike: true }, ctx);
  assert.equal(
    xml,
    '<a:rPr lang="en-US" u="sng" strike="sngStrike" dirty="0"><a:solidFill><a:srgbClr val="112233"><a:alpha val="50000"/></a:srgbClr></a:solidFill>' +
      '<a:highlight><a:srgbClr val="F5F5F5"/></a:highlight><a:latin typeface="Plus Jakarta Sans"/><a:cs typeface="Plus Jakarta Sans"/></a:rPr>',
  );
});

test("a typeface wins over a font role, and the last run's typeface carries into endParaRPr", () => {
  const xml = txBodyXml([{ runs: [{ text: "a", typeface: "Inter", font: "heading", size: 12 }] }], ctx);
  assert.match(xml, /<a:endParaRPr lang="en-US" sz="1200" dirty="0"><a:latin typeface="Inter"\/><a:cs typeface="Inter"\/><\/a:endParaRPr>/);
  assert.doesNotMatch(xml, /\+mj-lt/);
});

test("free text box paragraphs take exact line spacing, a bullet color and measured margins", () => {
  assert.equal(
    pPrXml({ bullet: "char", bulletColor: "#1A833B", lineSpacing: 21, marL: 0.3, indent: -0.15 }, ctx),
    '<a:pPr marL="274320" indent="-137160"><a:lnSpc><a:spcPts val="2100"/></a:lnSpc><a:buClr><a:srgbClr val="1A833B"/></a:buClr><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>',
  );
  assert.equal(
    pPrXml({ bullet: "number", number: 3, bulletColor: "accent1" }, ctx),
    '<a:pPr marL="228600" indent="-182880"><a:buClr><a:schemeClr val="accent1"/></a:buClr><a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod" startAt="3"/></a:pPr>',
  );
  assert.equal(pPrXml({ lineSpacing: 18.5 }, ctx), '<a:pPr><a:lnSpc><a:spcPts val="1850"/></a:lnSpc></a:pPr>');
  assert.equal(pPrXml({ bullet: "none", marL: 0.5 }, ctx), '<a:pPr marL="457200" indent="0"/>');
});
```

Also append to `test/text.test.mjs`:

```js
test("placeholders can switch off wrapping and insets too (render mode's measured titles)", () => {
  assert.equal(bodyPrXml({ placeholder: true, wrap: false, inset: 0, anchor: "t" }), '<a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"/>');
  assert.equal(bodyPrXml({ placeholder: true }), "<a:bodyPr/>");
});
```

Append to `test/writer.test.mjs`:

```js
test("a slide's title placeholder can carry its own body properties", async () => {
  const buffer = await writePptx([{ layout: "titleOnly", placeholders: { title: t("Measured"), titleBox: { x: 1, y: 1, w: 3, h: 0.5 }, titleBody: { wrap: false, inset: 0, anchor: "t" } }, shapes: [] }], DECK);
  assert.match((await unzip(buffer)).text("ppt/slides/slide1.xml"), /<p:ph type="title"\/>[\s\S]*?<a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"\/>/);
});
```

In `test/shapes.test.mjs` append:

```js
test("wrap: false on a free text shape writes wrap=\"none\"", () => {
  const xml = shapeXml({ x: 0, y: 0, w: 1, h: 1, wrap: false, paras: [{ runs: [{ text: "x" }] }] }, 5, ctx);
  assert.match(xml, /<a:bodyPr wrap="none" anchor="t" rtlCol="0"\/>/);
});
```

- [ ] **Step 2: Run them to see them fail**

Run: `node --test test/text.test.mjs test/shapes.test.mjs test/writer.test.mjs`
Expected: the six new tests FAIL (no `u`/`strike`/highlight/typeface support; no `lnSpc`; `wrap` ignored; no `titleBody`).

- [ ] **Step 3: Implement**

In `src/ooxml/text.mjs`:

1. Change the import to `import { esc, emu, pt100, solidFill, colorXml } from "./xml.mjs";`
2. Add after `fontXml`:

```js
// An explicit font name (render mode measures whatever font the page used); theme fonts use
// the font roles above instead.
function typefaceXml(name) {
  const f = esc(name);
  return `<a:latin typeface="${f}"/><a:cs typeface="${f}"/>`;
}
```

3. Replace `rPrXml` with:

```js
export function rPrXml(run, ctx, tag = "a:rPr") {
  let attrs = ` lang="${esc(ctx.lang)}"`;
  if (run.size) attrs += ` sz="${Math.min(400000, Math.max(100, pt100(run.size)))}"`;
  if (run.b) attrs += ' b="1"';
  if (run.i) attrs += ' i="1"';
  if (run.u) attrs += ' u="sng"';
  if (run.strike) attrs += ' strike="sngStrike"';
  if (run.spacing) attrs += ` spc="${pt100(run.spacing)}"`;
  attrs += ' dirty="0"';
  // Child order per CT_TextCharacterProperties: fill, effects, highlight, ..., latin, ea, cs, sym, hlinkClick.
  let kids = "";
  if (run.color) kids += solidFill(run.color, run.alpha);
  if (run.highlight) kids += `<a:highlight>${colorXml(run.highlight)}</a:highlight>`;
  kids += run.typeface ? typefaceXml(run.typeface) : fontXml(run.font, ctx);
  if (run.link) {
    if (!ctx.link) throw new Error("Run has a link but the text context has no link() callback");
    kids += `<a:hlinkClick r:id="${ctx.link(run.link)}"/>`;
  }
  return kids ? `<${tag}${attrs}>${kids}</${tag}>` : `<${tag}${attrs}/>`;
}
```

4. Replace `pPrXml` with:

```js
export function pPrXml(para, ctx, { placeholder = false } = {}) {
  const level = Math.max(0, Math.min(8, para.level || 0));
  const attrs = [];
  const kids = [];
  const bullet = para.bullet;
  const bulleted = bullet === "char" || bullet === "number";

  if (bullet === "none" && placeholder) {
    // Body placeholders inherit bullets from the master's bodyStyle, so a plain paragraph has
    // to switch them off and drop the hanging indent explicitly.
    attrs.push('marL="0"', 'indent="0"');
  } else if (!placeholder && (bulleted || para.marL != null)) {
    // Free text boxes inherit no list style, so they carry their own indent (render mode
    // passes the measured one).
    attrs.push(`marL="${emu(para.marL ?? 0.25 * (level + 1))}"`, `indent="${emu(para.indent ?? (bulleted ? -0.2 : 0))}"`);
  }
  if (level > 0) attrs.push(`lvl="${level}"`);
  if (para.align && para.align !== "l") attrs.push(`algn="${para.align}"`);

  // Child order per CT_TextParagraphProperties: lnSpc, spcBef, spcAft, buClr, buSz, buFont, bullet.
  if (para.lineSpacing != null) {
    kids.push(`<a:lnSpc><a:spcPts val="${Math.min(158400, Math.max(0, pt100(para.lineSpacing)))}"/></a:lnSpc>`);
  }
  if (para.spaceBefore != null) kids.push(`<a:spcBef><a:spcPts val="${pt100(para.spaceBefore)}"/></a:spcBef>`);
  const buClr = para.bulletColor ? `<a:buClr>${colorXml(para.bulletColor)}</a:buClr>` : null;
  if (bullet === "char" && !placeholder) {
    kids.push(buClr || '<a:buClr><a:schemeClr val="accent1"/></a:buClr>', '<a:buFont typeface="Arial"/>', '<a:buChar char="•"/>');
  } else if (bullet === "number") {
    if (buClr) kids.push(buClr);
    const startAt = para.number && para.number !== 1 ? ` startAt="${para.number}"` : "";
    kids.push('<a:buFont typeface="+mj-lt"/>', `<a:buAutoNum type="arabicPeriod"${startAt}/>`);
  } else if (bullet === "none" && placeholder) {
    kids.push("<a:buNone/>");
  }

  if (!attrs.length && !kids.length) return "";
  const a = attrs.length ? ` ${attrs.join(" ")}` : "";
  return kids.length ? `<a:pPr${a}>${kids.join("")}</a:pPr>` : `<a:pPr${a}/>`;
}
```

5. In `paraXml`, change the `end` line to:

```js
  const end = rPrXml({ size: last.size, font: last.font, typeface: last.typeface }, ctx, "a:endParaRPr");
```

6. In `bodyPrXml`, replace

```js
  if (!body.placeholder) {
    if (body.wrap === false) attrs.push('wrap="none"');
    else attrs.push('wrap="square"');
  }
```

with

```js
  // Placeholders inherit wrapping from their layout unless told otherwise.
  if (body.wrap === false) attrs.push('wrap="none"');
  else if (!body.placeholder) attrs.push('wrap="square"');
```

In `src/ooxml/index.mjs`, in the placeholder loop, change the `body` line to:

```js
      const body = p.role === "body" && ph.bodyFontScale ? { fontScale: ph.bodyFontScale } : p.role === "title" && ph.titleBody ? ph.titleBody : {};
```

In `src/ooxml/shapes.mjs`, in `shapeXml`, change the `body` line to:

```js
  const body = txBodyXml(shape.paras || [], ctx, { anchor: shape.anchor || "t", inset: shape.inset, wrap: shape.wrap });
```

- [ ] **Step 4: Run all tests**

Run: `npm test`
Expected: all pass, including the byte-identical fixture export (static output is unchanged).

- [ ] **Step 5: Commit**

```bash
git add src/ooxml/text.mjs src/ooxml/shapes.mjs src/ooxml/index.mjs test/text.test.mjs test/shapes.test.mjs test/writer.test.mjs
git commit -m "Writer: any typeface, underline/strike/highlight/alpha runs, exact line spacing, measured list margins, title body properties"
```

---

### Task 2: Writer — shapes, pictures and slide backgrounds for measured elements

**Files:**
- Modify: `src/ooxml/shapes.mjs`, `src/ooxml/index.mjs`
- Test: `test/shapes.test.mjs`, `test/writer.test.mjs`

**Interfaces:**
- Consumes: Task 1 writer.
- Produces:
  - Text/geometry shape: `fillAlpha`, `line.alpha`, `shadow: true | { blur, dist (inches), dir (degrees, 0 = right, 90 = down), color, alpha }`, `rot` (degrees, any sign).
  - Image shape: `{ kind: "image", x, y, w, h, path }` **or** `{ ..., data: Buffer, ext: "png"|"jpg"|"gif" }`, plus optional `crop: { l, t, r, b }` (fractions), `line`, `shadow`, `rot`, `alt`.
  - SlideSpec `background`: `{ color, alpha? }` or `{ image: { data, ext } }`.
  - Exports `shadowXml(shadow)` and `backgroundXml(bg, rId?)` from `shapes.mjs`; `slideXml(shapesXml, bgXml = "")`.

- [ ] **Step 1: Write the failing tests**

In `test/shapes.test.mjs` append:

```js
test("shape fill and outline alpha, a measured shadow and a rotation", () => {
  const xml = shapeXml({
    x: 1, y: 1, w: 2, h: 1, rot: -90, fill: "lt2", fillAlpha: 0.5,
    line: { width: 1, color: "#D9E3DB", alpha: 0.4 },
    shadow: { blur: 0.1, dist: 0.05, dir: 90, color: "#14181A", alpha: 0.14 }, paras: [],
  }, 3, ctx);
  assert.match(xml, /<a:xfrm rot="16200000">/);
  assert.match(xml, /<a:solidFill><a:schemeClr val="bg2"><a:alpha val="50000"\/><\/a:schemeClr><\/a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="D9E3DB"><a:alpha val="40000"\/><\/a:srgbClr><\/a:solidFill><\/a:ln>/);
  assert.match(xml, /<a:effectLst><a:outerShdw blurRad="91440" dist="45720" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="14181A"><a:alpha val="14000"\/><\/a:srgbClr><\/a:outerShdw><\/a:effectLst>/);
});

test("pictures can be cropped, outlined and shadowed; a plain picture is unchanged", () => {
  const xml = pictureXml({ id: 4, name: "P", rId: "rId2", box: { x: 0, y: 0, w: 1, h: 1 }, crop: { l: 0.1, t: 0, r: 0.1, b: 0 }, line: { width: 0.75, color: "accent4" }, shadow: true });
  assert.match(xml, /<a:blip r:embed="rId2"\/><a:srcRect l="10000" r="10000"\/><a:stretch>/);
  assert.match(xml, /<\/a:prstGeom><a:ln w="9525"><a:solidFill><a:schemeClr val="accent4"\/><\/a:solidFill><\/a:ln><a:effectLst>/);
  const plain = pictureXml({ id: 4, name: "P", rId: "rId2", box: { x: 0, y: 0, w: 1, h: 1 } });
  assert.match(plain, /<a:blip r:embed="rId2"\/><a:stretch><a:fillRect\/><\/a:stretch><\/p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"\/><a:ext cx="914400" cy="914400"\/><\/a:xfrm><a:prstGeom prst="rect"><a:avLst\/><\/a:prstGeom><\/p:spPr>/);
});
```

In `test/writer.test.mjs`, add `import { readFileSync } from "node:fs";` to the imports and append:

```js
test("images from buffers share media with file images; slides take solid and picture backgrounds", async () => {
  const data = readFileSync(PIC);
  const slides = [
    { layout: "blank", background: { color: "#FAFAFA" }, shapes: [{ kind: "image", x: 1, y: 1, w: 2, h: 1, path: PIC }, { kind: "image", x: 4, y: 1, w: 2, h: 1, data, ext: "png", crop: { l: 0.25, t: 0, r: 0.25, b: 0 } }] },
    { layout: "blank", background: { image: { data, ext: "png" } }, shapes: [] },
  ];
  const buffer = await writePptx(slides, { ...DECK, logo: null });
  assert.deepEqual(await lintPackage(buffer), []);
  const z = await unzip(buffer);
  assert.deepEqual(z.names.filter((n) => n.startsWith("ppt/media/")), ["ppt/media/image1.png"]);
  assert.match(z.text("ppt/slides/slide1.xml"), /<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FAFAFA"\/><\/a:solidFill><a:effectLst\/><\/p:bgPr><\/p:bg><p:spTree>/);
  assert.match(z.text("ppt/slides/slide1.xml"), /<a:srcRect l="25000" r="25000"\/>/);
  assert.match(z.text("ppt/slides/slide2.xml"), /<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rId2"\/><a:srcRect\/><a:stretch><a:fillRect\/><\/a:stretch><\/a:blipFill><a:effectLst\/><\/p:bgPr><\/p:bg>/);
});
```

- [ ] **Step 2: Run them to see them fail**

Run: `node --test test/shapes.test.mjs test/writer.test.mjs`
Expected: the three new tests FAIL.

- [ ] **Step 3: Implement**

In `src/ooxml/shapes.mjs`:

1. Replace `xfrmXml` with:

```js
export function xfrmXml(box, tag = "a:xfrm") {
  // rot is in 60000ths of a degree, 0..21599999.
  const rot = box.rot ? ` rot="${Math.round((((box.rot % 360) + 360) % 360) * 60000) % 21600000}"` : "";
  return `<${tag}${rot}><a:off x="${emu(box.x)}" y="${emu(box.y)}"/><a:ext cx="${emuSize(box.w)}" cy="${emuSize(box.h)}"/></${tag}>`;
}
```

2. In `lineXml`, change `${solidFill(line.color)}` to `${solidFill(line.color, line.alpha)}`.

3. After the `SHADOW` constant add:

```js
// `true` = the deck's soft card shadow; an object = a measured CSS box-shadow.
export function shadowXml(shadow) {
  if (shadow === true) return SHADOW;
  const dir = Math.round(((((shadow.dir ?? 90) % 360) + 360) % 360) * 60000) % 21600000;
  return (
    `<a:effectLst><a:outerShdw blurRad="${emuSize(shadow.blur ?? 0)}" dist="${emuSize(shadow.dist ?? 0)}" dir="${dir}" algn="ctr" rotWithShape="0">` +
    `${colorXml(shadow.color, shadow.alpha)}</a:outerShdw></a:effectLst>`
  );
}
```

4. In `shapeXml`, replace the fill/line/shadow part of `spPr` so it reads:

```js
    `${shape.fill ? solidFill(shape.fill, shape.fillAlpha) : "<a:noFill/>"}${lineXml(shape.line)}${shape.shadow ? shadowXml(shape.shadow) : ""}</p:spPr>`;
```

5. Replace `pictureXml` with:

```js
export function pictureXml({ id, name, descr, rId, box, crop, line, shadow }) {
  const alt = descr ? ` descr="${esc(descr)}"` : "";
  // srcRect crops in 1/1000 of a percent of the image, per side.
  const sides = crop ? ["l", "t", "r", "b"].filter((k) => crop[k] > 0) : [];
  const srcRect = sides.length ? `<a:srcRect${sides.map((k) => ` ${k}="${Math.round(crop[k] * 100000)}"`).join("")}/>` : "";
  return (
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${esc(name)}"${alt}/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${rId}"/>${srcRect}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr>${xfrmXml(box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${line ? lineXml(line) : ""}${shadow ? shadowXml(shadow) : ""}</p:spPr></p:pic>`
  );
}
```

6. Replace `slideXml` with the following, and add `backgroundXml` above it:

```js
// Slide background (p:bg): a fill, or a picture (rId) stretched over the slide.
export function backgroundXml(bg, rId) {
  const fill = rId
    ? `<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="${rId}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`
    : solidFill(bg.color, bg.alpha);
  return `<p:bg><p:bgPr>${fill}<a:effectLst/></p:bgPr></p:bg>`;
}

export function slideXml(shapesXml, bgXml = "") {
  return (
    `${XML_DECL}<p:sld ${NS_PML}><p:cSld>${bgXml}<p:spTree>${SP_TREE_HEADER}${shapesXml}</p:spTree></p:cSld>` +
    "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
  );
}
```

In `src/ooxml/index.mjs`:

1. Change the shapes import to include `backgroundXml`:

```js
import { placeholderXml, shapeXml, pictureXml, slideXml, graphicFrameXml, backgroundXml } from "./shapes.mjs";
```

2. Replace `MediaStore.add` with:

```js
  // `source` is a file path, or { data: Buffer, ext } for bytes captured in render mode.
  add(source) {
    const fromFile = typeof source === "string";
    const buf = fromFile ? readFileSync(source) : source.data;
    const ext = (fromFile ? extname(source).slice(1) : source.ext).toLowerCase().replace("jpeg", "jpg");
    const key = `${createHash("sha1").update(buf).digest("hex")}.${ext}`;
    if (!this.parts.has(key)) {
      const part = `/ppt/media/image${this.parts.size + 1}.${ext}`;
      this.pkg.addMedia(part, buf);
      this.parts.set(key, part);
    }
    return this.parts.get(key);
  }
```

3. Replace the `case "image":` block in `shapeToXml` with:

```js
    case "image": {
      const rId = pkg.rel(part, media.add(shape.path ?? { data: shape.data, ext: shape.ext }), REL.image);
      return pictureXml({ id, name: `Picture ${id}`, descr: shape.alt, rId, box: shape, crop: shape.crop, line: shape.line, shadow: shape.shadow });
    }
```

4. In the slides loop, replace `pkg.addXml(part, CT.slide, slideXml(xml));` with:

```js
    let bg = "";
    if (spec.background?.image) bg = backgroundXml(spec.background, pkg.rel(part, media.add(spec.background.image), REL.image));
    else if (spec.background?.color) bg = backgroundXml(spec.background);
    pkg.addXml(part, CT.slide, slideXml(xml, bg));
```

- [ ] **Step 4: Run all tests**

Run: `npm test`
Expected: all pass.

- [ ] **Step 5: Commit**

```bash
git add src/ooxml/shapes.mjs src/ooxml/index.mjs test/shapes.test.mjs test/writer.test.mjs
git commit -m "Writer: alpha fills and outlines, measured shadows, rotation, picture crops, images from buffers, slide backgrounds"
```

---

### Task 3: `sceneToSpec` — Scene to SlideSpec

**Files:**
- Create: `src/render/scene.mjs`
- Test: `test/render-scene.test.mjs`

**Interfaces:**
- Consumes: `toColor(hex, theme)` (`src/colors.mjs`); Tasks 1–2 SlideSpec fields; `writePptx` for the lint test.
- Produces:
  - `sceneToSpec(scene, style) → SlideSpec`, where `style = { geometry, theme, fonts }`.
  - `unitsFor(scene, geometry) → { inch(px), pt(px), box(rect) }`.
  - `shapesOf(node, u, style) → Shape[]`, which Task 12 uses for master shapes.
- Scene contract (phase 1). Rects are `{ x, y, w, h }` in canvas px. Colors are `{ hex: "RRGGBB", alpha }`. Every node carries `opacity` (cumulative) and an optional `rotate` (degrees).
  - `{ index, sourceIndex, layout, canvas: { w, h }, notes, background: { color } | { raster: true, png? }, nodes }`
  - `{ kind: "text", rect, role: "title"|"body", align, lineHeight, singleLine, runs }`
  - `{ kind: "list", rect, items: [{ level, marker: "bullet"|"number"|"none", start?, markerColor, indent, spaceBefore, lineHeight, align, runs }] }`
  - `{ kind: "box", rect, fill, border: { width, color } | null, sides: [{ side: "t"|"r"|"b"|"l", width, color }] | null, radius, shadow: { x, y, blur, color } | null }`
  - `{ kind: "image", id, rect, src, format, fit, natural: { w, h } | null, alt, border, shadow, data?, ext? }`. Capture adds `data` and `ext`.
  - `{ kind: "raster", id, rect, reason, label, alt, png? }`. Capture adds `png`.
  - `{ kind: "missing", label, src }`
  - A Run is `{ text, font, size, bold, italic, color, spacing, underline, strike, highlight }`, with size and spacing in px.

- [ ] **Step 1: Write the failing tests** — `test/render-scene.test.mjs`

```js
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { sceneToSpec } from "../src/render/scene.mjs";
import { slideGeometry } from "../src/geometry.mjs";
import { themeColors } from "../src/colors.mjs";
import { writePptx } from "../src/ooxml/index.mjs";
import { DECK, PIC } from "./deck.mjs";
import { lintPackage } from "./lint.mjs";

const STYLE = {
  geometry: slideGeometry({ width: 13.333, height: 7.5, marginX: 0.6 }),
  theme: themeColors({ accent: "1A833B", accentDark: "14532D", accent2: "1A589F", panel: "EEF4F0", border: "D9E3DB", text: "1F2937", muted: "6B7280" }),
  fonts: { heading: "Plus Jakarta Sans", body: "Inter", mono: "JetBrains Mono" },
};
const K = 13.333 / 980;
const inch = (px) => Math.round(px * K * 10000) / 10000;
const pt = (px) => Math.round(px * K * 72 * 100) / 100;
const box = (r) => ({ x: inch(r.x), y: inch(r.y), w: inch(r.w), h: inch(r.h) });
const C = (hex, alpha = 1) => ({ hex, alpha });
const run = (text, extra = {}) => ({ text, font: "Inter", size: 17.6, bold: false, italic: false, color: C("1F2937"), spacing: 0, underline: false, strike: false, highlight: null, ...extra });
const text = (runs, extra = {}) => ({ kind: "text", rect: { x: 56, y: 300, w: 400, h: 24 }, role: "body", align: "left", lineHeight: 24, singleLine: true, runs, opacity: 1, ...extra });
const scene = (nodes, extra = {}) => ({ index: 3, sourceIndex: 2, layout: "default", canvas: { w: 980, h: 552 }, notes: null, background: { color: C("FFFFFF") }, nodes, ...extra });

test("the first title node fills the title placeholder at its measured box; the rest become shapes", () => {
  const titleRect = { x: 54, y: 163, w: 870, h: 40 };
  const spec = sceneToSpec(scene([
    text([run("Overview", { font: "Plus Jakarta Sans", size: 12, bold: true, color: C("1A833B"), spacing: 1.68 })], { rect: { x: 56, y: 137, w: 90, h: 18 }, lineHeight: 18 }),
    text([run("Card row", { font: "Plus Jakarta Sans", size: 36, bold: true })], { role: "title", rect: titleRect, lineHeight: 40 }),
  ]), STYLE);
  assert.equal(spec.layout, "titleOnly");
  // A one-line title gets room to spare up to the mirrored right margin (980 - 2 x 54 = 872).
  assert.deepEqual(spec.placeholders.titleBox, box({ ...titleRect, w: 872 }));
  assert.deepEqual(spec.placeholders.titleBody, { wrap: false, inset: 0, anchor: "t", autofit: "none" });
  assert.deepEqual(spec.placeholders.title, [{ lineSpacing: pt(40), runs: [{ text: "Card row", size: pt(36), font: "heading", b: true, color: "dk1" }] }]);
  assert.equal(spec.shapes.length, 1);
  const [eyebrow] = spec.shapes;
  assert.deepEqual(eyebrow, {
    kind: "text", ...box({ x: 56, y: 137, w: 90, h: 18 }), geometry: "rect", anchor: "t", inset: 0, wrap: false,
    paras: [{ lineSpacing: pt(18), runs: [{ text: "Overview", size: pt(12), font: "heading", b: true, spacing: pt(1.68), color: "accent1" }] }],
  });
});

test("a centered one-line title widens symmetrically; a wrapped title keeps its measured box", () => {
  const centered = sceneToSpec(scene([text([run("Thanks", { size: 36 })], { role: "title", align: "center", rect: { x: 440, y: 300, w: 100, h: 40 } })]), STYLE);
  assert.deepEqual(centered.placeholders.titleBox, box({ x: 0, y: 300, w: 980, h: 40 }));
  const wrapped = sceneToSpec(scene([text([run("A long title", { size: 36 })], { role: "title", singleLine: false, rect: { x: 56, y: 100, w: 300, h: 80 } })]), STYLE);
  assert.deepEqual(wrapped.placeholders.titleBox, box({ x: 56, y: 100, w: 300, h: 80 }));
  assert.deepEqual(wrapped.placeholders.titleBody, { inset: 0, anchor: "t", autofit: "none" });
});

test("multi-line text wraps at its measured width; unknown fonts keep their name; code keeps its highlight; alignment maps", () => {
  const spec = sceneToSpec(scene([text([run("a "), run("code", { font: "Fira Code", highlight: C("F5F5F5") })], { singleLine: false, align: "center" })]), STYLE);
  const [shape] = spec.shapes;
  assert.equal(shape.wrap, undefined);
  assert.equal(shape.paras[0].align, "ctr");
  assert.deepEqual(shape.paras[0].runs[1], { text: "code", size: pt(17.6), typeface: "Fira Code", color: "dk1", highlight: "#F5F5F5" });
});

test("boxes: a rounded card with theme fill and outline; side-only borders become bars; shadow and opacity", () => {
  const card = { kind: "box", rect: { x: 56, y: 235, w: 276, h: 180 }, fill: C("EEF4F0"), border: { width: 1, color: C("D9E3DB") }, sides: null, radius: 8, shadow: null, opacity: 1 };
  const quote = { kind: "box", rect: { x: 60, y: 300, w: 400, h: 50 }, fill: null, border: null, sides: [{ side: "l", width: 4, color: C("1A833B") }], radius: 0, shadow: null, opacity: 1 };
  const lifted = { kind: "box", rect: { x: 0, y: 0, w: 100, h: 100 }, fill: C("FFFFFF"), border: null, sides: null, radius: 0, shadow: { x: 0, y: 12, blur: 32, color: C("14181A", 0.14) }, opacity: 0.5 };
  const [cardShape, bar, liftedShape] = sceneToSpec(scene([card, quote, lifted]), STYLE).shapes;
  assert.deepEqual(cardShape, { kind: "text", ...box(card.rect), geometry: "roundRect", paras: [], radius: inch(8), fill: "lt2", line: { width: pt(1), color: "accent4" } });
  assert.deepEqual(bar, { kind: "text", ...box({ x: 60, y: 300, w: 4, h: 50 }), geometry: "rect", fill: "accent1", paras: [] });
  assert.equal(liftedShape.fill, "lt1");
  assert.equal(liftedShape.fillAlpha, 0.5);
  assert.deepEqual(liftedShape.shadow, { blur: inch(32), dist: inch(12), dir: 90, color: "#14181A", alpha: 0.07 });
});

test("lists become one text box with native bullets, marker colors, measured indents and spacing", () => {
  const list = {
    kind: "list", rect: { x: 56, y: 200, w: 500, h: 80 }, opacity: 1,
    items: [
      { level: 0, marker: "bullet", markerColor: C("1A833B"), indent: 24, spaceBefore: 0, lineHeight: 26, align: "left", runs: [run("One")] },
      { level: 1, marker: "number", start: 3, markerColor: C("1F2937"), indent: 48, spaceBefore: 6, lineHeight: 26, align: "left", runs: [run("Three")] },
      { level: 0, marker: "none", markerColor: null, indent: 0, spaceBefore: 0, lineHeight: 26, align: "left", runs: [run("Plain")] },
    ],
  };
  const [shape] = sceneToSpec(scene([list]), STYLE).shapes;
  assert.equal(shape.wrap, undefined);
  assert.deepEqual(shape.paras.map(({ runs, ...p }) => p), [
    { level: 0, bullet: "char", bulletColor: "accent1", marL: inch(24), indent: -0.25, lineSpacing: pt(26) },
    { level: 1, bullet: "number", number: 3, bulletColor: "dk1", marL: inch(48), indent: -0.25, lineSpacing: pt(26), spaceBefore: pt(6) },
    { level: 0, bullet: "none", marL: 0, indent: 0, lineSpacing: pt(26) },
  ]);
});

test("images are contain-fitted or cover-cropped from their natural size; pictures of elements become images; nodes without bytes are dropped", () => {
  const data = Buffer.from("png");
  const img = (fit) => ({ kind: "image", id: 1, rect: { x: 0, y: 0, w: 200, h: 100 }, src: "/a.png", format: "png", fit, natural: { w: 100, h: 100 }, alt: "A", border: null, shadow: null, opacity: 1, data, ext: "png" });
  const [contain, cover, fill] = sceneToSpec(scene([img("contain"), img("cover"), img("fill")]), STYLE).shapes;
  assert.deepEqual(contain, { kind: "image", ...box({ x: 50, y: 0, w: 100, h: 100 }), data, ext: "png", alt: "A" });
  assert.deepEqual(cover.crop, { l: 0, t: 0.25, r: 0, b: 0.25 });
  assert.deepEqual({ x: cover.x, w: cover.w }, { x: 0, w: inch(200) });
  assert.equal(fill.crop, undefined);
  const raster = { kind: "raster", id: 2, rect: { x: 10, y: 10, w: 50, h: 50 }, reason: "inline SVG", label: "svg", alt: "", png: data };
  const shapes = sceneToSpec(scene([raster, { ...raster, png: undefined }, { ...img("fill"), data: undefined }, { kind: "missing", label: "img", src: "/x.png" }]), STYLE).shapes;
  assert.deepEqual(shapes, [{ kind: "image", ...box(raster.rect), data, ext: "png", alt: "" }]);
});

test("layouts: the deck's first slide and cover/intro slides are Title Slides; untitled slides are Blank", () => {
  assert.equal(sceneToSpec(scene([], { index: 1 }), STYLE).layout, "title");
  assert.equal(sceneToSpec(scene([], { index: 4, layout: "intro" }), STYLE).layout, "title");
  assert.equal(sceneToSpec(scene([]), STYLE).layout, "blank");
});

test("backgrounds: non-white colors and pictures become the slide background; notes are carried", () => {
  assert.equal(sceneToSpec(scene([]), STYLE).background, undefined);
  assert.deepEqual(sceneToSpec(scene([], { background: { color: C("F8FAFC") } }), STYLE).background, { color: "#F8FAFC" });
  const png = Buffer.from("bg");
  assert.deepEqual(sceneToSpec(scene([], { background: { raster: true, png } }), STYLE).background, { image: { data: png, ext: "png" } });
  assert.equal(sceneToSpec(scene([], { notes: "Say this" }), STYLE).notes, "Say this");
});

test("a scene with every node kind writes a lint-clean package", async () => {
  const data = readFileSync(PIC);
  const spec = sceneToSpec(scene([
    text([run("Title & <more>", { size: 36 })], { role: "title" }),
    { kind: "box", rect: { x: 56, y: 235, w: 276, h: 180 }, fill: C("EEF4F0"), border: { width: 1, color: C("D9E3DB") }, sides: [{ side: "b", width: 2, color: C("1A833B") }], radius: 8, shadow: { x: 0, y: 4, blur: 8, color: C("000000", 0.2) }, opacity: 0.9, rotate: 5 },
    { kind: "list", rect: { x: 56, y: 100, w: 500, h: 60 }, opacity: 1, items: [{ level: 0, marker: "bullet", markerColor: C("1A833B"), indent: 24, spaceBefore: 0, lineHeight: 26, align: "left", runs: [run("Item")] }] },
    { kind: "image", id: 1, rect: { x: 600, y: 100, w: 100, h: 50 }, src: "/p.png", format: "png", fit: "cover", natural: { w: 4, h: 2 }, alt: "pic", border: { width: 1, color: C("D9E3DB") }, shadow: null, opacity: 1, data, ext: "png" },
    { kind: "raster", id: 2, rect: { x: 700, y: 300, w: 40, h: 40 }, reason: "canvas", label: "canvas", alt: "", png: data },
  ], { background: { color: C("F8FAFC") }, notes: "n" }), STYLE);
  const buffer = await writePptx([spec], { ...DECK, logo: null });
  assert.deepEqual(await lintPackage(buffer), []);
});
```

- [ ] **Step 2: Run them to see them fail**

Run: `node --test test/render-scene.test.mjs`
Expected: FAIL — `Cannot find module '.../src/render/scene.mjs'`.

- [ ] **Step 3: Implement** — `src/render/scene.mjs`

```js
// Render mode: Scene (what extract.js measured in the browser, in canvas px) -> SlideSpec (what
// the OOXML writer draws, in inches and points). Pure - no browser, no file access - so every
// mapping rule is unit-tested with hand-written scenes.

import { toColor } from "../colors.mjs";

const ALIGN = { left: "l", start: "l", center: "ctr", right: "r", end: "r", justify: "just" };

const round = (v, digits = 4) => Math.round(v * 10 ** digits) / 10 ** digits;

// 1 canvas px = k inches, where k = slide width / canvas width.
export function unitsFor(scene, geometry) {
  const k = geometry.width / scene.canvas.w;
  return {
    inch: (px) => round(px * k),
    pt: (px) => round(px * k * 72, 2),
    box: (r) => ({ x: round(r.x * k), y: round(r.y * k), w: round(r.w * k), h: round(r.h * k) }),
  };
}

function colorOf(c, style) {
  return toColor(c.hex, style.theme);
}

// The alpha to write after the node's cumulative opacity; undefined when fully opaque.
function alphaOf(c, opacity = 1) {
  const a = round((c.alpha ?? 1) * opacity, 3);
  return a < 1 ? a : undefined;
}

// Fonts that are the theme's heading/body/mono fonts become theme font references.
function fontOf(family, fonts) {
  const f = String(family || "").toLowerCase();
  for (const role of ["heading", "body", "mono"]) {
    if (fonts[role] && fonts[role].toLowerCase() === f) return { font: role };
  }
  return family ? { typeface: family } : { font: "body" };
}

function runOf(r, u, style, opacity) {
  const run = { text: r.text, size: u.pt(r.size), ...fontOf(r.font, style.fonts) };
  if (r.bold) run.b = true;
  if (r.italic) run.i = true;
  if (r.underline) run.u = true;
  if (r.strike) run.strike = true;
  if (r.spacing) run.spacing = u.pt(r.spacing);
  if (r.color) {
    run.color = colorOf(r.color, style);
    const a = alphaOf(r.color, opacity);
    if (a !== undefined) run.alpha = a;
  }
  if (r.highlight) run.highlight = colorOf(r.highlight, style);
  return run;
}

function textShape(node, u, style) {
  const para = { runs: node.runs.map((r) => runOf(r, u, style, node.opacity ?? 1)) };
  const align = ALIGN[node.align];
  if (align && align !== "l") para.align = align;
  if (node.lineHeight) para.lineSpacing = u.pt(node.lineHeight);
  const shape = { kind: "text", ...u.box(node.rect), geometry: "rect", anchor: "t", inset: 0, paras: [para] };
  if (node.singleLine) shape.wrap = false;
  if (node.rotate) shape.rot = node.rotate;
  return shape;
}

function listShape(node, u, style) {
  const paras = node.items.map((it) => {
    const para = { level: it.level };
    para.bullet = it.marker === "none" ? "none" : it.marker === "number" ? "number" : "char";
    if (para.bullet === "number" && it.start != null) para.number = it.start;
    if (para.bullet !== "none" && it.markerColor) para.bulletColor = colorOf(it.markerColor, style);
    para.marL = u.inch(it.indent);
    para.indent = para.bullet === "none" ? 0 : -Math.min(para.marL, 0.25);
    if (it.lineHeight) para.lineSpacing = u.pt(it.lineHeight);
    if (it.spaceBefore > 0) para.spaceBefore = u.pt(it.spaceBefore);
    const align = ALIGN[it.align];
    if (align && align !== "l") para.align = align;
    para.runs = it.runs.map((r) => runOf(r, u, style, node.opacity ?? 1));
    return para;
  });
  return { kind: "text", ...u.box(node.rect), geometry: "rect", anchor: "t", inset: 0, paras };
}

function shadowOf(s, u, style, opacity) {
  const shadow = {
    blur: u.inch(s.blur),
    dist: u.inch(Math.hypot(s.x, s.y)),
    dir: round(((Math.atan2(s.y, s.x) * 180) / Math.PI + 360) % 360, 2),
    color: colorOf(s.color, style),
  };
  const a = alphaOf(s.color, opacity);
  if (a !== undefined) shadow.alpha = a;
  return shadow;
}

// A border on only some sides (a blockquote's left rule, a heading's underline) is a thin bar.
function sideBar(rect, side, u, style, opacity) {
  const t = side.width;
  const r =
    side.side === "l" ? { x: rect.x, y: rect.y, w: t, h: rect.h }
    : side.side === "r" ? { x: rect.x + rect.w - t, y: rect.y, w: t, h: rect.h }
    : side.side === "t" ? { x: rect.x, y: rect.y, w: rect.w, h: t }
    : { x: rect.x, y: rect.y + rect.h - t, w: rect.w, h: t };
  const bar = { kind: "text", ...u.box(r), geometry: "rect", fill: colorOf(side.color, style), paras: [] };
  const a = alphaOf(side.color, opacity);
  if (a !== undefined) bar.fillAlpha = a;
  return bar;
}

function boxShapes(node, u, style) {
  const op = node.opacity ?? 1;
  const shapes = [];
  if (node.fill || node.border || node.shadow) {
    const shape = { kind: "text", ...u.box(node.rect), geometry: node.radius > 0 ? "roundRect" : "rect", paras: [] };
    if (node.radius > 0) shape.radius = u.inch(node.radius);
    if (node.fill) {
      shape.fill = colorOf(node.fill, style);
      const a = alphaOf(node.fill, op);
      if (a !== undefined) shape.fillAlpha = a;
    }
    if (node.border) {
      shape.line = { width: u.pt(node.border.width), color: colorOf(node.border.color, style) };
      const a = alphaOf(node.border.color, op);
      if (a !== undefined) shape.line.alpha = a;
    }
    if (node.shadow) shape.shadow = shadowOf(node.shadow, u, style, op);
    if (node.rotate) shape.rot = node.rotate;
    shapes.push(shape);
  }
  for (const side of node.sides || []) shapes.push(sideBar(node.rect, side, u, style, op));
  return shapes;
}

// object-fit: the picture's box and crop for an image of `natural` size drawn into `rect`.
export function fitImage(rect, natural, fit) {
  if (!natural || !natural.w || !natural.h || fit === "fill") return { rect };
  if (fit === "cover") {
    const s = Math.max(rect.w / natural.w, rect.h / natural.h);
    const fx = Math.min(1, rect.w / (natural.w * s));
    const fy = Math.min(1, rect.h / (natural.h * s));
    const crop = { l: round((1 - fx) / 2), t: round((1 - fy) / 2), r: round((1 - fx) / 2), b: round((1 - fy) / 2) };
    return { rect, crop: crop.l || crop.t ? crop : undefined };
  }
  // contain, scale-down and none are all drawn letterboxed inside the box.
  let s = Math.min(rect.w / natural.w, rect.h / natural.h);
  if (fit === "scale-down" || fit === "none") s = Math.min(1, s);
  const w = natural.w * s;
  const h = natural.h * s;
  return { rect: { x: rect.x + (rect.w - w) / 2, y: rect.y + (rect.h - h) / 2, w, h } };
}

function imageShape(node, u, style) {
  const { rect, crop } = fitImage(node.rect, node.natural, node.fit);
  const shape = { kind: "image", ...u.box(rect), data: node.data, ext: node.ext, alt: node.alt || "" };
  if (crop) shape.crop = crop;
  const op = node.opacity ?? 1;
  if (node.border) shape.line = { width: u.pt(node.border.width), color: colorOf(node.border.color, style) };
  if (node.shadow) shape.shadow = shadowOf(node.shadow, u, style, op);
  if (node.rotate) shape.rot = node.rotate;
  return shape;
}

export function shapesOf(node, u, style) {
  switch (node.kind) {
    case "text":
      return [textShape(node, u, style)];
    case "list":
      return [listShape(node, u, style)];
    case "box":
      return boxShapes(node, u, style);
    case "image":
      return node.data ? [imageShape(node, u, style)] : [];
    case "raster":
      return node.png ? [{ kind: "image", ...u.box(node.rect), data: node.png, ext: "png", alt: node.alt || "" }] : [];
    case "missing":
      return [];
    default:
      throw new Error(`Unknown scene node kind ${JSON.stringify(node.kind)}`);
  }
}

function backgroundOf(bg, style) {
  if (!bg) return null;
  if (bg.png) return { image: { data: bg.png, ext: "png" } };
  if (!bg.color) return null;
  const color = colorOf(bg.color, style);
  const alpha = alphaOf(bg.color);
  // Plain white is the master's background already.
  if (color === "lt1" && alpha === undefined) return null;
  return alpha === undefined ? { color } : { color, alpha };
}

// A title that fits on one line in the browser gets room to spare, so a wider substitute font
// can't wrap it (LibreOffice wraps placeholder text even with wrap="none"): left-aligned titles
// extend to the mirrored right margin, centered ones widen symmetrically, right-aligned ones
// extend to the mirrored left margin.
function titleRectOf(node, canvasW) {
  const r = node.rect;
  if (!node.singleLine) return r;
  if (node.align === "center") {
    const c = r.x + r.w / 2;
    const half = Math.max(r.w / 2, Math.min(c, canvasW - c));
    return { ...r, x: c - half, w: 2 * half };
  }
  if (node.align === "right" || node.align === "end") {
    const right = r.x + r.w;
    const x = Math.min(r.x, canvasW - right);
    return { ...r, x, w: right - x };
  }
  return { ...r, w: Math.max(r.w, canvasW - 2 * r.x) };
}

export function sceneToSpec(scene, style) {
  const u = unitsFor(scene, style.geometry);
  const spec = { layout: "blank", placeholders: {}, shapes: [] };
  if (scene.notes) spec.notes = scene.notes;
  const bg = backgroundOf(scene.background, style);
  if (bg) spec.background = bg;
  let title = null;
  for (const node of scene.nodes) {
    if (!title && node.kind === "text" && node.role === "title") {
      title = node;
      continue;
    }
    spec.shapes.push(...shapesOf(node, u, style));
  }
  if (title) {
    spec.placeholders.title = textShape(title, u, style).paras;
    spec.placeholders.titleBox = u.box(titleRectOf(title, scene.canvas.w));
    // The box is measured, so none of the master's insets, bottom anchoring or shrink-to-fit
    // apply; a title that fits on one line in the browser must not wrap in PowerPoint.
    const body = { inset: 0, anchor: "t", autofit: "none" };
    spec.placeholders.titleBody = title.singleLine ? { wrap: false, ...body } : body;
  }
  const cover = scene.index === 1 || scene.layout === "cover" || scene.layout === "intro";
  spec.layout = cover ? "title" : title ? "titleOnly" : "blank";
  return spec;
}
```

- [ ] **Step 4: Run all tests**

Run: `npm test`
Expected: all pass.

- [ ] **Step 5: Commit**

```bash
git add src/render/scene.mjs test/render-scene.test.mjs
git commit -m "Render mode: map measured scenes onto slide specs (text, lists, boxes, images, backgrounds, titles)"
```

---
### Task 4: In-page extractor and the browser test harness

**Files:**
- Create: `src/render/extract.js`, `src/render/browser.mjs`, `test/render/helpers.mjs`, `test/render/extract.test.mjs`
- Modify: `package.json` (devDependency, optional peer dependencies, `test:render` script)

**Interfaces:**
- Produces:
  - `window.__pptxExtract(pageEl, canvas)` returns a Scene without `index`/`sourceIndex`/`notes`: `{ canvas, layout, background, nodes }`, using the node contract from Task 3. Every `image` and `raster` node's element gets the attribute `data-pptx-node="<id>"`.
  - `loadChromium(deckDir, { resolve }?) → Promise<chromium>` (`src/render/browser.mjs`).
  - Test helpers `usableChromium()`, `extractHtml(browser, html, opts)`, `EXTRACT`, `ROOT`.
- In phase 1 these elements become `raster` nodes: `canvas`, inline `svg`, `iframe`, `video`, `audio`, `object`/`embed`, `math`, `table`, `pre`, SVG/unknown-format/translucent `img`, CSS filter/backdrop-filter/mask/clip-path/blend modes, `background-clip: text`, gradients and layered background images, non-rotation transforms, and rotated elements with child elements.

- [ ] **Step 1: Install Playwright for development and add the scripts**

Run: `npm install -D playwright-chromium` (downloads Chromium for this Playwright version if it isn't cached).

Then edit `package.json` so `scripts` and the new peer-dependency fields read:

```json
  "scripts": {
    "test": "node --test test/*.test.mjs",
    "test:render": "node --test test/render/*.test.mjs",
    "render": "node scripts/render.mjs"
  },
  "peerDependencies": {
    "playwright": ">=1.40.0",
    "playwright-chromium": ">=1.40.0"
  },
  "peerDependenciesMeta": {
    "playwright": { "optional": true },
    "playwright-chromium": { "optional": true }
  },
```

(keep `dependencies`; `devDependencies` now also lists `playwright-chromium`).

- [ ] **Step 2: Write `src/render/browser.mjs`**

```js
// Render mode drives Chromium through Playwright. Playwright is an optional peer dependency -
// installing slidev-pptx-export never downloads a browser - so it's resolved at runtime: the
// deck's own node_modules first (where `slidev export` users already have playwright-chromium),
// then this package's.

import { createRequire } from "node:module";
import { join } from "node:path";
import { pathToFileURL, fileURLToPath } from "node:url";

const PACKAGE_ROOT = fileURLToPath(new URL("../../", import.meta.url));

function defaultResolve(name, base) {
  try {
    return createRequire(join(base, "package.json")).resolve(name);
  } catch {
    return null;
  }
}

export async function loadChromium(deckDir, { resolve = defaultResolve } = {}) {
  for (const name of ["playwright-chromium", "playwright"]) {
    for (const base of [deckDir, PACKAGE_ROOT]) {
      const path = resolve(name, base);
      if (!path) continue;
      const mod = await import(pathToFileURL(path).href);
      const chromium = mod.chromium ?? mod.default?.chromium;
      if (chromium) return chromium;
    }
  }
  throw new Error(
    "render mode needs Playwright's Chromium: run `npm i -D playwright-chromium` (the package `slidev export` uses), then `npx playwright install chromium`.",
  );
}
```

- [ ] **Step 3: Write the test helpers** — `test/render/helpers.mjs`

```js
// Shared setup for the browser tests (npm run test:render). Tests skip - never fail - when
// Playwright or its Chromium isn't installed, so this suite is safe to run anywhere.
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { loadChromium } from "../../src/render/browser.mjs";

export const ROOT = fileURLToPath(new URL("../../", import.meta.url));
export const EXTRACT = readFileSync(new URL("../../src/render/extract.js", import.meta.url), "utf8");

// Playwright's chromium if it loads and can launch here, else null.
export async function usableChromium() {
  try {
    const chromium = await loadChromium(ROOT);
    const browser = await chromium.launch();
    await browser.close();
    return chromium;
  } catch {
    return null;
  }
}

// Renders `html` inside a 980x552 `.slidev-page` like one Slidev slide and returns its Scene.
// `setup` (optional) runs in the page before extraction.
export async function extractHtml(browser, html, { layoutClass = "slidev-layout default", layoutStyle = "", css = "", setup = null } = {}) {
  const page = await browser.newPage({ viewport: { width: 980, height: 552 } });
  try {
    await page.setContent(
      `<!doctype html><html><head><style>body{margin:0;font-family:Arial,sans-serif}${css}</style></head><body>` +
        '<div class="slidev-page slidev-page-1" style="position:relative;width:980px;height:552px;overflow:hidden">' +
        `<div class="${layoutClass}" style="width:100%;height:100%;box-sizing:border-box;padding:40px;${layoutStyle}">${html}</div></div></body></html>`,
    );
    if (setup) await page.evaluate(setup);
    await page.addScriptTag({ content: EXTRACT });
    return await page.evaluate(() => window.__pptxExtract(document.querySelector(".slidev-page"), { w: 980, h: 552 }));
  } finally {
    await page.close();
  }
}
```

- [ ] **Step 4: Write the failing tests** — `test/render/extract.test.mjs`

```js
import { test, after } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { usableChromium, extractHtml } from "./helpers.mjs";

const chromium = await usableChromium();
const browser = chromium ? await chromium.launch() : null;
after(() => browser?.close());
const opts = { skip: browser ? false : "Playwright/Chromium not available - run npm i -D playwright-chromium" };
const PNG = `data:image/png;base64,${readFileSync(new URL("../fixtures/kitchen-sink/public/pic.png", import.meta.url)).toString("base64")}`;
const textOf = (node) => node.runs.map((r) => r.text).join("");

test("a card becomes a rounded box with fill and border; its heading and paragraph become text nodes", opts, async () => {
  const s = await extractHtml(browser,
    '<div style="width:276px;padding:20px;border-radius:8px;border:1px solid #D9E3DB;background:#EEF4F0">' +
      '<h3 style="margin:0;font:24px Arial">First point</h3>' +
      '<p style="margin:8px 0 0;font:17.6px Arial;color:#1f2937">Each card is a <code style="background:#f5f5f5;font-family:monospace">kv-card</code> inside a grid row of three cards.</p></div>');
  const box = s.nodes.find((n) => n.kind === "box");
  assert.deepEqual({ fill: box.fill, border: box.border, radius: box.radius, x: box.rect.x, y: box.rect.y, w: box.rect.w }, {
    fill: { hex: "EEF4F0", alpha: 1 }, border: { width: 1, color: { hex: "D9E3DB", alpha: 1 } }, radius: 8, x: 40, y: 40, w: 318,
  });
  const texts = s.nodes.filter((n) => n.kind === "text");
  assert.deepEqual(texts.map(textOf), ["First point", "Each card is a kv-card inside a grid row of three cards."]);
  assert.ok(s.nodes.indexOf(box) < s.nodes.indexOf(texts[0]), "the box is painted before its contents");
  assert.equal(texts[0].singleLine, true);
  assert.ok(texts[0].rect.w < 276, "a single line hugs its text");
  assert.equal(texts[1].singleLine, false);
  assert.deepEqual({ x: texts[1].rect.x, w: texts[1].rect.w }, { x: 61, w: 276 }, "wrapped text keeps its content box width");
  const code = texts[1].runs.find((r) => r.text === "kv-card");
  assert.deepEqual(code.highlight, { hex: "F5F5F5", alpha: 1 });
  assert.equal(code.font, "monospace");
});

test("an eyebrow keeps uppercase, letter spacing and weight in its run", opts, async () => {
  const s = await extractHtml(browser, '<div style="font:800 12px Arial;letter-spacing:1.68px;text-transform:uppercase;color:#1a833b">Overview</div>');
  assert.deepEqual(s.nodes.map((n) => n.kind), ["text"]);
  assert.deepEqual(s.nodes[0].runs, [{ text: "OVERVIEW", font: "Arial", size: 12, bold: true, italic: false, color: { hex: "1A833B", alpha: 1 }, spacing: 1.68, underline: false, strike: false, highlight: null }]);
});

test("the first h1 is the title; other headings are body text", opts, async () => {
  const s = await extractHtml(browser, '<h1 style="margin:0">Card <em>row</em></h1><h2>Sub</h2><h1>Second</h1>');
  assert.deepEqual(s.nodes.map((n) => [n.role, textOf(n)]), [["title", "Card row"], ["body", "Sub"], ["body", "Second"]]);
  assert.equal(s.nodes[0].runs[1].italic, true);
});

test("lists: items, nesting levels, ordered start, marker color and measured indent", opts, async () => {
  const s = await extractHtml(browser,
    '<ul style="margin:0;padding-left:24px;font:16px/24px Arial"><li>One</li><li>Two<ol start="3" style="padding-left:24px"><li>Three</li></ol></li></ul>',
    { css: "li::marker{color:#1a833b}" });
  const list = s.nodes.find((n) => n.kind === "list");
  assert.deepEqual(list.items.map((i) => [i.level, i.marker, i.start ?? null, textOf(i), i.markerColor.hex, i.indent, i.lineHeight]), [
    [0, "bullet", null, "One", "1A833B", 24, 24],
    [0, "bullet", null, "Two", "1A833B", 24, 24],
    [1, "number", 3, "Three", "1A833B", 48, 24],
  ]);
  assert.equal(s.nodes.filter((n) => n.kind === "text").length, 0, "list text lives only in the list node");
});

test("svg, canvas and video become raster nodes tagged for screenshots, without walking their contents", opts, async () => {
  const s = await extractHtml(browser,
    '<svg width="50" height="50"><text x="0" y="20">svg text</text></svg><canvas width="40" height="40" style="display:block"></canvas><video width="40" height="30" style="display:block"></video>');
  assert.deepEqual(s.nodes.map((n) => [n.kind, n.reason, n.label]), [["raster", "inline SVG", "svg"], ["raster", "canvas", "canvas"], ["raster", "video", "video"]]);
  assert.deepEqual(s.nodes.map((n) => n.id), [0, 1, 2]);
});

test("images: content box, natural size, object-fit and border; a broken image is reported missing", opts, async () => {
  const s = await extractHtml(browser,
    `<img src="${PNG}" alt="Pic" style="display:block;width:100px;height:100px;object-fit:cover;padding:5px;border:1px solid #000">` +
      '<img src="http://127.0.0.1:9/nope.png" alt="x" style="width:20px;height:20px">');
  const [img, missing] = s.nodes;
  assert.deepEqual({ kind: img.kind, rect: img.rect, natural: img.natural, fit: img.fit, format: img.format, alt: img.alt, border: img.border }, {
    kind: "image", rect: { x: 46, y: 46, w: 100, h: 100 }, natural: { w: 4, h: 2 }, fit: "cover", format: "png", alt: "Pic", border: { width: 1, color: { hex: "000000", alpha: 1 } },
  });
  assert.equal(missing.kind, "missing");
  assert.equal(missing.src, "http://127.0.0.1:9/nope.png");
});

test("hidden, transparent and zero-size elements are skipped; opacity multiplies down; one-sided borders are sides", opts, async () => {
  const s = await extractHtml(browser,
    '<div style="display:none">a</div><div style="visibility:hidden">b</div><div style="opacity:0">c</div><div style="width:0;height:0;background:red"></div>' +
      '<div style="opacity:0.5"><div style="background:#000;opacity:0.5;width:10px;height:10px"></div></div>' +
      '<blockquote style="margin:0;border-left:4px solid #1a833b;padding-left:8px">quoted</blockquote>');
  assert.deepEqual(s.nodes.filter((n) => n.kind === "text").map(textOf), ["quoted"]);
  const [faded, quote] = s.nodes.filter((n) => n.kind === "box");
  assert.equal(faded.opacity, 0.25);
  assert.deepEqual({ fill: quote.fill, border: quote.border, sides: quote.sides }, { fill: null, border: null, sides: [{ side: "l", width: 4, color: { hex: "1A833B", alpha: 1 } }] });
});

test("a rotated text-only element keeps its own box and a rotation; skewed elements are rasters", opts, async () => {
  const s = await extractHtml(browser, '<div style="width:100px;transform:rotate(90deg)">Turned</div><div style="transform:skewX(10deg)">Skewed</div>');
  const turned = s.nodes.find((n) => n.kind === "text");
  assert.equal(turned.rotate, 90);
  assert.equal(turned.rect.w, 100);
  assert.deepEqual(s.nodes.filter((n) => n.kind === "raster").map((n) => n.reason), ["transform"]);
});

test("backgrounds: the layout root's color is the slide background; a gradient asks for a picture background", opts, async () => {
  assert.deepEqual((await extractHtml(browser, "<p>x</p>", { layoutStyle: "background:#f8fafc" })).background, { color: { hex: "F8FAFC", alpha: 1 } });
  const g = await extractHtml(browser, "<p>x</p>", { layoutStyle: "background:radial-gradient(#fff, #000)" });
  assert.deepEqual(g.background, { raster: true, reason: "gradient background" });
  assert.equal(g.layout, "default");
});

test("a custom layout without the slidev-layout class still extracts, on white", opts, async () => {
  const s = await extractHtml(browser, '<h1 style="margin:0">Custom</h1>', { layoutClass: "my-layout" });
  assert.equal(s.layout, "default");
  assert.deepEqual(s.background, { color: { hex: "FFFFFF", alpha: 1 } });
  assert.deepEqual(s.nodes.map((n) => [n.kind, n.role]), [["text", "title"]]);
});
```

- [ ] **Step 5: Run them to see them fail**

Run: `npm run test:render`
Expected: the extract tests FAIL. The page has no `window.__pptxExtract`, because `src/render/extract.js` doesn't exist yet: `helpers.mjs` fails to read it, so the whole file errors. If instead every test reports *skipped*, Chromium can't launch. Run `npx playwright install chromium` and retry; this task needs the tests to run.

- [ ] **Step 6: Implement** — `src/render/extract.js`

```js
// Render mode's in-page extractor. capture.mjs loads this file into a slide's page with
// page.addScriptTag and calls window.__pptxExtract(pageEl, canvas). It walks the rendered DOM
// under one `.slidev-page` and returns a Scene: plain JSON describing every visible element -
// its measured box in canvas px (relative to the slide's top-left) and its computed style.
// A classic script with no imports, so it runs as-is in Chromium. Elements it can't describe
// natively come back as `raster` nodes; their elements are tagged data-pptx-node="<id>" so
// capture.mjs can screenshot exactly that element.
(() => {
  const SKIP_TAGS = new Set(["script", "style", "noscript", "template", "link", "meta", "head", "title"]);
  // Drawn as pictures; later tasks turn tables, code blocks and charts native.
  const RASTER_TAGS = {
    canvas: "canvas", svg: "inline SVG", iframe: "iframe", video: "video", audio: "audio",
    object: "embedded object", embed: "embedded object", math: "math", table: "table", pre: "code block",
  };
  const REPLACED = new Set(["img", "svg", "canvas", "video", "iframe", "object", "embed", "math", "audio"]);
  const IMAGE_FORMATS = { png: "png", jpg: "jpg", jpeg: "jpg", gif: "gif", svg: "svg", "svg+xml": "svg" };
  const GENERIC_FONTS = new Set(["serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui", "ui-serif", "ui-sans-serif", "ui-monospace", "ui-rounded", "emoji", "math", "fangsong"]);
  // Stands between two paragraphs (e.g. loose-list <p>s) joined into one run list.
  const BREAK = { br: true };
  const r2 = (v) => Math.round(v * 100) / 100;

  // Splits a comma-separated CSS value at top level only (commas inside rgba(...) etc. stay).
  function splitTop(value) {
    const out = [];
    let depth = 0;
    let cur = "";
    for (const ch of String(value)) {
      if (ch === "(") depth++;
      else if (ch === ")") depth--;
      if (ch === "," && depth === 0) {
        out.push(cur.trim());
        cur = "";
      } else cur += ch;
    }
    if (cur.trim()) out.push(cur.trim());
    return out;
  }

  const probe = document.createElement("canvas").getContext("2d", { willReadFrequently: true });

  // Any CSS color -> { hex: "RRGGBB", alpha }, or null when fully transparent. Computed styles
  // are usually rgb()/rgba(); anything else (oklch(), color()) goes through a 1px canvas.
  function parseColor(css) {
    if (!css || css === "transparent" || css === "none") return null;
    let r, g, b, a;
    const m = /^rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*(?:,\s*([\d.]+)\s*)?\)$/.exec(String(css).trim());
    if (m) {
      r = +m[1]; g = +m[2]; b = +m[3]; a = m[4] === undefined ? 1 : +m[4];
    } else {
      probe.clearRect(0, 0, 1, 1);
      probe.fillStyle = "rgba(0, 0, 0, 0)";
      probe.fillStyle = css;
      probe.fillRect(0, 0, 1, 1);
      const d = probe.getImageData(0, 0, 1, 1).data;
      r = d[0]; g = d[1]; b = d[2]; a = d[3] / 255;
    }
    if (!(a > 0.001)) return null;
    const hex = [r, g, b].map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, "0")).join("").toUpperCase();
    return { hex, alpha: Math.round(a * 1000) / 1000 };
  }

  function firstFamily(list) {
    const fams = String(list || "").split(",").map((f) => f.trim().replace(/^["']|["']$/g, "")).filter(Boolean);
    return fams.find((f) => !GENERIC_FONTS.has(f.toLowerCase())) || fams[0] || "";
  }

  function formatOf(src) {
    if (!src) return null;
    const data = /^data:image\/([a-z0-9+.-]+)/i.exec(src);
    let ext = data ? data[1] : null;
    if (!ext) {
      try {
        ext = (/\.([a-z0-9]+)$/i.exec(new URL(src, location.href).pathname) || [])[1];
      } catch {
        ext = null;
      }
    }
    return IMAGE_FORMATS[String(ext || "").toLowerCase()] || null;
  }

  // A transform that only rotates (and translates) -> { rotateOnly, rotate }; anything else
  // (scale, skew, mirror, 3D) -> { rotateOnly: false }; none -> null.
  function transformOf(cs) {
    if (!cs.transform || cs.transform === "none") return null;
    const m = new DOMMatrixReadOnly(cs.transform);
    if (!m.is2D) return { rotateOnly: false };
    const sx = Math.hypot(m.a, m.b);
    const sy = Math.hypot(m.c, m.d);
    const skewed = Math.abs(m.a * m.c + m.b * m.d) > 1e-6;
    const mirrored = m.a * m.d - m.b * m.c < 0;
    if (skewed || mirrored || Math.abs(sx - 1) > 0.01 || Math.abs(sy - 1) > 0.01) return { rotateOnly: false };
    const rotate = (Math.atan2(m.b, m.a) * 180) / Math.PI;
    return Math.abs(rotate) < 0.01 ? null : { rotateOnly: true, rotate: r2(rotate) };
  }

  function decoOf(cs) {
    const line = cs.textDecorationLine || "";
    return { underline: line.includes("underline"), strike: line.includes("line-through") };
  }

  function mergeDeco(d, cs) {
    const own = decoOf(cs);
    return { underline: d.underline || own.underline, strike: d.strike || own.strike };
  }

  function transformText(text, cs) {
    const ws = cs.whiteSpace;
    let t = /^(pre|pre-wrap|break-spaces)$/.test(ws) ? text : ws === "pre-line" ? text.replace(/[ \t]+/g, " ") : text.replace(/\s+/g, " ");
    if (cs.textTransform === "uppercase") t = t.toUpperCase();
    else if (cs.textTransform === "lowercase") t = t.toLowerCase();
    else if (cs.textTransform === "capitalize") t = t.replace(/(^|\s)(\S)/g, (_, s, c) => s + c.toUpperCase());
    return t;
  }

  function textAlign(cs) {
    const a = cs.textAlign;
    if (a.includes("center")) return "center";
    if (a.includes("right")) return "right";
    return a;
  }

  function sameColor(a, b) {
    if (!a || !b) return !a && !b;
    return a.hex === b.hex && a.alpha === b.alpha;
  }

  function sameStyle(a, b) {
    return a.font === b.font && a.size === b.size && a.bold === b.bold && a.italic === b.italic &&
      a.underline === b.underline && a.strike === b.strike && a.spacing === b.spacing &&
      sameColor(a.color, b.color) && sameColor(a.highlight, b.highlight);
  }

  function runStyle(cs, deco, highlight) {
    return {
      font: firstFamily(cs.fontFamily),
      size: r2(parseFloat(cs.fontSize)),
      bold: (parseInt(cs.fontWeight, 10) || 400) >= 600,
      italic: cs.fontStyle !== "normal",
      color: parseColor(cs.color),
      spacing: cs.letterSpacing === "normal" ? 0 : r2(parseFloat(cs.letterSpacing) || 0),
      underline: deco.underline,
      strike: deco.strike,
      highlight: highlight || null,
    };
  }

  // Collapses whitespace across run boundaries like the browser does, trims each line, and
  // merges neighbouring runs that look the same.
  function tidyRuns(runs, cs) {
    if (!/^(pre|pre-wrap|break-spaces)$/.test(cs.whiteSpace)) {
      let skipSpace = true;
      for (const r of runs) {
        if (r.br) {
          skipSpace = true;
          continue;
        }
        if (skipSpace) r.text = r.text.replace(/^ +/, "");
        if (r.text) skipSpace = r.text.endsWith(" ");
      }
      for (let i = 0; i < runs.length; i++) {
        if (runs[i].br) continue;
        const next = runs.slice(i + 1).find((x) => x.br || x.text);
        if (!next || next.br) runs[i].text = runs[i].text.replace(/ +$/, "");
      }
    }
    const out = [];
    for (const r of runs) {
      if (!r.text) continue;
      const last = out[out.length - 1];
      if (last && sameStyle(last, r)) last.text += r.text;
      else out.push({ ...r });
    }
    return out.map(({ br, ...r }) => r);
  }

  window.__pptxExtract = function extract(pageEl, canvas) {
    const origin = pageEl.getBoundingClientRect();
    const scale = origin.width / canvas.w || 1;
    const layoutRoot = pageEl.querySelector(".slidev-layout") || pageEl.firstElementChild || pageEl;
    const titleEl = pageEl.querySelector("h1");
    const nodes = [];
    let nextId = 0;

    const rectOf = (r) => ({ x: r2((r.left - origin.left) / scale), y: r2((r.top - origin.top) / scale), w: r2(r.width / scale), h: r2(r.height / scale) });
    const onCanvas = (rc) => rc.w > 0.5 && rc.h > 0.5 && rc.x < canvas.w && rc.y < canvas.h && rc.x + rc.w > 0 && rc.y + rc.h > 0;
    const labelOf = (el) => `${el.localName}${[...el.classList].slice(0, 2).map((c) => `.${c}`).join("")}`;
    const mark = (el) => {
      const id = nextId++;
      el.setAttribute("data-pptx-node", String(id));
      return id;
    };

    function contentBox(el, cs, rc) {
      const n = (p) => parseFloat(cs[p]) || 0;
      const l = n("borderLeftWidth") + n("paddingLeft");
      const t = n("borderTopWidth") + n("paddingTop");
      const r = n("borderRightWidth") + n("paddingRight");
      const b = n("borderBottomWidth") + n("paddingBottom");
      return { x: r2(rc.x + l), y: r2(rc.y + t), w: r2(Math.max(0, rc.w - l - r)), h: r2(Math.max(0, rc.h - t - b)) };
    }

    // The element's own box with the rotation taken out (getBoundingClientRect is the rotated hull).
    function unrotatedRect(el, r) {
      const c = rectOf(r);
      const w = el.offsetWidth;
      const h = el.offsetHeight;
      return { x: r2(c.x + c.w / 2 - w / 2), y: r2(c.y + c.h / 2 - h / 2), w, h };
    }

    function parseShadow(css) {
      if (!css || css === "none") return null;
      for (const part of splitTop(css)) {
        if (/\binset\b/.test(part)) continue;
        const colorMatch = /(rgba?\([^)]*\)|#[0-9a-f]{3,8}\b)/i.exec(part);
        const color = parseColor(colorMatch ? colorMatch[1] : "rgb(0, 0, 0)");
        if (!color) continue;
        const lengths = (part.replace(colorMatch ? colorMatch[1] : "", "").match(/-?[\d.]+px/g) || []).map(parseFloat);
        return { x: lengths[0] || 0, y: lengths[1] || 0, blur: lengths[2] || 0, color };
      }
      return null;
    }

    function decoration(cs, rc) {
      const sides = ["Top", "Right", "Bottom", "Left"]
        .map((s, i) => {
          const width = parseFloat(cs[`border${s}Width`]) || 0;
          const style = cs[`border${s}Style`];
          const color = width && style !== "none" && style !== "hidden" ? parseColor(cs[`border${s}Color`]) : null;
          return { side: "trbl"[i], width, color };
        })
        .filter((s) => s.color);
      const uniform = sides.length === 4 && sides.every((s) => s.width === sides[0].width && sameColor(s.color, sides[0].color));
      const rad = cs.borderTopLeftRadius || "0px";
      const radius = rad.endsWith("%") ? (parseFloat(rad) / 100) * Math.min(rc.w, rc.h) : parseFloat(rad) || 0;
      return {
        fill: parseColor(cs.backgroundColor),
        border: uniform ? { width: sides[0].width, color: sides[0].color } : null,
        sides: !uniform && sides.length ? sides : null,
        radius: r2(radius),
        shadow: parseShadow(cs.boxShadow),
      };
    }

    // A single url() background image -> its source and fit; anything else is unsupported.
    function backgroundImageOf(cs) {
      const bg = cs.backgroundImage;
      if (!bg || bg === "none") return null;
      // A `background` shorthand ending in a color reports an extra "none" image layer.
      if (splitTop(bg).filter((l) => l !== "none").length > 1) return { unsupported: "background layers" };
      const m = /^url\("?(.*?)"?\)$/.exec(bg);
      if (!m) return { unsupported: /gradient\(/.test(bg) ? "gradient" : "background image" };
      const src = new URL(m[1], location.href).href;
      const format = formatOf(src);
      if (!format || format === "svg") return { unsupported: "background image" };
      const size = cs.backgroundSize;
      return { src, format, fit: size === "cover" ? "cover" : size === "contain" ? "contain" : "fill" };
    }

    function rasterReason(el, cs, tag, op) {
      if (RASTER_TAGS[tag]) return RASTER_TAGS[tag];
      if (tag === "img") {
        const fmt = formatOf(el.currentSrc || el.src);
        if (fmt === "svg") return "SVG image";
        if (!fmt) return "image format";
        if (op < 0.999) return "image opacity";
      }
      if (cs.filter && cs.filter !== "none") return "CSS filter";
      if (cs.backdropFilter && cs.backdropFilter !== "none") return "backdrop-filter";
      const mask = cs.maskImage || cs.webkitMaskImage;
      if (mask && mask !== "none") return "mask";
      if (cs.clipPath && cs.clipPath !== "none") return "clip-path";
      if (cs.mixBlendMode && cs.mixBlendMode !== "normal") return "blend mode";
      if ((cs.backgroundClip || cs.webkitBackgroundClip) === "text") return "background-clip: text";
      const bg = backgroundImageOf(cs);
      if (bg && bg.unsupported) return bg.unsupported;
      const t = transformOf(cs);
      if (t && !t.rotateOnly) return "transform";
      if (t && el.children.length) return "rotated content";
      return null;
    }

    function isInlineFlow(el) {
      if (REPLACED.has(el.localName) || SKIP_TAGS.has(el.localName)) return false;
      const d = getComputedStyle(el).display;
      return d === "inline" || d === "contents";
    }

    function inheritedDeco(el) {
      let d = { underline: false, strike: false };
      for (let e = el; e && e !== pageEl.parentElement; e = e.parentElement) d = mergeDeco(d, getComputedStyle(e));
      return d;
    }

    // Text runs from inline content. Elements that aren't inline flow (images, inline-blocks)
    // are collected in `deferred` and become nodes of their own.
    function collectRuns(list, deco, highlight, runs, deferred) {
      for (const n of list) {
        if (n === BREAK) {
          if (runs.length) runs.push({ ...runs[runs.length - 1], text: "\n", br: true });
          continue;
        }
        if (n.nodeType === Node.TEXT_NODE) {
          const pcs = getComputedStyle(n.parentElement);
          runs.push({ ...runStyle(pcs, deco, highlight), text: transformText(n.data, pcs) });
          continue;
        }
        if (n.nodeType !== Node.ELEMENT_NODE) continue;
        const ccs = getComputedStyle(n);
        if (ccs.display === "none" || ccs.visibility === "hidden") continue;
        if (n.localName === "br") {
          runs.push({ ...runStyle(ccs, deco, highlight), text: "\n", br: true });
          continue;
        }
        if (!isInlineFlow(n)) {
          deferred.push(n);
          continue;
        }
        collectRuns(n.childNodes, mergeDeco(deco, ccs), parseColor(ccs.backgroundColor) || highlight, runs, deferred);
      }
    }

    // Measures a list of inline nodes: its rect, how many lines it wraps to, and where its
    // first and last lines start.
    function measure(list) {
      const dom = list.filter((n) => n !== BREAK);
      if (!dom.length) return null;
      const range = document.createRange();
      range.setStartBefore(dom[0]);
      range.setEndAfter(dom[dom.length - 1]);
      const rr = range.getBoundingClientRect();
      if (!rr.width || !rr.height) return null;
      const rects = [...range.getClientRects()].filter((x) => x.width > 0 && x.height > 0).sort((a, b) => a.top - b.top);
      let lines = 0;
      let bottom = -Infinity;
      let firstTop = rr.top;
      let lastTop = rr.top;
      for (const x of rects) {
        if (x.top >= bottom - 2) {
          if (lines === 0) firstTop = x.top;
          lines++;
          lastTop = x.top;
          bottom = x.bottom;
        } else bottom = Math.max(bottom, x.bottom);
      }
      return { rect: rectOf(rr), lines: Math.max(1, lines), firstTop: r2((firstTop - origin.top) / scale), lastTop: r2((lastTop - origin.top) / scale) };
    }

    function lineHeightOf(cs, m) {
      const lh = parseFloat(cs.lineHeight);
      return r2(lh > 0 ? lh : m.rect.h / m.lines);
    }

    function textNode(el, cs, segment, op, box, rotate) {
      const raw = [];
      const deferred = [];
      collectRuns(segment, inheritedDeco(el), null, raw, deferred);
      const runs = tidyRuns(raw, cs);
      const m = runs.length ? measure(segment) : null;
      if (m) {
        const multi = m.lines > 1;
        let rect = m.rect;
        if (rotate && box) rect = contentBox(el, cs, box);
        else if (multi && box) {
          const cb = contentBox(el, cs, box);
          rect = { x: cb.x, y: m.rect.y, w: cb.w, h: m.rect.h };
        }
        if (onCanvas(rect)) {
          nodes.push({ kind: "text", rect, role: el === titleEl ? "title" : "body", align: textAlign(cs), lineHeight: lineHeightOf(cs, m), singleLine: !multi, runs, opacity: op, rotate });
        }
      }
      for (const d of deferred) walk(d, op);
    }

    // Splits an element's children into runs of inline content (each a text node) and block
    // children (walked on their own), in document order.
    function content(el, cs, op, box, rotate) {
      let segment = [];
      const flush = () => {
        if (segment.length) textNode(el, cs, segment, op, box, rotate);
        segment = [];
      };
      for (const child of el.childNodes) {
        if (child.nodeType === Node.TEXT_NODE) segment.push(child);
        else if (child.nodeType === Node.ELEMENT_NODE) {
          if (isInlineFlow(child)) segment.push(child);
          else {
            flush();
            walk(child, op);
          }
        }
      }
      flush();
    }

    function boxNode(el, cs, rc, rotate, op) {
      if (el === pageEl || !onCanvas(rc)) return;
      const d = decoration(cs, rc);
      if (d.fill || d.border || d.sides || d.shadow) nodes.push({ kind: "box", rect: rc, ...d, opacity: op, rotate });
      const bg = backgroundImageOf(cs);
      if (bg && !bg.unsupported) {
        nodes.push({ kind: "image", id: mark(el), rect: rc, src: bg.src, format: bg.format, fit: bg.fit, natural: null, alt: "", border: null, shadow: null, opacity: op, rotate });
      }
    }

    function imageNode(el, cs, rc, rotate, op) {
      const src = el.currentSrc || el.src;
      if (!el.complete || !el.naturalWidth) {
        nodes.push({ kind: "missing", label: labelOf(el), src });
        return;
      }
      if (!onCanvas(rc)) return;
      const d = decoration(cs, rc);
      nodes.push({
        kind: "image", id: mark(el), rect: contentBox(el, cs, rc), src, format: formatOf(src), fit: cs.objectFit || "fill",
        natural: { w: el.naturalWidth, h: el.naturalHeight }, alt: el.alt || "", border: d.border, shadow: d.shadow, opacity: op, rotate,
      });
    }

    function collectItems(listEl, box, level, items, op) {
      let first = true;
      for (const li of listEl.children) {
        if (li.localName !== "li") continue;
        const lcs = getComputedStyle(li);
        if (lcs.display === "none" || lcs.visibility === "hidden") continue;
        const segment = [];
        const nested = [];
        const blocks = [];
        for (const c of li.childNodes) {
          if (c.nodeType === Node.TEXT_NODE) segment.push(c);
          else if (c.nodeType !== Node.ELEMENT_NODE) continue;
          else if (c.localName === "ul" || c.localName === "ol") nested.push(c);
          else if (isInlineFlow(c)) segment.push(c);
          else if (c.localName === "p") {
            if (segment.length) segment.push(BREAK);
            segment.push(...c.childNodes);
          } else blocks.push(c);
        }
        const raw = [];
        const deferred = [];
        collectRuns(segment, inheritedDeco(li), null, raw, deferred);
        const runs = tidyRuns(raw, lcs);
        const type = lcs.listStyleType;
        const marker = type === "none" ? "none" : /decimal|alpha|roman|latin|greek|armenian|georgian|cjk/.test(type) ? "number" : "bullet";
        if (runs.length) {
          const m = measure(segment);
          const lr = rectOf(li.getBoundingClientRect());
          const lineHeight = m ? lineHeightOf(lcs, m) : r2(parseFloat(lcs.fontSize) * 1.2);
          const top = m ? m.firstTop : lr.y;
          const prev = items[items.length - 1];
          const item = {
            level, marker,
            markerColor: parseColor(getComputedStyle(li, "::marker").color),
            indent: r2(lr.x + (parseFloat(lcs.paddingLeft) || 0) - box.x),
            spaceBefore: prev ? r2(Math.max(0, top - prev.lastTop - lineHeight)) : 0,
            lineHeight, align: textAlign(lcs), runs, lastTop: m ? m.lastTop : lr.y,
          };
          if (marker === "number" && first) item.start = listEl.localName === "ol" ? listEl.start || 1 : 1;
          items.push(item);
          first = false;
        }
        for (const n of nested) collectItems(n, box, level + 1, items, op);
        for (const b of [...blocks, ...deferred]) walk(b, op);
      }
    }

    function listNode(listEl, box, op) {
      const items = [];
      collectItems(listEl, box, 0, items, op);
      if (items.length && onCanvas(box)) nodes.push({ kind: "list", rect: box, items: items.map(({ lastTop, ...it }) => it), opacity: op });
    }

    function walk(el, opacity) {
      const tag = el.localName;
      if (SKIP_TAGS.has(tag)) return;
      const cs = getComputedStyle(el);
      if (cs.display === "none" || cs.visibility === "hidden" || cs.visibility === "collapse") return;
      const op = opacity * (parseFloat(cs.opacity) || 0);
      if (!(op > 0.001)) return;
      if (cs.display === "contents") {
        content(el, cs, op, null, 0);
        return;
      }
      if (el === layoutRoot) {
        content(el, cs, op, rectOf(el.getBoundingClientRect()), 0);
        return;
      }
      const r = el.getBoundingClientRect();
      const rc = rectOf(r);
      const reason = rasterReason(el, cs, tag, op);
      if (reason) {
        if (onCanvas(rc)) nodes.push({ kind: "raster", id: mark(el), rect: rc, reason, label: labelOf(el), alt: el.getAttribute("alt") || el.getAttribute("aria-label") || "" });
        return;
      }
      const t = transformOf(cs);
      const rotate = t ? t.rotate : 0;
      const box = rotate ? unrotatedRect(el, r) : rc;
      if (tag === "img") {
        imageNode(el, cs, box, rotate, op);
        return;
      }
      boxNode(el, cs, box, rotate, op);
      if (tag === "ul" || tag === "ol") {
        listNode(el, box, op);
        return;
      }
      content(el, cs, op, box, rotate);
    }

    function backgroundOf() {
      for (let e = layoutRoot; e && e !== pageEl.parentElement; e = e.parentElement) {
        const cs = getComputedStyle(e);
        if (cs.backgroundImage && cs.backgroundImage !== "none") {
          return { raster: true, reason: /gradient\(/.test(cs.backgroundImage) ? "gradient background" : "background image" };
        }
        const c = parseColor(cs.backgroundColor);
        if (c) return { color: c };
      }
      return { color: { hex: "FFFFFF", alpha: 1 } };
    }

    content(pageEl, getComputedStyle(pageEl), 1, rectOf(origin), 0);
    const layout = layoutRoot.classList && layoutRoot.classList.contains("slidev-layout")
      ? [...layoutRoot.classList].find((c) => c !== "slidev-layout") || "default"
      : "default";
    return { canvas: { w: canvas.w, h: canvas.h }, layout, background: backgroundOf(), nodes };
  };
})();
```

- [ ] **Step 7: Run the browser tests**

Run: `npm run test:render`
Expected: all extract tests pass.

If a position assertion is off by a fraction of a pixel, that comes from the font, not the logic. Compare against the actual value, and change the assertion to a tolerance (`Math.abs(a - b) < 1`) rather than changing the extractor.

- [ ] **Step 8: Run the unit suite**

Run: `npm test`
Expected: all pass (no browser needed).

- [ ] **Step 9: Commit**

```bash
git add package.json package-lock.json src/render/extract.js src/render/browser.mjs test/render/helpers.mjs test/render/extract.test.mjs
git commit -m "Render mode: in-page extractor (text, lists, boxes, images, pictures of the rest) and the browser test harness"
```

---

### Task 5: Deck server and per-slide capture

**Files:**
- Create: `src/render/server.mjs`, `src/render/capture.mjs`
- Test: `test/render-server.test.mjs` (no browser), `test/render/deck.test.mjs` (browser; skips without `../innotek-SlideDevBase`)

**Interfaces:**
- Consumes: `loadChromium` (Task 4), `extract.js` (Task 4), `imageSize` from `image-size`.
- Produces:
  - `slidevCli(deckDir) → { bin, version } | null`
  - `startDeckServer({ inputPath, timeout = 60000 }) → Promise<{ url, stop(): Promise<void>, log(): string, version }>`
  - `captureDeck({ baseUrl, chromium, timeout = 15000, screenshots = false, slidevVersion }) → Promise<{ canvas: { w, h }, slides }>`
  - Each slide is a Scene with `index`, `sourceIndex`, `notes`, `screenshot` (Buffer or null), and image nodes carrying `data`/`ext`/`natural`, raster nodes carrying `png`, and a raster background carrying `background.png`. A slide that failed is `{ index, sourceIndex, error }`.

- [ ] **Step 1: Write the failing server tests** — `test/render-server.test.mjs`

```js
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, mkdirSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { slidevCli, startDeckServer } from "../src/render/server.mjs";
import { loadChromium } from "../src/render/browser.mjs";

// A throwaway deck whose node_modules/@slidev/cli is a tiny fake CLI script (null = none).
function fakeDeck(cliSource, { hoisted = false } = {}) {
  const root = mkdtempSync(join(tmpdir(), "pptx-render-"));
  const deckDir = hoisted ? join(root, "decks", "talk") : root;
  mkdirSync(deckDir, { recursive: true });
  writeFileSync(join(deckDir, "slides.md"), "# Hi\n");
  if (cliSource != null) {
    const pkgDir = join(root, "node_modules", "@slidev", "cli");
    mkdirSync(join(pkgDir, "bin"), { recursive: true });
    writeFileSync(join(pkgDir, "package.json"), JSON.stringify({ name: "@slidev/cli", version: "52.0.0-test", bin: { slidev: "./bin/slidev.mjs" } }));
    writeFileSync(join(pkgDir, "bin", "slidev.mjs"), cliSource);
  }
  return { root, deckDir, inputPath: join(deckDir, "slides.md") };
}

const SERVES = [
  'import { createServer } from "node:http";',
  'const port = Number(process.argv[process.argv.indexOf("--port") + 1]);',
  'createServer((req, res) => res.end("ok")).listen(port, "127.0.0.1");',
].join("\n");

test("finds @slidev/cli next to the deck or hoisted into a parent folder", () => {
  const local = fakeDeck(SERVES);
  assert.equal(slidevCli(local.deckDir).bin, join(local.root, "node_modules/@slidev/cli/bin/slidev.mjs"));
  const hoisted = fakeDeck(SERVES, { hoisted: true });
  assert.deepEqual(slidevCli(hoisted.deckDir), { bin: join(hoisted.root, "node_modules/@slidev/cli/bin/slidev.mjs"), version: "52.0.0-test" });
  assert.equal(slidevCli(fakeDeck(null).deckDir), null);
});

test("without @slidev/cli the export stops and says to run npm install in the deck folder", async () => {
  const { inputPath, deckDir } = fakeDeck(null);
  await assert.rejects(startDeckServer({ inputPath }), (err) => err.message.includes("@slidev/cli") && err.message.includes(`run npm install in ${deckDir}`));
});

test("a Slidev that exits during startup reports its own error output", async () => {
  const { inputPath } = fakeDeck('console.error(\'The theme "@slidev/theme-default" was not found\'); process.exit(1);');
  await assert.rejects(startDeckServer({ inputPath }), /Slidev exited before it was ready \(code 1\):[\s\S]*The theme "@slidev\/theme-default" was not found/);
});

test("a Slidev that never answers times out, showing its output", async () => {
  const { inputPath } = fakeDeck('console.log("still starting"); setInterval(() => {}, 1000);');
  await assert.rejects(startDeckServer({ inputPath, timeout: 1500 }), /Slidev didn't start within 1.5s:[\s\S]*still starting/);
});

test("starts the deck's Slidev on a free local port and stops it", async () => {
  const { inputPath } = fakeDeck(SERVES);
  const server = await startDeckServer({ inputPath });
  assert.match(server.url, /^http:\/\/127\.0\.0\.1:\d+$/);
  assert.equal(await (await fetch(server.url)).text(), "ok");
  assert.equal(server.version, "52.0.0-test");
  await server.stop();
  await assert.rejects(fetch(server.url));
});

test("loadChromium explains how to install Playwright when it's missing", async () => {
  await assert.rejects(loadChromium("/nowhere", { resolve: () => null }), /npm i -D playwright-chromium[\s\S]*npx playwright install chromium/);
});

test("loadChromium takes playwright-chromium from the deck folder first", async () => {
  const dir = mkdtempSync(join(tmpdir(), "pptx-pw-"));
  const file = join(dir, "fake-playwright.mjs");
  writeFileSync(file, "export const chromium = { name: 'fake' };");
  const seen = [];
  const chromium = await loadChromium("/deck", {
    resolve: (name, base) => {
      seen.push(`${name}@${base}`);
      return name === "playwright-chromium" && base === "/deck" ? file : null;
    },
  });
  assert.equal(chromium.name, "fake");
  assert.deepEqual(seen, ["playwright-chromium@/deck"]);
});
```

- [ ] **Step 2: Run them to see them fail**

Run: `node --test test/render-server.test.mjs`
Expected: FAIL — `Cannot find module '.../src/render/server.mjs'`.

- [ ] **Step 3: Implement** — `src/render/server.mjs`

```js
// Render mode runs the deck's own Slidev (its version, theme and addons) as a dev server on a
// free local port, and stops it again afterwards.

import { spawn } from "node:child_process";
import { existsSync, readFileSync } from "node:fs";
import { createServer } from "node:net";
import { join, dirname, basename } from "node:path";

// @slidev/cli as the deck sees it: in its own node_modules or any parent's (workspaces hoist it).
export function slidevCli(deckDir) {
  for (let dir = deckDir; ; dir = dirname(dir)) {
    const pkgPath = join(dir, "node_modules", "@slidev", "cli", "package.json");
    if (existsSync(pkgPath)) {
      const pkg = JSON.parse(readFileSync(pkgPath, "utf8"));
      const rel = typeof pkg.bin === "string" ? pkg.bin : pkg.bin?.slidev;
      return rel ? { bin: join(dirname(pkgPath), rel), version: pkg.version } : null;
    }
    if (dirname(dir) === dir) return null;
  }
}

function freePort() {
  return new Promise((resolve, reject) => {
    const s = createServer();
    s.unref();
    s.on("error", reject);
    s.listen(0, "127.0.0.1", () => {
      const { port } = s.address();
      s.close(() => resolve(port));
    });
  });
}

export async function startDeckServer({ inputPath, timeout = 60000 }) {
  const deckDir = dirname(inputPath);
  const cli = slidevCli(deckDir);
  if (!cli) throw new Error(`render mode runs the deck's own Slidev, but @slidev/cli isn't installed for this deck - run npm install in ${deckDir}`);
  const port = await freePort();
  const child = spawn(process.execPath, [cli.bin, basename(inputPath), "--port", String(port), "--bind", "127.0.0.1"], {
    cwd: deckDir,
    stdio: ["ignore", "pipe", "pipe"],
    env: { ...process.env, NO_COLOR: "1" },
  });
  const lines = [];
  const keep = (chunk) => {
    lines.push(...String(chunk).split(/\r?\n/));
    if (lines.length > 50) lines.splice(0, lines.length - 50);
  };
  child.stdout.on("data", keep);
  child.stderr.on("data", keep);
  let exited = null;
  // "close" (not "exit") fires after the output streams are drained.
  child.on("close", (code, signal) => {
    exited = code ?? signal ?? "unknown";
  });
  const log = () => lines.filter((l) => l.trim()).join("\n");
  const stop = () =>
    new Promise((resolve) => {
      if (exited !== null) return resolve();
      child.once("close", () => resolve());
      child.kill();
      setTimeout(() => child.kill("SIGKILL"), 3000).unref();
    });

  const url = `http://127.0.0.1:${port}`;
  const deadline = Date.now() + timeout;
  while (Date.now() < deadline) {
    if (exited !== null) throw new Error(`Slidev exited before it was ready (code ${exited}):\n${log()}`);
    try {
      if ((await fetch(`${url}/`)).ok) return { url, stop, log, version: cli.version };
    } catch {
      // not listening yet
    }
    await new Promise((r) => setTimeout(r, 250));
  }
  await stop();
  throw new Error(`Slidev didn't start within ${timeout / 1000}s:\n${log()}`);
}
```

- [ ] **Step 4: Run the server tests**

Run: `node --test test/render-server.test.mjs`
Expected: all pass.

- [ ] **Step 5: Write the failing end-to-end capture test** — `test/render/deck.test.mjs`

```js
import { test } from "node:test";
import assert from "node:assert/strict";
import { existsSync } from "node:fs";
import { join } from "node:path";
import { ROOT, usableChromium } from "./helpers.mjs";
import { startDeckServer } from "../../src/render/server.mjs";
import { captureDeck } from "../../src/render/capture.mjs";

export const INNOTEK = join(ROOT, "../innotek-SlideDevBase/slides.md");
const chromium = existsSync(INNOTEK) ? await usableChromium() : null;
export const deckOpts = {
  skip: !existsSync(INNOTEK) ? "../innotek-SlideDevBase is not checked out" : !chromium ? "Playwright/Chromium not available" : false,
  timeout: 300000,
};
const PNG_MAGIC = "89504e47";

test("captures every slide of the innotek deck: scenes, notes, image bytes and screenshots", deckOpts, async () => {
  const server = await startDeckServer({ inputPath: INNOTEK });
  let captured;
  try {
    captured = await captureDeck({ baseUrl: server.url, chromium, screenshots: true, slidevVersion: server.version });
  } finally {
    await server.stop();
  }
  assert.deepEqual(captured.canvas, { w: 980, h: 552 });
  assert.equal(captured.slides.length, 7);
  assert.deepEqual(captured.slides.filter((s) => s.error), []);
  assert.deepEqual(captured.slides.map((s) => s.sourceIndex), [0, 1, 2, 3, 4, 5, 6]);
  assert.equal(captured.slides[0].layout, "cover");
  const cardRow = captured.slides[2];
  assert.equal(cardRow.nodes.find((n) => n.role === "title").runs[0].text, "Card row");
  assert.equal(cardRow.nodes.filter((n) => n.kind === "box" && n.radius === 8).length, 3);
  for (const s of captured.slides) {
    assert.equal(s.screenshot.subarray(0, 4).toString("hex"), PNG_MAGIC);
    for (const n of s.nodes.filter((x) => x.kind === "raster")) assert.equal(n.png.subarray(0, 4).toString("hex"), PNG_MAGIC, n.label);
    for (const n of s.nodes.filter((x) => x.kind === "image")) assert.ok(n.data.length > 0 && (!n.natural || n.natural.w > 0), n.src);
  }
});
```

- [ ] **Step 6: Run it to see it fail**

Run: `npm run test:render`
Expected: the deck test FAILS (`Cannot find module '.../src/render/capture.mjs'`), or it skips if `../innotek-SlideDevBase` isn't checked out. If it skips, you can't complete this step: tell the controller.

- [ ] **Step 7: Implement** — `src/render/capture.mjs`

```js
// Render mode capture: drives Chromium through every slide of the deck's running Slidev dev
// server and returns one Scene per slide (see extract.js), plus the bytes the writer needs -
// image files, and PNG screenshots of elements drawn as pictures.

import { readFileSync } from "node:fs";
import { imageSize } from "image-size";

const EXTRACT = readFileSync(new URL("./extract.js", import.meta.url), "utf8");

export async function captureDeck({ baseUrl, chromium, timeout = 15000, screenshots = false, slidevVersion = "unknown" }) {
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage({ viewport: { width: 980, height: 552 }, deviceScaleFactor: 2 });
    page.setDefaultTimeout(timeout);
    await page.goto(`${baseUrl}/1?print`, { waitUntil: "networkidle" });
    try {
      await page.waitForFunction(() => window.__slidev__?.nav?.total > 0 && !!document.querySelector(".slidev-page"));
    } catch {
      throw new Error(`unsupported Slidev version ${slidevVersion}: the page has no slide navigation data (window.__slidev__) or .slidev-page elements`);
    }
    const info = await page.evaluate(() => {
      const s = window.__slidev__;
      return {
        canvasWidth: s.configs.canvasWidth || 980,
        aspectRatio: s.configs.aspectRatio || 16 / 9,
        slides: s.nav.slides.map((r) => ({ no: r.no, sourceIndex: r.meta?.slide?.sourceIndex ?? r.no - 1, notes: r.meta?.slide?.note || null })),
      };
    });
    // Slidev's own slide height: Math.ceil(canvasWidth / aspectRatio) (980 x 16/9 -> 552).
    const canvas = { w: info.canvasWidth, h: Math.ceil(info.canvasWidth / info.aspectRatio) };
    await page.setViewportSize({ width: canvas.w, height: canvas.h });
    const slides = [];
    for (const s of info.slides) {
      try {
        slides.push(await captureSlide(page, baseUrl, s, canvas, screenshots));
      } catch (err) {
        slides.push({ index: s.no, sourceIndex: s.sourceIndex, error: String(err.message || err).split("\n")[0] });
      }
    }
    return { canvas, slides };
  } finally {
    await browser.close();
  }
}

async function captureSlide(page, baseUrl, s, canvas, screenshots) {
  await page.goto(`${baseUrl}/${s.no}?print`, { waitUntil: "networkidle" });
  const sel = `.slidev-page-${s.no}`;
  await page.waitForFunction((pageSel) => {
    const p = document.querySelector(pageSel);
    return !!p && p.getBoundingClientRect().width > 0;
  }, sel);
  await page.evaluate(() => document.fonts.ready.then(() => true));
  await waitForStableCanvases(page);
  const screenshot = screenshots ? await page.locator(sel).screenshot({ animations: "disabled" }) : null;
  await page.addScriptTag({ content: EXTRACT });
  const scene = await page.evaluate(({ pageSel, size }) => window.__pptxExtract(document.querySelector(pageSel), size), { pageSel: sel, size: canvas });
  for (const node of scene.nodes) {
    if (node.kind === "image") await loadImage(page, node);
    else if (node.kind === "raster") node.png = await elementShot(page, node.id);
  }
  // Last: it hides the slide's contents.
  if (scene.background.raster) scene.background.png = await backgroundShot(page, sel);
  return { ...scene, index: s.no, sourceIndex: s.sourceIndex, notes: s.notes, screenshot };
}

// Charts animate in; wait until every canvas stops changing (up to ~5 s).
async function waitForStableCanvases(page) {
  let previous = null;
  for (let i = 0; i < 20; i++) {
    const sig = await page.evaluate(() =>
      [...document.querySelectorAll("canvas")]
        .map((c) => {
          try {
            const d = c.toDataURL();
            return `${c.width}x${c.height}:${d.length}:${d.slice(-48)}`;
          } catch {
            return "tainted";
          }
        })
        .join("|"),
    );
    if (sig === "" || sig === previous) return;
    previous = sig;
    await page.waitForTimeout(250);
  }
}

async function fetchBytes(page, src) {
  if (src.startsWith("data:")) {
    const comma = src.indexOf(",");
    const body = src.slice(comma + 1);
    return src.slice(0, comma).endsWith(";base64") ? Buffer.from(body, "base64") : Buffer.from(decodeURIComponent(body), "utf8");
  }
  const res = await page.request.get(src);
  if (!res.ok()) throw new Error(`HTTP ${res.status()}`);
  return Buffer.from(await res.body());
}

// Image bytes for the writer; if they can't be fetched the element is drawn as a picture.
async function loadImage(page, node) {
  try {
    node.data = await fetchBytes(page, node.src);
    node.ext = node.format;
    if (!node.natural) {
      const d = imageSize(node.data);
      node.natural = { w: d.width, h: d.height };
    }
  } catch (err) {
    delete node.data;
    Object.assign(node, { kind: "raster", reason: `image not fetched: ${err.message}`, label: node.alt || node.src, png: await elementShot(page, node.id) });
  }
}

async function elementShot(page, id) {
  try {
    return await page.locator(`[data-pptx-node="${id}"]`).screenshot({ animations: "disabled" });
  } catch {
    return null;
  }
}

// The slide's background alone: everything inside the layout hidden, then one screenshot.
async function backgroundShot(page, sel) {
  await page.evaluate((pageSel) => {
    const pageEl = document.querySelector(pageSel);
    const root = pageEl.querySelector(".slidev-layout") || pageEl.firstElementChild || pageEl;
    for (const c of root.children) c.style.setProperty("visibility", "hidden", "important");
  }, sel);
  return page.locator(sel).screenshot({ animations: "disabled" });
}
```

- [ ] **Step 8: Run both suites**

Run: `npm test && npm run test:render`
Expected: all pass. The deck test takes about 30 s.

- [ ] **Step 9: Commit**

```bash
git add src/render/server.mjs src/render/capture.mjs test/render-server.test.mjs test/render/deck.test.mjs
git commit -m "Render mode: start the deck's own Slidev and capture every slide in Chromium"
```

---

### Task 6: `renderSpecs`, `convert({ render })`, `--render`, README

**Files:**
- Create: `src/render/index.mjs`
- Modify: `src/layout.mjs` (`layoutEach`), `src/index.mjs`, `bin/cli.mjs`, `README.md`
- Test: `test/render-index.test.mjs`, `test/render/deck.test.mjs` (append)

**Interfaces:**
- Consumes: `startDeckServer`, `captureDeck`, `loadChromium`, `sceneToSpec`, `deckStyle`.
- Produces:
  - `renderSpecs({ inputPath, userConfig, config, parsed, codeStore, screenshots, deps })` returns `{ specs, style, screenshots, translatedCharts }`. `deps` (tests only) is `{ start, chromiumFor, capture }`. The function mutates `config.slide` for non-16:9 canvases.
  - `layoutEach(parsedSlides, codeStore, config)` returns `SlideSpec[][]`, one array per parsed slide.
  - `convert({ ..., render: true | { screenshots: true } })`. With `screenshots`, the result also carries `screenshots: (Buffer|null)[]`.
  - CLI flag `--render`.

- [ ] **Step 1: Write the failing tests** — `test/render-index.test.mjs`

```js
import { test } from "node:test";
import assert from "node:assert/strict";
import { renderSpecs } from "../src/render/index.mjs";
import { mergeConfig } from "../src/config.mjs";
import { splitSlides, parseSlide } from "../src/parse.mjs";

const C = (hex) => ({ hex, alpha: 1 });
const run = (text) => ({ text, font: "Inter", size: 36, bold: true, italic: false, color: C("1F2937"), spacing: 0, underline: false, strike: false, highlight: null });
const scene = (index, sourceIndex, title, nodes = []) => ({
  index, sourceIndex, layout: "default", canvas: { w: 980, h: 552 }, notes: "n", background: { color: C("FFFFFF") }, screenshot: null,
  nodes: [{ kind: "text", role: "title", rect: { x: 50, y: 50, w: 300, h: 40 }, align: "left", lineHeight: 40, singleLine: true, opacity: 1, runs: [run(title)] }, ...nodes],
});
const titleOf = (spec) => spec.placeholders.title.map((p) => p.runs.map((r) => r.text).join("")).join("");

function setup(md, { userConfig = {}, slides, canvas = { w: 980, h: 552 }, captureError = null } = {}) {
  const config = mergeConfig(userConfig);
  const codeStore = [];
  const parsed = splitSlides(md).slides.map((s) => ({ ...parseSlide(s.content, codeStore, config), frontmatter: s.frontmatter }));
  const calls = { stopped: 0 };
  const deps = {
    chromiumFor: async () => ({ fake: true }),
    start: async () => ({ url: "http://127.0.0.1:1", version: "52", stop: async () => { calls.stopped++; } }),
    capture: async ({ baseUrl }) => {
      calls.baseUrl = baseUrl;
      if (captureError) throw captureError;
      return { canvas, slides };
    },
  };
  return { args: { inputPath: "/deck/slides.md", userConfig, config, parsed, codeStore, deps }, calls, config };
}

async function withWarnings(fn) {
  const warnings = [];
  const orig = console.warn;
  console.warn = (...a) => warnings.push(a.join(" "));
  try {
    return { result: await fn(), warnings };
  } finally {
    console.warn = orig;
  }
}

test("rendered scenes become specs; a slide that failed falls back to its markdown by sourceIndex, with a warning", async () => {
  const md = "# Cover\n\n---\n\n# Second\n\nText\n\n---\nhide: true\n---\n\n# Hidden\n\n---\n\n# Fourth\n";
  const raster = { kind: "raster", id: 0, rect: { x: 0, y: 0, w: 10, h: 10 }, reason: "canvas", label: "canvas", alt: "", png: Buffer.from("x") };
  const missing = { kind: "missing", label: "img", src: "http://x/a.png" };
  const { args, calls } = setup(md, { slides: [scene(1, 0, "Cover", [raster]), scene(2, 1, "Second", [missing]), { index: 3, sourceIndex: 3, error: "Timeout 15000ms exceeded." }] });
  const { result, warnings } = await withWarnings(() => renderSpecs(args));
  assert.deepEqual(result.specs.map(titleOf), ["Cover", "Second", "Fourth"]);
  assert.equal(result.specs[0].layout, "title");
  assert.equal(result.specs[1].layout, "titleOnly");
  assert.deepEqual(warnings, [
    "Notice: slide 1: canvas exported as a picture (canvas).",
    "Notice: slide 2: image http://x/a.png didn't load - skipped.",
    "Warning: slide 3 could not be rendered (Timeout 15000ms exceeded.) - exported from markdown instead.",
  ]);
  assert.equal(calls.stopped, 1);
  assert.equal(calls.baseUrl, "http://127.0.0.1:1");
  assert.equal(result.translatedCharts, 0);
});

test("the server is stopped even when capture fails, and the error propagates", async () => {
  const { args, calls } = setup("# A\n", { captureError: new Error("browser crashed") });
  await assert.rejects(renderSpecs(args), /browser crashed/);
  assert.equal(calls.stopped, 1);
});

test("config keys render mode doesn't need are named once", async () => {
  const { args } = setup("# A\n", { userConfig: { logo: { path: "x.png" }, cssVarColors: { "--a": "accent" } }, slides: [scene(1, 0, "A")] });
  const { warnings } = await withWarnings(() => renderSpecs(args));
  assert.deepEqual(warnings, ["Notice: render mode ignores logo, cssVarColors from pptx.config.mjs - the rendered page already shows them."]);
});

test("a 4:3 deck gets a 4:3 slide so nothing is cut off; 16:9 keeps the configured size", async () => {
  const four = setup("# A\n", { slides: [scene(1, 0, "A")], canvas: { w: 980, h: 735 } });
  const { result } = await withWarnings(() => renderSpecs(four.args));
  assert.equal(result.style.geometry.height, 10);
  const wide = setup("# A\n", { slides: [scene(1, 0, "A")] });
  assert.equal((await withWarnings(() => renderSpecs(wide.args))).result.style.geometry.height, 7.5);
});
```

- [ ] **Step 2: Run them to see them fail**

Run: `node --test test/render-index.test.mjs`
Expected: FAIL — `Cannot find module '.../src/render/index.mjs'`.

- [ ] **Step 3: Add `layoutEach` to `src/layout.mjs`**

Replace `layoutDeck` with:

```js
// One SlideSpec list per parsed slide (a long table can make several). Slides are numbered by
// their position in the export, so a warning about slide N matches slide N in the .pptx.
export function layoutEach(parsedSlides, codeStore, config) {
  const S = deckStyle(config);
  const out = [];
  let next = 0;
  for (const slide of parsedSlides) {
    const specs = slideSpecs(slide, next, codeStore, S);
    out.push(specs);
    next += specs.length;
  }
  return out;
}

export function layoutDeck(parsedSlides, codeStore, config) {
  return layoutEach(parsedSlides, codeStore, config).flat();
}
```

- [ ] **Step 4: Implement** — `src/render/index.mjs`

```js
// Render mode orchestration: start the deck's own Slidev, capture every slide in Chromium, and
// turn the scenes into SlideSpecs. A slide that couldn't be captured comes from the markdown
// (the static export) instead.

import { dirname } from "node:path";
import { startDeckServer } from "./server.mjs";
import { loadChromium } from "./browser.mjs";
import { captureDeck } from "./capture.mjs";
import { sceneToSpec } from "./scene.mjs";
import { deckStyle, layoutEach } from "../layout.mjs";

// Config keys that only restate what the page shows; render mode reads the page instead.
const UNUSED_KEYS = ["logo", "eyebrowClass", "cssVarColors"];

export async function renderSpecs({ inputPath, userConfig = {}, config, parsed, codeStore, screenshots = false, deps = {} }) {
  const { start = startDeckServer, chromiumFor = loadChromium, capture = captureDeck } = deps;
  const unused = UNUSED_KEYS.filter((k) => userConfig[k] != null);
  if (unused.length) console.warn(`Notice: render mode ignores ${unused.join(", ")} from pptx.config.mjs - the rendered page already shows them.`);

  // Playwright first, so a missing browser fails before Slidev is started.
  const chromium = await chromiumFor(dirname(inputPath));
  const server = await start({ inputPath });
  let captured;
  try {
    captured = await capture({ baseUrl: server.url, chromium, timeout: config.render?.timeout ?? 15000, screenshots, slidevVersion: server.version });
  } finally {
    await server.stop();
  }

  // A non-16:9 deck keeps its shape: the slide height follows Slidev's canvas.
  const { w, h } = captured.canvas;
  if (Math.abs(w / h - config.slide.width / config.slide.height) > 0.01) {
    config.slide = { ...config.slide, height: Math.round(((config.slide.width * h) / w) * 1000) / 1000 };
  }
  const style = deckStyle(config);
  const sceneStyle = { geometry: style.geometry, theme: style.theme, fonts: config.fonts };
  const fromMarkdown = layoutEach(parsed, codeStore, config);

  const specs = [];
  const shots = [];
  for (const scene of captured.slides) {
    const no = specs.length + 1;
    if (scene.error) {
      console.warn(`Warning: slide ${no} could not be rendered (${scene.error}) - exported from markdown instead.`);
      const fallback = fromMarkdown[scene.sourceIndex] || [];
      specs.push(...fallback);
      shots.push(...fallback.map(() => null));
      continue;
    }
    for (const node of scene.nodes) {
      if (node.kind === "raster") console.warn(`Notice: slide ${no}: ${node.label} exported as a picture (${node.reason}).`);
      if (node.kind === "missing") console.warn(`Notice: slide ${no}: image ${node.src} didn't load - skipped.`);
    }
    specs.push(sceneToSpec(scene, sceneStyle));
    shots.push(scene.screenshot || null);
  }
  const translatedCharts = specs.reduce((n, s) => n + (s.shapes || []).filter((x) => x.kind === "chart").length, 0);
  return { specs, style, screenshots: shots, translatedCharts };
}
```

- [ ] **Step 5: Wire `convert()`** — `src/index.mjs`

1. Add the import: `import { renderSpecs } from "./render/index.mjs";` and change the layout import to `import { layoutDeck, deckStyle } from "./layout.mjs";` (unchanged if already so).
2. In the JSDoc, add after the `configDir` line:

```js
 * @param {boolean|{screenshots?: boolean}} [options.render] - render the deck with its own Slidev in Chromium and convert what is drawn (see README "Render mode")
```

3. Replace everything from `config.logo = usableLogo(config.logo, configDir);` to the end of `convert()` with:

```js
  // Render mode draws whatever logo the page shows, so the config's logo isn't used there.
  config.logo = options.render ? null : usableLogo(config.logo, configDir);

  const { globalFrontmatter, slides } = splitSlides(raw);
  const codeStore = [];
  const parsed = slides.map((s) => ({ ...parseSlide(s.content, codeStore, config), frontmatter: s.frontmatter }));
  const chartBlocks = parsed.flatMap((s) => s.blocks).filter((b) => b.type === "chart");

  let specs;
  let style;
  let rendered = null;
  if (options.render) {
    rendered = await renderSpecs({ inputPath, userConfig: options.config || {}, config, parsed, codeStore, screenshots: options.render.screenshots === true });
    ({ specs, style } = rendered);
  } else {
    specs = layoutDeck(parsed, codeStore, config);
    style = deckStyle(config);
  }

  const buffer = await writePptx(specs, {
    geometry: style.geometry,
    theme: style.theme,
    fonts: config.fonts,
    lang: config.lang,
    title: globalFrontmatter.title || basename(inputPath, extname(inputPath)),
    author: config.author,
    timestamps: config.timestamps === true,
    // Drawn bottom-left on every slide via the slide master.
    logo: config.logo ? { path: config.logo.path, box: style.geometry.logo } : null,
  });

  const outputPath = resolve(options.outputPath || `${basename(inputPath, extname(inputPath))}-editable.pptx`);
  await writeFile(outputPath, buffer);

  // specs, not parsed slides: a long table can continue onto extra slides.
  const result = {
    outputPath,
    slideCount: specs.length,
    translatedCharts: rendered ? rendered.translatedCharts : chartBlocks.filter((b) => b.resolved).length,
    unresolvedCharts: rendered ? [] : [...new Set(chartBlocks.filter((b) => !b.resolved).map((b) => b.name))],
  };
  if (rendered && options.render.screenshots === true) result.screenshots = rendered.screenshots;
  return result;
}
```

- [ ] **Step 6: Add `--render` to the CLI** — `bin/cli.mjs`

1. Change the usage comment's first line to `// Usage: slidev-pptx-export [input.md] [output.pptx] [--config path/to/pptx.config.mjs] [--render]` and add below the existing comment paragraph:

```js
//
// --render runs the deck with its own Slidev in headless Chromium and converts what is actually
// drawn (custom CSS, HTML and Vue components included). Needs playwright-chromium installed.
```

2. Replace `parseArgs` with:

```js
function parseArgs(argv) {
  const positional = [];
  let configPath = null;
  let render = false;
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === "--config") {
      configPath = argv[++i];
    } else if (argv[i].startsWith("--config=")) {
      configPath = argv[i].slice("--config=".length);
    } else if (argv[i] === "--render") {
      render = true;
    } else {
      positional.push(argv[i]);
    }
  }
  return { positional, configPath, render };
}
```

3. In `main()`, change `const { positional, configPath } = ...` to `const { positional, configPath, render } = parseArgs(process.argv.slice(2));` and the convert call to `const result = await convert({ inputPath, outputPath, config, configDir, render });`.

- [ ] **Step 7: Run the unit suite**

Run: `npm test`
Expected: all pass, including the byte-identical static export test.

- [ ] **Step 8: Add the end-to-end export test** — append to `test/render/deck.test.mjs`

```js
import { readFileSync, mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname } from "node:path";
import { pathToFileURL } from "node:url";
import { convert } from "../../src/index.mjs";
import { lintPackage } from "../lint.mjs";
import { unzip } from "../helpers.mjs";

export async function exportInnotek(name = "innotek-render.pptx", overrides = {}) {
  const config = { ...(await import(pathToFileURL(join(dirname(INNOTEK), "pptx.config.mjs")).href)).default, ...overrides };
  const outputPath = join(mkdtempSync(join(tmpdir(), "pptx-render-")), name);
  const result = await convert({ inputPath: INNOTEK, outputPath, config, configDir: dirname(INNOTEK), render: true });
  const buffer = readFileSync(outputPath);
  return { result, buffer, z: await unzip(buffer) };
}

test("exports the innotek deck with --render: lint-clean, title placeholders, native cards and eyebrow", deckOpts, async () => {
  const { result, buffer, z } = await exportInnotek();
  assert.equal(result.slideCount, 7);
  assert.deepEqual(await lintPackage(buffer), []);
  const cards = z.text("ppt/slides/slide3.xml");
  assert.match(cards, /<p:ph type="title"\/><\/p:nvPr><\/p:nvSpPr><p:spPr><a:xfrm>[\s\S]*?<a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:noAutofit\/><\/a:bodyPr>[\s\S]*?Card row/);
  assert.equal((cards.match(/prst="roundRect"/g) || []).length, 3);
  assert.match(cards, /<a:solidFill><a:schemeClr val="bg2"\/><\/a:solidFill><a:ln w="12446"><a:solidFill><a:schemeClr val="accent4"\/>/);
  assert.match(cards, /<a:solidFill><a:schemeClr val="accent1"\/><\/a:solidFill><a:latin typeface="\+mj-lt"\/><a:ea typeface="\+mj-ea"\/><a:cs typeface="\+mj-cs"\/><\/a:rPr><a:t>OVERVIEW<\/a:t>/);
  assert.match(z.text("ppt/slides/slide1.xml"), /<p:bg><p:bgPr><a:blipFill/);
});
```

(Move the new `import` lines to the top of the file with the others. `12446` is 1 px = 0.98 pt in EMU: `Math.round(0.98 * 12700)`.)

- [ ] **Step 9: Run the browser suite**

Run: `npm run test:render`
Expected: all pass. If the outline width differs by 1 because of rounding, compute `Math.round(pt(1) * 12700)` with the Task 3 `pt` formula and use that value.

- [ ] **Step 10: Document render mode** — `README.md`

Add this section after the "Usage" section:

````markdown
## Render mode (`--render`)

By default the exporter reads only `slides.md`, so custom CSS, HTML and Vue components reach
the `.pptx` only through what `pptx.config.mjs` restates. With `--render` it instead runs the
deck with its own Slidev in headless Chromium and converts what is actually drawn:

```bash
npm i -D playwright-chromium   # once, in the deck folder (the package `slidev export` uses)
npx playwright install chromium
npx slidev-pptx-export slides.md slides-editable.pptx --render
```

- Text keeps its rendered font, size, weight, color, letter spacing and position; slide
  titles still go into title placeholders.
- Boxes with a background, border, rounded corners or shadow (cards, callouts) become native
  shapes; lists get native bullets in their rendered colors.
- Images become native pictures. Anything PowerPoint can't express natively is drawn as a
  picture of just that element, with a notice naming the slide.
- `colors` and `fonts` in `pptx.config.mjs` still build the theme (rendered colors that match
  become theme colors); `logo`, `eyebrowClass` and `cssVarColors` aren't needed.
  `render: { timeout }` sets the per-slide wait (default 15000 ms).
- A slide that can't be rendered falls back to the markdown export, with a warning.

Render mode needs the deck's dependencies installed (`npm install` in the deck folder) and
takes roughly 10–30 seconds per deck.
````

- [ ] **Step 11: Commit**

```bash
git add src/render/index.mjs src/layout.mjs src/index.mjs bin/cli.mjs README.md test/render-index.test.mjs test/render/deck.test.mjs
git commit -m "Render mode: --render exports what Slidev draws, falling back to markdown per slide"
```

---
# Phase 2 — Content

### Task 7: Writer — "plain" tables with per-cell formatting

**Files:**
- Modify: `src/ooxml/table.mjs`
- Test: `test/table.test.mjs`

**Interfaces:**
- Produces:
  - Table shape `variant: "plain"`, which uses the built-in "No Style, No Grid" style (`NO_STYLE`) and no `firstRow`.
  - A cell is either the existing paragraph array, or `{ paras, fill?, fillAlpha?, borders?: { l, r, t, b: { width (pt), color, alpha? } }, margin?: { l, r, t, b } (inches), anchor?: "t"|"ctr"|"b", gridSpan?, rowSpan?, hMerge?, vMerge? }`.
  - Export `NO_STYLE`.

- [ ] **Step 1: Write the failing test** — append to `test/table.test.mjs`

```js
test("plain tables write each cell's fill, borders, margins, anchor and merges, with the No Style table style", async () => {
  const PLAIN = {
    kind: "table", x: 0, y: 0, w: 4, h: 1, variant: "plain", header: true, colW: [2, 2], rowH: [0.5, 0.5],
    rows: [
      [{ paras: cell("Head"), fill: "#EEF4F0", borders: { b: { width: 1, color: "accent4" } }, margin: { l: 0.1, r: 0.1, t: 0.05, b: 0.05 }, anchor: "ctr", gridSpan: 2 }, { paras: [], hMerge: true }],
      [{ paras: cell("a") }, { paras: cell("b"), fill: "lt2", fillAlpha: 0.5 }],
    ],
  };
  const xml = tableXml(PLAIN, ctx);
  assert.match(xml, /^<a:tbl><a:tblPr bandRow="0"><a:tableStyleId>\{2D5ABB26-0587-4C30-8999-92F81FD0307C\}<\/a:tableStyleId><\/a:tblPr>/);
  assert.match(xml, /<a:tc gridSpan="2"><a:txBody>[\s\S]*?<\/a:txBody><a:tcPr marL="91440" marR="91440" marT="45720" marB="45720" anchor="ctr"><a:lnL w="0"><a:noFill\/><\/a:lnL><a:lnR w="0"><a:noFill\/><\/a:lnR><a:lnT w="0"><a:noFill\/><\/a:lnT><a:lnB w="12700"><a:solidFill><a:schemeClr val="accent4"\/><\/a:solidFill><\/a:lnB><a:solidFill><a:srgbClr val="EEF4F0"\/><\/a:solidFill><\/a:tcPr><\/a:tc><a:tc hMerge="1">/);
  assert.match(xml, /<a:solidFill><a:schemeClr val="bg2"><a:alpha val="50000"\/><\/a:schemeClr><\/a:solidFill><\/a:tcPr><\/a:tc><\/a:tr><\/a:tbl>$/);
  const buffer = await writePptx([{ layout: "blank", shapes: [PLAIN] }], DECK);
  assert.deepEqual(await lintPackage(buffer), []);
});
```

- [ ] **Step 2: Run it to see it fail**

Run: `node --test test/table.test.mjs`
Expected: the new test FAILS (object cells aren't supported).

- [ ] **Step 3: Implement** — `src/ooxml/table.mjs`

1. Change the xml import to `import { XML_DECL, NS, guid, emu, emuSize, solidFill } from "./xml.mjs";`
2. Add above `tableXml`:

```js
// Built-in "No Style, No Grid" table style: render mode writes every cell's fill and borders
// itself, so nothing may come from a style.
export const NO_STYLE = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}";

function lnXml(tag, side) {
  return side && side.width > 0
    ? `<a:${tag} w="${Math.round(side.width * 12700)}">${solidFill(side.color, side.alpha)}</a:${tag}>`
    : `<a:${tag} w="0"><a:noFill/></a:${tag}>`;
}

// A cell is a list of paragraphs (deck-styled tables, formatting from the table style) or, in
// render mode, an object carrying its own fill, borders, margins, anchor and merge flags.
// Cell bodies are written like placeholder bodies: a bare <a:bodyPr/>, no wrap/inset attrs.
function cellXml(cell, ctx, tcPr) {
  const body = (paras) => txBodyXml(paras, ctx, { placeholder: true }, "a:txBody");
  if (Array.isArray(cell)) return `<a:tc>${body(cell)}${tcPr}</a:tc>`;
  const attrs =
    (cell.gridSpan > 1 ? ` gridSpan="${cell.gridSpan}"` : "") +
    (cell.rowSpan > 1 ? ` rowSpan="${cell.rowSpan}"` : "") +
    (cell.hMerge ? ' hMerge="1"' : "") +
    (cell.vMerge ? ' vMerge="1"' : "");
  const m = cell.margin || {};
  const mar = ["l", "r", "t", "b"].map((k) => (m[k] != null ? ` mar${k.toUpperCase()}="${emu(m[k])}"` : "")).join("");
  const anchor = cell.anchor ? ` anchor="${cell.anchor}"` : "";
  const b = cell.borders || {};
  // CT_TableCellProperties child order: lnL, lnR, lnT, lnB, ..., fill.
  const own =
    `<a:tcPr${mar}${anchor}>${lnXml("lnL", b.l)}${lnXml("lnR", b.r)}${lnXml("lnT", b.t)}${lnXml("lnB", b.b)}` +
    `${cell.fill ? solidFill(cell.fill, cell.fillAlpha) : "<a:noFill/>"}</a:tcPr>`;
  return `<a:tc${attrs}>${body(cell.paras || [])}${own}</a:tc>`;
}
```

3. Replace `tableXml` with:

```js
export function tableXml(shape, ctx) {
  const cols = Math.max(1, ...shape.rows.map((r) => r.length));
  const widths = shape.colW && shape.colW.length === cols ? shape.colW : Array(cols).fill(shape.w / cols);
  // A row's h is its minimum height (rows still grow to fit their text). `rowH` (inches per
  // row) lets the layout give wrapped rows more room; otherwise the frame height is shared.
  const evenH = shape.h / Math.max(1, shape.rows.length);
  const rowH = (i) => emuSize(shape.rowH && shape.rowH.length === shape.rows.length ? shape.rowH[i] : evenH);
  const deflist = shape.variant === "deflist";
  const plain = shape.variant === "plain";
  const mar = deflist ? [76200, 76200, 50800, 50800] : [91440, 91440, 45720, 45720];
  const tcPr = `<a:tcPr marL="${mar[0]}" marR="${mar[1]}" marT="${mar[2]}" marB="${mar[3]}"${deflist ? ' anchor="ctr"' : ""}/>`;
  const rows = shape.rows
    .map((row, ri) => {
      // Short rows are padded to the grid width - a row with fewer a:tc than gridCols is a
      // repair prompt.
      const cells = Array.from({ length: cols }, (_, i) => row[i] || []);
      return `<a:tr h="${rowH(ri)}">${cells.map((cell) => cellXml(cell, ctx, tcPr)).join("")}</a:tr>`;
    })
    .join("");
  const styleId = plain ? NO_STYLE : deflist ? TABLE_STYLES.deflist : TABLE_STYLES.default;
  const firstRow = shape.header && !deflist && !plain ? ' firstRow="1"' : "";
  const grid = widths.map((w) => `<a:gridCol w="${emuSize(w)}"/>`).join("");
  return `<a:tbl><a:tblPr${firstRow} bandRow="0"><a:tableStyleId>${styleId}</a:tableStyleId></a:tblPr><a:tblGrid>${grid}</a:tblGrid>${rows}</a:tbl>`;
}
```

- [ ] **Step 4: Run all tests**

Run: `npm test`
Expected: all pass. The existing table tests prove array cells are written exactly as before.

- [ ] **Step 5: Commit**

```bash
git add src/ooxml/table.mjs test/table.test.mjs
git commit -m "Writer: plain tables with per-cell fills, borders, margins, anchors and merges"
```

---

### Task 8: Native tables and code blocks in render mode

**Files:**
- Modify: `src/render/extract.js`, `src/render/scene.mjs`
- Test: `test/render/extract.test.mjs`, `test/render-scene.test.mjs`, `test/render/deck.test.mjs`

**Interfaces:**
- Consumes: Task 7 plain tables.
- Produces these Scene nodes:
  - `{ kind: "table", rect, colW: [px], rowH: [px], rows: [[cell]], opacity }`. A cell is `{ runs, align, valign, lineHeight, fill, borders: { t?, r?, b?, l?: { width, color } }, padding: { t, r, b, l }, colspan, rowspan }`, or `{ covered: "h"|"v"|"hv" }` for a slot covered by a span.
  - `{ kind: "code", rect, lines: [Run[]], lineHeight, opacity }`, drawn on top of the `pre`'s own `box` node.

- [ ] **Step 1: Write the failing tests**

Append to `test/render/extract.test.mjs`:

```js
test("tables become table nodes: a grid with spans, measured columns and rows, per-cell fill, borders, padding and alignment", opts, async () => {
  const s = await extractHtml(browser,
    '<table style="border-collapse:collapse;font:14px/20px Arial">' +
      '<tr><th style="background:#eef4f0;border-bottom:2px solid #1a833b;padding:4px 8px;text-align:left">Option</th><th style="padding:4px 8px" colspan="2">Pros and cons</th></tr>' +
      '<tr><td style="padding:4px 8px;vertical-align:middle" rowspan="2">A</td><td style="padding:4px 8px">Simple</td><td style="padding:4px 8px">Slower</td></tr>' +
      '<tr><td style="padding:4px 8px">Fast</td><td style="padding:4px 8px;text-align:right">More <b>setup</b></td></tr></table>');
  const t = s.nodes.find((n) => n.kind === "table");
  assert.ok(t, JSON.stringify(s.nodes.map((n) => n.kind)));
  assert.equal(s.nodes.filter((n) => n.kind === "text").length, 0, "cell text lives only in the table node");
  assert.deepEqual([t.colW.length, t.rowH.length], [3, 3]);
  assert.ok(Math.abs(t.colW.reduce((a, b) => a + b, 0) - t.rect.w) < 1);
  const [head, span, covered] = t.rows[0];
  assert.deepEqual([textOf(head), head.fill, head.borders.b, head.padding, head.align], [
    "Option", { hex: "EEF4F0", alpha: 1 }, { width: 2, color: { hex: "1A833B", alpha: 1 } }, { t: 4, r: 8, b: 4, l: 8 }, "left",
  ]);
  assert.deepEqual([span.colspan, span.align], [2, "center"]);
  assert.deepEqual(covered, { covered: "h" });
  assert.deepEqual([t.rows[1][0].rowspan, t.rows[1][0].valign, t.rows[2][0]], [2, "middle", { covered: "v" }]);
  assert.deepEqual([textOf(t.rows[2][2]), t.rows[2][2].align, t.rows[2][2].runs[1].bold], ["More setup", "right", true]);
});

test("code blocks become a panel box plus one line of runs per source line, keeping token colors and leading spaces", opts, async () => {
  const s = await extractHtml(browser,
    '<pre style="margin:0;background:#eef4f0;padding:10px;font:13px/18px monospace"><code><span style="color:#1a589f">const</span> x = 1;\n<span style="color:#1a833b">  return</span> x;\n</code></pre>');
  const [panel, code] = s.nodes;
  assert.deepEqual([panel.kind, panel.fill.hex], ["box", "EEF4F0"]);
  assert.equal(code.kind, "code");
  assert.equal(code.lineHeight, 18);
  assert.deepEqual({ x: code.rect.x, y: code.rect.y }, { x: 50, y: 50 });
  assert.deepEqual(code.lines.map((l) => l.map((r) => [r.text, r.color.hex])), [
    [["const", "1A589F"], [" x = 1;", "000000"]],
    [["  return", "1A833B"], [" x;", "000000"]],
  ]);
});
```

Append to `test/render-scene.test.mjs`:

```js
test("table nodes become plain native tables with per-cell formatting and merges", () => {
  const cell = (t, extra = {}) => ({ runs: [run(t, { size: 14 })], align: "left", valign: "top", lineHeight: 20, fill: null, borders: {}, padding: { t: 4, r: 8, b: 4, l: 8 }, colspan: 1, rowspan: 1, ...extra });
  const node = {
    kind: "table", rect: { x: 40, y: 40, w: 300, h: 60 }, colW: [100, 200], rowH: [30, 30], opacity: 1,
    rows: [
      [cell("Head", { fill: C("EEF4F0"), borders: { b: { width: 2, color: C("1A833B") } }, colspan: 2, align: "center", valign: "middle" }), { covered: "h" }],
      [cell("a"), cell("b", { align: "right" })],
    ],
  };
  const [t] = sceneToSpec(scene([node]), STYLE).shapes;
  assert.deepEqual({ kind: t.kind, variant: t.variant, colW: t.colW, rowH: t.rowH }, { kind: "table", variant: "plain", colW: [inch(100), inch(200)], rowH: [inch(30), inch(30)] });
  assert.deepEqual(t.rows[0][0], {
    paras: [{ align: "ctr", lineSpacing: pt(20), runs: [{ text: "Head", size: pt(14), font: "body", color: "dk1" }] }],
    margin: { l: inch(8), r: inch(8), t: inch(4), b: inch(4) }, anchor: "ctr",
    borders: { b: { width: pt(2), color: "accent1" } }, fill: "lt2", gridSpan: 2,
  });
  assert.deepEqual(t.rows[0][1], { paras: [], hMerge: true, vMerge: false });
  assert.equal(t.rows[1][1].paras[0].align, "r");
});

test("code nodes become an unwrapped text box with one paragraph per line", () => {
  const node = {
    kind: "code", rect: { x: 50, y: 50, w: 400, h: 54 }, lineHeight: 18, opacity: 1,
    lines: [[run("const", { font: "JetBrains Mono", size: 13, color: C("1A589F") })], [], [run("  x", { font: "JetBrains Mono", size: 13 })]],
  };
  const [shape] = sceneToSpec(scene([node]), STYLE).shapes;
  assert.equal(shape.wrap, false);
  assert.deepEqual(shape.paras.map((p) => p.runs.map((r) => [r.text, r.font, r.color])), [[["const", "mono", "accent2"]], [], [["  x", "mono", "dk1"]]]);
  assert.ok(shape.paras.every((p) => p.lineSpacing === pt(18)));
});

test("table and code nodes write a lint-clean package", async () => {
  const cell = (t) => ({ runs: [run(t)], align: "left", valign: "top", lineHeight: 20, fill: C("EEF4F0"), borders: { b: { width: 1, color: C("D9E3DB") } }, padding: { t: 4, r: 8, b: 4, l: 8 }, colspan: 1, rowspan: 1 });
  const spec = sceneToSpec(scene([
    { kind: "table", rect: { x: 40, y: 40, w: 300, h: 60 }, colW: [150, 150], rowH: [30, 30], opacity: 1, rows: [[{ ...cell("A"), colspan: 2 }, { covered: "h" }], [cell("b"), cell("c")]] },
    { kind: "code", rect: { x: 40, y: 200, w: 400, h: 36 }, lineHeight: 18, opacity: 1, lines: [[run("x = 1", { font: "JetBrains Mono" })], []] },
  ]), STYLE);
  assert.deepEqual(await lintPackage(await writePptx([spec], { ...DECK, logo: null })), []);
});
```

- [ ] **Step 2: Run them to see them fail**

Run: `npm test` (the scene tests fail with `Unknown scene node kind "table"`), then `npm run test:render` (the table and code tests fail: they still come back as rasters).

- [ ] **Step 3: Implement the extractor** — `src/render/extract.js`

1. Replace the `RASTER_TAGS` constant and its comment with:

```js
  // Drawn as pictures; Task 9 turns Chart.js canvases into native charts.
  const RASTER_TAGS = {
    canvas: "canvas", svg: "inline SVG", iframe: "iframe", video: "video", audio: "audio",
    object: "embedded object", embed: "embedded object", math: "math",
  };
```

2. In `walk`, directly after the `if (tag === "img") { ... }` block, add:

```js
      if (tag === "table") {
        boxNode(el, cs, box, rotate, op);
        tableNode(el, op);
        return;
      }
      if (tag === "pre") {
        boxNode(el, cs, box, rotate, op);
        codeNode(el, cs, box, op);
        return;
      }
```

3. Add these functions inside `extract`, after `listNode`:

```js
    function hiddenWithin(el, root) {
      for (let e = el; e && e !== root; e = e.parentElement) {
        const s = getComputedStyle(e);
        if (s.display === "none" || s.visibility === "hidden") return true;
      }
      return false;
    }

    function cellBorders(cs) {
      const out = {};
      for (const [side, name] of [["t", "Top"], ["r", "Right"], ["b", "Bottom"], ["l", "Left"]]) {
        const width = parseFloat(cs[`border${name}Width`]) || 0;
        const style = cs[`border${name}Style`];
        const color = width && style !== "none" && style !== "hidden" ? parseColor(cs[`border${name}Color`]) : null;
        if (color) out[side] = { width, color };
      }
      return out;
    }

    // A cell's inline content (its <p> children included) as runs; other blocks are returned
    // to be walked as nodes of their own, on top of the table.
    function cellContent(cell) {
      const segment = [];
      const blocks = [];
      for (const c of cell.childNodes) {
        if (c.nodeType === Node.TEXT_NODE) segment.push(c);
        else if (c.nodeType !== Node.ELEMENT_NODE) continue;
        else if (isInlineFlow(c)) segment.push(c);
        else if (c.localName === "p") {
          if (segment.length) segment.push(BREAK);
          segment.push(...c.childNodes);
        } else blocks.push(c);
      }
      const raw = [];
      const deferred = [];
      collectRuns(segment, inheritedDeco(cell), null, raw, deferred);
      return { runs: tidyRuns(raw, getComputedStyle(cell)), m: measure(segment), blocks: [...blocks, ...deferred] };
    }

    function tableNode(table, op) {
      const rows = [...table.rows].filter((tr) => getComputedStyle(tr).display !== "none");
      if (!rows.length) return;
      // The HTML table grid: which slot each cell starts in, and which slots spans cover.
      const grid = [];
      const placed = [];
      rows.forEach((tr, r) => {
        grid[r] = grid[r] || [];
        let c = 0;
        for (const td of tr.cells) {
          while (grid[r][c]) c++;
          const cspan = Math.max(1, td.colSpan || 1);
          const rspan = Math.max(1, Math.min(td.rowSpan || 1, rows.length - r));
          for (let i = 0; i < rspan; i++) {
            grid[r + i] = grid[r + i] || [];
            for (let j = 0; j < cspan; j++) grid[r + i][c + j] = i || j ? (i && j ? "hv" : i ? "v" : "h") : "origin";
          }
          placed.push({ r, c, cspan, rspan, td });
          c += cspan;
        }
      });
      const nCols = Math.max(...grid.map((row) => row.length));
      // Column edges from the cells that start in each column; the last edge is the widest right side.
      const left = Array(nCols).fill(null);
      let right = -Infinity;
      for (const p of placed) {
        const rc = rectOf(p.td.getBoundingClientRect());
        if (left[p.c] === null || rc.x < left[p.c]) left[p.c] = rc.x;
        right = Math.max(right, rc.x + rc.w);
      }
      for (let c = 0; c < nCols; c++) if (left[c] === null) left[c] = c ? left[c - 1] : rectOf(table.getBoundingClientRect()).x;
      const colW = left.map((x, c) => r2((c + 1 < nCols ? left[c + 1] : right) - x));
      const rowRects = rows.map((tr) => rectOf(tr.getBoundingClientRect()));
      const rowH = rowRects.map((rc) => rc.h);
      const empty = () => ({ runs: [], align: "left", valign: "top", lineHeight: 0, fill: null, borders: {}, padding: { t: 0, r: 0, b: 0, l: 0 }, colspan: 1, rowspan: 1 });
      const cells = grid.map((row) => Array.from({ length: nCols }, (_, c) => (row[c] && row[c] !== "origin" ? { covered: row[c] } : empty())));
      const later = [];
      for (const p of placed) {
        const cs = getComputedStyle(p.td);
        const { runs, m, blocks } = cellContent(p.td);
        later.push(...blocks);
        const n = (prop) => parseFloat(cs[prop]) || 0;
        cells[p.r][p.c] = {
          runs, align: textAlign(cs), valign: cs.verticalAlign,
          lineHeight: m ? lineHeightOf(cs, m) : r2(parseFloat(cs.fontSize) * 1.2),
          fill: parseColor(cs.backgroundColor) || parseColor(getComputedStyle(p.td.parentElement).backgroundColor),
          borders: cellBorders(cs),
          padding: { t: n("paddingTop"), r: n("paddingRight"), b: n("paddingBottom"), l: n("paddingLeft") },
          colspan: p.cspan, rowspan: p.rspan,
        };
      }
      const rect = { x: left[0], y: rowRects[0].y, w: r2(colW.reduce((a, b) => a + b, 0)), h: r2(rowH.reduce((a, b) => a + b, 0)) };
      if (onCanvas(rect)) nodes.push({ kind: "table", rect, colW, rowH, rows: cells, opacity: op });
      for (const b of later) walk(b, op);
    }

    // A code block: one line of runs per source line, token colors kept, spaces untouched.
    function codeNode(pre, cs, box, op) {
      const lines = [[]];
      const walker = document.createTreeWalker(pre, NodeFilter.SHOW_TEXT);
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        if (hiddenWithin(n.parentElement, pre)) continue;
        const pcs = getComputedStyle(n.parentElement);
        const style = runStyle(pcs, decoOf(pcs), null);
        n.data.split("\n").forEach((part, i) => {
          if (i > 0) lines.push([]);
          if (part) lines[lines.length - 1].push({ ...style, text: part });
        });
      }
      while (lines.length && !lines[lines.length - 1].length) lines.pop();
      if (!lines.length) return;
      const lh = parseFloat(cs.lineHeight);
      nodes.push({
        kind: "code", rect: contentBox(pre, cs, box), lines: lines.map((runs) => tidyRuns(runs, { whiteSpace: "pre" })),
        lineHeight: r2(lh > 0 ? lh : parseFloat(cs.fontSize) * 1.5), opacity: op,
      });
    }
```

- [ ] **Step 4: Implement the mapping** — `src/render/scene.mjs`

Add these functions above `shapesOf`:

```js
const VALIGN = { middle: "ctr", bottom: "b" };

function tableShape(node, u, style) {
  const op = node.opacity ?? 1;
  const cellOf = (c) => {
    if (c.covered) return { paras: [], hMerge: c.covered.includes("h"), vMerge: c.covered.includes("v") };
    const para = { runs: c.runs.map((r) => runOf(r, u, style, op)) };
    const align = ALIGN[c.align];
    if (align && align !== "l") para.align = align;
    if (c.lineHeight) para.lineSpacing = u.pt(c.lineHeight);
    const out = {
      paras: [para],
      margin: { l: u.inch(c.padding.l), r: u.inch(c.padding.r), t: u.inch(c.padding.t), b: u.inch(c.padding.b) },
      anchor: VALIGN[c.valign] || "t",
      borders: {},
    };
    for (const [side, b] of Object.entries(c.borders || {})) {
      out.borders[side] = { width: u.pt(b.width), color: colorOf(b.color, style) };
      const a = alphaOf(b.color, op);
      if (a !== undefined) out.borders[side].alpha = a;
    }
    if (c.fill) {
      out.fill = colorOf(c.fill, style);
      const a = alphaOf(c.fill, op);
      if (a !== undefined) out.fillAlpha = a;
    }
    if (c.colspan > 1) out.gridSpan = c.colspan;
    if (c.rowspan > 1) out.rowSpan = c.rowspan;
    return out;
  };
  return {
    kind: "table", ...u.box(node.rect), variant: "plain",
    colW: node.colW.map((w) => u.inch(w)), rowH: node.rowH.map((h) => u.inch(h)),
    rows: node.rows.map((row) => row.map(cellOf)),
  };
}

function codeShape(node, u, style) {
  const op = node.opacity ?? 1;
  const paras = node.lines.map((runs) => ({ lineSpacing: u.pt(node.lineHeight), runs: runs.map((r) => runOf(r, u, style, op)) }));
  return { kind: "text", ...u.box(node.rect), geometry: "rect", anchor: "t", inset: 0, wrap: false, paras };
}
```

and add two cases to the `switch` in `shapesOf`:

```js
    case "table":
      return [tableShape(node, u, style)];
    case "code":
      return [codeShape(node, u, style)];
```

- [ ] **Step 5: Add the deck assertion** — append to `test/render/deck.test.mjs`

```js
test("the innotek table and code block are native", deckOpts, async () => {
  const { z } = await exportInnotek("innotek-content.pptx");
  const slide = z.text("ppt/slides/slide5.xml");
  assert.match(slide, /<a:tableStyleId>\{2D5ABB26-0587-4C30-8999-92F81FD0307C\}<\/a:tableStyleId>/);
  assert.match(slide, /<a:t>Option<\/a:t>/);
  assert.match(slide, /agent\.ask/);
});
```

- [ ] **Step 6: Run both suites**

Run: `npm test && npm run test:render`
Expected: all pass.

- [ ] **Step 7: Commit**

```bash
git add src/render/extract.js src/render/scene.mjs test/render/extract.test.mjs test/render-scene.test.mjs test/render/deck.test.mjs
git commit -m "Render mode: native tables with per-cell formatting and syntax-colored code blocks"
```

---

### Task 9: Native Chart.js charts

**Files:**
- Modify: `src/ooxml/chart.mjs`, `src/layout.mjs`, `src/render/extract.js`, `src/render/scene.mjs`, `src/render/index.mjs`
- Test: `test/chart.test.mjs`, `test/render/extract.test.mjs`, `test/render-scene.test.mjs`, `test/render-index.test.mjs`, `test/render/deck.test.mjs`

**Interfaces:**
- Produces:
  - ChartSpec fields `numFmt` (Excel format code), `legendPos` (`"t"|"b"|"l"|"r"`), `stacked`, `gridColor`/`gridAlpha`, and `background: null` (transparent).
  - Scene `{ kind: "chart", id, rect, components: [name], chart: { type, categories, series: [{ name, values: (number|null)[], color }], min, max, stacked, legend, legendPos, numFmt, gridColor, labelColor, legendColor }, opacity }`, plus `resolvedChart` when a resolver wins.
  - `groupedBarChart(resolved, S)` exported from `src/layout.mjs`.

- [ ] **Step 1: Write the failing tests**

Append to `test/chart.test.mjs`:

```js
test("render-mode chart options: number format, legend position, stacking, gridlines and a transparent background", async () => {
  const spec = { ...SPEC, numFmt: '0"%"', legendPos: "t", stacked: true, gridColor: "#14181A", gridAlpha: 0.1, background: null };
  const xml = chartXml(spec, { lang: "en-US", workbookRId: "rId1" });
  assert.match(xml, /<c:grouping val="stacked"\/>/);
  assert.match(xml, /<c:gapWidth val="45"\/><c:overlap val="100"\/>/);
  assert.match(xml, /<c:axPos val="l"\/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="14181A"><a:alpha val="10000"\/><\/a:srgbClr><\/a:solidFill><\/a:ln><\/c:spPr><\/c:majorGridlines><c:numFmt formatCode="0&quot;%&quot;" sourceLinked="0"\/>/);
  assert.match(xml, /<c:formatCode>0&quot;%&quot;<\/c:formatCode>/);
  assert.match(xml, /<c:legendPos val="t"\/>/);
  assert.match(xml, /<c:spPr><a:noFill\/><a:ln><a:noFill\/><\/a:ln><\/c:spPr><\/c:plotArea>/);
  const buffer = await writePptx([{ layout: "blank", shapes: [{ kind: "chart", x: 1, y: 1, w: 6, h: 3, chart: spec }] }], DECK);
  assert.deepEqual(await lintPackage(buffer), []);
});
```

Append to `test/render/extract.test.mjs`:

```js
test("a Chart.js bar chart becomes a chart node with its data, colors, axis range, legend and tick format; other canvases stay pictures", opts, async () => {
  const setup = () => {
    const [bar, pie] = document.querySelectorAll("canvas");
    const chart = (canvas, type) => ({
      canvas, config: { type },
      data: {
        labels: ["Q1", "Q2"],
        datasets: [
          { label: "Before", data: [22, 31], backgroundColor: "#9CA3AF" },
          { label: "After", data: [41, null], backgroundColor: ["#1A833B", "#000000"] },
          { label: "Hidden", data: [1, 2], hidden: true },
        ],
      },
      options: {
        scales: { y: { min: 0, max: 100, ticks: { color: "#5b6b63", callback: (v) => v + "%" }, grid: { color: "rgba(20, 24, 26, 0.1)" } }, x: { grid: { display: false } } },
        plugins: { legend: { position: "bottom", labels: { color: "#1f2937" } } },
      },
      scales: { y: {} },
    });
    bar.__vueParentComponent = { type: {}, exposed: { chart: { value: chart(bar, "bar") } }, parent: { type: { __name: "ExampleChart" }, parent: null } };
    pie.__vueParentComponent = { type: {}, exposed: { chart: { value: chart(pie, "pie") } }, parent: null };
  };
  const s = await extractHtml(browser,
    '<canvas width="400" height="200" style="display:block"></canvas><canvas width="100" height="100" style="display:block"></canvas><canvas width="50" height="50" style="display:block"></canvas>',
    { setup });
  const [bar, pie, plain] = s.nodes;
  assert.equal(bar.kind, "chart");
  assert.deepEqual(bar.components, ["ExampleChart"]);
  assert.deepEqual(bar.chart, {
    type: "column", categories: ["Q1", "Q2"],
    series: [{ name: "Before", values: [22, 31], color: { hex: "9CA3AF", alpha: 1 } }, { name: "After", values: [41, null], color: { hex: "1A833B", alpha: 1 } }],
    min: 0, max: 100, stacked: false, legend: true, legendPos: "b", numFmt: '0"%"',
    gridColor: { hex: "14181A", alpha: 0.1 }, labelColor: { hex: "5B6B63", alpha: 1 }, legendColor: { hex: "1F2937", alpha: 1 },
  });
  assert.deepEqual([pie.kind, pie.reason], ["raster", "chart type pie"]);
  assert.deepEqual([plain.kind, plain.reason], ["raster", "canvas"]);
});
```

Append to `test/render-scene.test.mjs`:

```js
test("chart nodes become native charts; uncolored series take the palette; a resolver's chart wins", async () => {
  const node = {
    kind: "chart", id: 3, rect: { x: 40, y: 200, w: 900, h: 240 }, components: ["ExampleChart"], opacity: 1,
    chart: {
      type: "column", categories: ["Q1", "Q2"],
      series: [{ name: "Before", values: [22, null], color: C("9CA3AF") }, { name: "After", values: [41, 56], color: null }],
      min: 0, max: 100, stacked: false, legend: true, legendPos: "b", numFmt: '0"%"', gridColor: C("14181A", 0.1), labelColor: C("5B6B63"), legendColor: C("1F2937"),
    },
  };
  const [shape] = sceneToSpec(scene([node]), STYLE).shapes;
  assert.deepEqual({ kind: shape.kind, x: shape.x, y: shape.y, w: shape.w, h: shape.h }, { kind: "chart", ...box(node.rect) });
  const c = shape.chart;
  assert.deepEqual(c.series.map((s) => [s.name, s.color]), [["Before", "accent5"], ["After", "accent2"]]);
  assert.ok(Number.isNaN(c.series[0].values[1]));
  assert.deepEqual(
    { min: c.min, max: c.max, numFmt: c.numFmt, legendPos: c.legendPos, gridColor: c.gridColor, gridAlpha: c.gridAlpha, background: c.background },
    { min: 0, max: 100, numFmt: '0"%"', legendPos: "b", gridColor: "#14181A", gridAlpha: 0.1, background: null },
  );
  assert.deepEqual(await lintPackage(await writePptx([sceneToSpec(scene([node]), STYLE)], { ...DECK, logo: null })), []);
  const resolved = { type: "column", categories: ["A"], series: [] };
  assert.equal(sceneToSpec(scene([{ ...node, resolvedChart: resolved }]), STYLE).shapes[0].chart, resolved);
});
```

Append to `test/render-index.test.mjs`:

```js
test("a chartResolvers entry for the chart's component replaces the captured data", async () => {
  const chartNode = {
    kind: "chart", id: 1, rect: { x: 40, y: 200, w: 900, h: 240 }, components: ["ExampleChart"], opacity: 1,
    chart: { type: "column", categories: ["Q1"], series: [{ name: "Page", values: [1], color: null }], min: null, max: null, stacked: false, legend: true, legendPos: "b", numFmt: null, gridColor: null, labelColor: null, legendColor: null },
  };
  const userConfig = { chartResolvers: { ExampleChart: () => ({ kind: "barchart", labels: ["From", "File"], datasets: [{ name: "Resolved", color: "1A833B", data: [5, 6] }], min: 0, max: 10 }) } };
  const { args } = setup("# A\n", { userConfig, slides: [scene(1, 0, "A", [chartNode])] });
  const { result } = await withWarnings(() => renderSpecs(args));
  const chart = result.specs[0].shapes.find((s) => s.kind === "chart").chart;
  assert.deepEqual(chart.categories, ["From", "File"]);
  assert.deepEqual(chart.series.map((s) => s.name), ["Resolved"]);
  assert.equal(result.translatedCharts, 1);
});
```

- [ ] **Step 2: Run them to see them fail**

Run: `npm test` (the chart, scene and index tests FAIL), then `npm run test:render` (the extract chart test FAILS).

- [ ] **Step 3: Implement the writer options** — `src/ooxml/chart.mjs`

1. In the series `numCache`, replace `<c:formatCode>General</c:formatCode>` with `<c:formatCode>${esc(spec.numFmt || "General")}</c:formatCode>`.
2. Replace the `const ticks = ...` line with:

```js
  const tickMarks = '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>';
  const ticks = `<c:numFmt formatCode="General" sourceLinked="1"/>${tickMarks}`;
  const valTicks = spec.numFmt ? `<c:numFmt formatCode="${esc(spec.numFmt)}" sourceLinked="0"/>${tickMarks}` : ticks;
  // CT_ValAx: majorGridlines sits between axPos and numFmt.
  const grid = spec.gridColor ? `<c:majorGridlines><c:spPr><a:ln w="9525">${solidFill(spec.gridColor, spec.gridAlpha)}</a:ln></c:spPr></c:majorGridlines>` : "";
```

3. In `valAx`, replace `<c:axPos val="${horizontal ? "b" : "l"}"/>${ticks}${NO_LINE}` with `<c:axPos val="${horizontal ? "b" : "l"}"/>${grid}${valTicks}${NO_LINE}`.
4. In `legend`, replace `<c:legendPos val="b"/>` with `<c:legendPos val="${spec.legendPos || "b"}"/>`.
5. Replace the `background` and `overlap` lines with:

```js
  const background = `<c:spPr>${spec.background ? solidFill(spec.background) : "<a:noFill/>"}<a:ln><a:noFill/></a:ln></c:spPr>`;
  const overlap = spec.stacked ? '<c:overlap val="100"/>' : spec.overlap != null ? `<c:overlap val="${spec.overlap}"/>` : "";
```

6. Replace `<c:grouping val="clustered"/>` with `<c:grouping val="${spec.stacked ? "stacked" : "clustered"}"/>`.

- [ ] **Step 4: Export the resolver chart builder** — `src/layout.mjs`

Replace `groupedBarShape` with:

```js
// A "barchart" resolver result as a native chart spec (also used by render mode, where a
// resolver overrides the chart captured from the page).
export function groupedBarChart(resolved, S) {
  const { color } = S;
  return {
    type: "column", categories: resolved.labels.map(String),
    series: resolved.datasets.map((d, j) => ({ name: String(d.name), color: resolverColor(S, d.color, PALETTE[j % PALETTE.length]), values: d.data.map(Number) })),
    min: resolved.min, max: resolved.max, valueAxis: true, legend: true, gapWidth: 45, overlap: -6,
    dataLabels: { color: color.text, size: 10 }, background: "lt1",
    catLabelColor: color.text, catLabelSize: 11, valLabelColor: color.muted, valLabelSize: 9, legendColor: color.text, legendSize: 11,
  };
}

function groupedBarShape(resolved, box, S) {
  return { kind: "chart", ...box, chart: groupedBarChart(resolved, S) };
}
```

- [ ] **Step 5: Implement the extractor** — `src/render/extract.js`

1. Remove `canvas: "canvas", ` from `RASTER_TAGS`, and change its comment to `// Drawn as pictures.`
2. In `walk`, directly after the `if (tag === "img") { ... }` block, add:

```js
      if (tag === "canvas") {
        chartNode(el, rc, op);
        return;
      }
```

3. Add these functions inside `extract`, after `codeNode`:

```js
    // The Chart.js chart drawn on a canvas: through the Vue component that owns it (vue-chartjs
    // exposes it as `chart`), else Chart.getChart when Chart.js is a page global.
    function findChart(canvasEl) {
      const components = [];
      let chart = null;
      let inst = canvasEl.__vueParentComponent || (canvasEl.parentElement && canvasEl.parentElement.__vueParentComponent);
      for (let i = 0; inst && i < 25; i++, inst = inst.parent) {
        const name = inst.type && (inst.type.__name || inst.type.name);
        if (name) components.push(name);
        for (const src of [inst.exposed, inst.setupState]) {
          if (chart || !src) continue;
          let c = null;
          try {
            c = src.chart;
          } catch {
            c = null;
          }
          const v = c && (c.value !== undefined ? c.value : c);
          if (v && v.config && v.data && v.canvas === canvasEl) chart = v;
        }
      }
      if (!chart && window.Chart && typeof window.Chart.getChart === "function") chart = window.Chart.getChart(canvasEl) || null;
      return chart ? { chart, components } : null;
    }

    // An axis tick callback like (v) => v + "%" -> the Excel format 0"%" (null if it isn't a
    // plain prefix/suffix around the number).
    function tickFormat(cb, scale) {
      if (typeof cb !== "function") return null;
      try {
        const a = String(cb.call(scale, 12, 0, [{ value: 12 }]));
        const b = String(cb.call(scale, 34.5, 1, [{ value: 34.5 }]));
        const i = a.indexOf("12");
        if (i < 0) return null;
        const pre = a.slice(0, i);
        const suf = a.slice(i + 2);
        if ((!pre && !suf) || b !== `${pre}34.5${suf}` || /["\\]/.test(pre + suf)) return null;
        const q = (s) => (s ? `"${s}"` : "");
        return `${q(pre)}0${q(suf)}`;
      } catch {
        return null;
      }
    }

    function chartSpec(chart) {
      const o = chart.options || {};
      const horizontal = o.indexAxis === "y";
      const scales = o.scales || {};
      const valScale = scales[horizontal ? "x" : "y"] || {};
      const catScale = scales[horizontal ? "y" : "x"] || {};
      const legend = (o.plugins && o.plugins.legend) || {};
      const datasets = [...(chart.data.datasets || [])].filter((d) => !d.hidden);
      const num = (v) => (typeof v === "number" && Number.isFinite(v) ? v : null);
      const value = (v) => (v == null ? null : typeof v === "object" ? num(horizontal ? v.x : v.y) : num(Number(v)));
      const datasetColor = (c) => (typeof c === "string" ? parseColor(c) : Array.isArray(c) && typeof c[0] === "string" ? parseColor(c[0]) : null);
      return {
        type: horizontal ? "bar" : "column",
        categories: Array.from(chart.data.labels || [], (l) => String(l)),
        series: datasets.map((d) => ({ name: String(d.label ?? ""), values: Array.from(d.data || [], value), color: datasetColor(d.backgroundColor) })),
        min: num(valScale.min),
        max: num(valScale.max),
        stacked: !!(valScale.stacked || catScale.stacked),
        legend: legend.display !== false && datasets.length > 0,
        legendPos: { top: "t", bottom: "b", left: "l", right: "r" }[legend.position] || "t",
        numFmt: tickFormat(valScale.ticks && valScale.ticks.callback, chart.scales && chart.scales[horizontal ? "x" : "y"]),
        gridColor: valScale.grid && valScale.grid.display === false ? null : parseColor((valScale.grid && valScale.grid.color) || "rgba(0, 0, 0, 0.1)"),
        labelColor: parseColor((valScale.ticks && valScale.ticks.color) || (catScale.ticks && catScale.ticks.color) || "#666666"),
        legendColor: parseColor((legend.labels && legend.labels.color) || "#666666"),
      };
    }

    function chartNode(canvasEl, rc, op) {
      if (!onCanvas(rc)) return;
      const found = findChart(canvasEl);
      const type = found && found.chart.config ? found.chart.config.type : null;
      if (!found || type !== "bar") {
        nodes.push({ kind: "raster", id: mark(canvasEl), rect: rc, reason: found ? `chart type ${type}` : "canvas", label: labelOf(canvasEl), alt: canvasEl.getAttribute("aria-label") || "" });
        return;
      }
      nodes.push({ kind: "chart", id: mark(canvasEl), rect: rc, components: found.components, chart: chartSpec(found.chart), opacity: op });
    }
```

- [ ] **Step 6: Implement the mapping** — `src/render/scene.mjs`

Add above `shapesOf`:

```js
const PALETTE = ["accent1", "accent2", "accent5", "accent6", "accent3"];

function chartShape(node, u, style) {
  if (node.resolvedChart) return { kind: "chart", ...u.box(node.rect), chart: node.resolvedChart };
  const c = node.chart;
  const color = (col, fallback) => (col ? colorOf(col, style) : fallback);
  const chart = {
    type: c.type, categories: c.categories,
    series: c.series.map((s, j) => ({ name: s.name, values: s.values.map((v) => (typeof v === "number" ? v : NaN)), color: color(s.color, PALETTE[j % PALETTE.length]) })),
    valueAxis: true, legend: c.legend, legendPos: c.legendPos, stacked: c.stacked, gapWidth: 60, background: null,
    catLabelColor: color(c.labelColor, "dk1"), catLabelSize: u.pt(12),
    valLabelColor: color(c.labelColor, "accent3"), valLabelSize: u.pt(12),
    legendColor: color(c.legendColor, "dk1"), legendSize: u.pt(12),
  };
  if (c.min != null) chart.min = c.min;
  if (c.max != null) chart.max = c.max;
  if (c.numFmt) chart.numFmt = c.numFmt;
  if (c.gridColor) {
    chart.gridColor = colorOf(c.gridColor, style);
    const a = alphaOf(c.gridColor);
    if (a !== undefined) chart.gridAlpha = a;
  }
  return { kind: "chart", ...u.box(node.rect), chart };
}
```

and add to the `switch` in `shapesOf`:

```js
    case "chart":
      return [chartShape(node, u, style)];
```

- [ ] **Step 7: Let resolvers win** — `src/render/index.mjs`

1. Change the layout import to `import { deckStyle, layoutEach, groupedBarChart } from "../layout.mjs";`
2. Add above `renderSpecs`:

```js
// A chartResolvers entry for the chart's Vue component (or any component around it) overrides
// the data captured from the page - but only a "barchart"; other resolver kinds describe HTML
// the page already renders.
function resolveChart(node, config, S) {
  const name = node.components.find((n) => typeof config.chartResolvers?.[n] === "function");
  if (!name) return;
  let resolved = null;
  try {
    resolved = config.chartResolvers[name](config.resolvedDataDir);
  } catch (err) {
    console.warn(`Warning: chart resolver ${name} failed (${err.message}) - using the chart drawn on the page.`);
    return;
  }
  if (resolved && resolved.kind === "barchart") node.resolvedChart = groupedBarChart(resolved, S);
}
```

3. In the loop over `scene.nodes` (the one that prints notices), add as its first line:

```js
      if (node.kind === "chart") resolveChart(node, config, style);
```

- [ ] **Step 8: Add the deck assertion** — append to `test/render/deck.test.mjs`

```js
test("ExampleChart is a native, editable chart from the page itself when no resolver is configured", deckOpts, async () => {
  const { result, z } = await exportInnotek("innotek-chart.pptx", { chartResolvers: {} });
  assert.equal(result.translatedCharts, 1);
  const chart = z.text("ppt/charts/chart1.xml");
  assert.match(chart, /<c:v>Q1<\/c:v>/);
  assert.match(chart, /<c:numFmt formatCode="0&quot;%&quot;" sourceLinked="0"\/>/);
  assert.match(chart, /<c:legendPos val="b"\/>/);
});
```

- [ ] **Step 9: Run both suites**

Run: `npm test && npm run test:render`
Expected: all pass. The static chart tests prove existing charts are written unchanged.

- [ ] **Step 10: Commit**

```bash
git add src/ooxml/chart.mjs src/layout.mjs src/render/extract.js src/render/scene.mjs src/render/index.mjs test/chart.test.mjs test/render/extract.test.mjs test/render-scene.test.mjs test/render-index.test.mjs test/render/deck.test.mjs
git commit -m "Render mode: Chart.js bar charts become native charts (with number format, legend, gridlines); resolvers still win"
```

---
# Phase 3 — Polish

### Task 10: Native gradients

**Files:**
- Modify: `src/ooxml/xml.mjs`, `src/ooxml/shapes.mjs`, `src/ooxml/index.mjs`, `src/render/extract.js`, `src/render/scene.mjs`
- Test: `test/shapes.test.mjs`, `test/writer.test.mjs`, `test/render/extract.test.mjs`, `test/render-scene.test.mjs`, `test/render/deck.test.mjs`

**Interfaces:**
- Produces:
  - Writer: `gradFillXml(g)` and `fillXml(fill, alpha)` in `xml.mjs`, where `g = { type: "linear", angle (DrawingML degrees, 0 = left→right, 90 = top→bottom) | type: "radial", cx, cy (0..1), stops: [{ pos 0..1, color, alpha? }] }`. A shape's `fill` may be such an object. The background may be `{ gradient: g }`.
  - Scene: a box's `fill` may be `{ gradient, base }`; `background` may be `{ gradient, base }`. Here `gradient = { type: "linear", angle (CSS degrees, 0 = up) | type: "radial", cx, cy, rx (px), stops: [{ pos, color }] }`, and `base` is the color under it (or null).
- DrawingML radial gradients run from the focus point (pos 0) out to the shape's edge (pos 1).

- [ ] **Step 1: Write the failing tests**

In `test/shapes.test.mjs`, add `backgroundXml` to the import from `../src/ooxml/shapes.mjs`, then append:

```js
test("gradient fills: a linear angle, a radial center, alpha stops", () => {
  const lin = shapeXml({ x: 0, y: 0, w: 1, h: 1, fill: { type: "linear", angle: 0, stops: [{ pos: 0, color: "#FF0000" }, { pos: 1, color: "accent1", alpha: 0.5 }] }, paras: [] }, 2, ctx);
  assert.match(lin, /<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"\/><\/a:gs><a:gs pos="100000"><a:schemeClr val="accent1"><a:alpha val="50000"\/><\/a:schemeClr><\/a:gs><\/a:gsLst><a:lin ang="0" scaled="0"\/><\/a:gradFill>/);
  const rad = backgroundXml({ gradient: { type: "radial", cx: 0.5, cy: 0.38, stops: [{ pos: 0, color: "#E5F1E8" }, { pos: 0.6, color: "lt1" }, { pos: 1, color: "lt1" }] } });
  assert.match(rad, /^<p:bg><p:bgPr><a:gradFill rotWithShape="1"><a:gsLst>.*<\/a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="38000" r="50000" b="62000"\/><\/a:path><\/a:gradFill><a:effectLst\/><\/p:bgPr><\/p:bg>$/);
});
```

Append to `test/writer.test.mjs`:

```js
test("gradient slide backgrounds and gradient shape fills pass lint", async () => {
  const gradient = { type: "radial", cx: 0.5, cy: 0.4, stops: [{ pos: 0, color: "#E5F1E8" }, { pos: 1, color: "lt1" }] };
  const slides = [{ layout: "blank", background: { gradient }, shapes: [{ kind: "text", x: 1, y: 1, w: 2, h: 1, fill: { type: "linear", angle: 90, stops: [{ pos: 0, color: "accent1" }, { pos: 1, color: "accent2", alpha: 0 }] }, paras: [] }] }];
  const buffer = await writePptx(slides, { ...DECK, logo: null });
  assert.deepEqual(await lintPackage(buffer), []);
  assert.match((await unzip(buffer)).text("ppt/slides/slide1.xml"), /<p:bg><p:bgPr><a:gradFill/);
});
```

Append to `test/render/extract.test.mjs`:

```js
test("gradients: a radial glow over a color is a native background gradient; a linear gradient fills a box; layered backgrounds stay pictures", opts, async () => {
  const glow = await extractHtml(browser, "<p>x</p>", { layoutStyle: "background:radial-gradient(60% 50% at 50% 38%, rgba(26, 131, 59, 0.12) 0%, transparent 70%), #ffffff" });
  assert.deepEqual(glow.background, {
    gradient: { type: "radial", cx: 0.5, cy: 0.38, rx: 588, stops: [{ pos: 0, color: { hex: "1A833B", alpha: 0.12 } }, { pos: 0.7, color: { hex: "000000", alpha: 0 } }] },
    base: { hex: "FFFFFF", alpha: 1 },
  });
  const lin = await extractHtml(browser, '<div style="width:100px;height:50px;background:linear-gradient(to right, #ff0000, #0000ff)"></div>');
  assert.deepEqual(lin.nodes.find((n) => n.kind === "box").fill, {
    gradient: { type: "linear", angle: 90, stops: [{ pos: 0, color: { hex: "FF0000", alpha: 1 } }, { pos: 1, color: { hex: "0000FF", alpha: 1 } }] },
    base: null,
  });
  const layered = await extractHtml(browser, '<div style="width:100px;height:50px;background:linear-gradient(#f00, #00f), linear-gradient(#0f0, #000)"></div>');
  assert.deepEqual(layered.nodes.map((n) => [n.kind, n.reason]), [["raster", "background layers"]]);
});
```

In the existing test `"backgrounds: the layout root's color is the slide background; a gradient asks for a picture background"` in `test/render/extract.test.mjs`, rename it to `"backgrounds: the layout root's color is the slide background; an image background is a picture"`, and replace its two gradient lines with:

```js
  const g = await extractHtml(browser, "<p>x</p>", { layoutStyle: "background:url(data:image/gif;base64,R0lGODlhAQABAAAAACw=), #fff" });
  assert.deepEqual(g.background, { raster: true, reason: "background image" });
```

Append to `test/render-scene.test.mjs`:

```js
test("gradients: radial backgrounds composite over their base and rescale to the corner radius; linear angles turn into DrawingML's", () => {
  const bg = {
    gradient: { type: "radial", cx: 0.5, cy: 0.5, rx: 100, stops: [{ pos: 0, color: C("1A833B", 0.5) }, { pos: 0.5, color: { hex: "000000", alpha: 0 } }] },
    base: C("FFFFFF"),
  };
  const spec = sceneToSpec(scene([], { canvas: { w: 200, h: 200 }, background: bg }), STYLE);
  assert.deepEqual(spec.background, {
    gradient: { type: "radial", cx: 0.5, cy: 0.5, stops: [{ pos: 0, color: "#8DC19D" }, { pos: Math.round((50 / Math.hypot(100, 100)) * 10000) / 10000, color: "lt1" }, { pos: 1, color: "lt1" }] },
  });
  const node = {
    kind: "box", rect: { x: 0, y: 0, w: 100, h: 50 }, border: null, sides: null, radius: 0, shadow: null, opacity: 1,
    fill: { gradient: { type: "linear", angle: 90, stops: [{ pos: 0, color: C("FF0000") }, { pos: 1, color: C("0000FF") }] }, base: null },
  };
  const [shape] = sceneToSpec(scene([node]), STYLE).shapes;
  assert.deepEqual(shape.fill, { type: "linear", angle: 0, stops: [{ pos: 0, color: "#FF0000" }, { pos: 1, color: "#0000FF" }] });
  assert.equal(shape.fillAlpha, undefined);
});
```

In `test/render/deck.test.mjs`, in the test `"exports the innotek deck with --render: ..."`, replace

```js
  assert.match(z.text("ppt/slides/slide1.xml"), /<p:bg><p:bgPr><a:blipFill/);
```

with

```js
  assert.match(z.text("ppt/slides/slide1.xml"), /<p:bg><p:bgPr><a:gradFill rotWithShape="1">[\s\S]*<a:path path="circle">/);
```

- [ ] **Step 2: Run them to see them fail**

Run: `npm test`, then `npm run test:render`
Expected: the new tests FAIL.

- [ ] **Step 3: Implement the writer**

In `src/ooxml/xml.mjs`, append:

```js
// Linear or radial gradient fill. Stops: [{ pos 0..1, color, alpha? }], at least two. Linear
// angles are DrawingML's (0 = left to right, 90 = top to bottom); a radial gradient runs from
// its focus point (cx, cy as fractions of the shape) at pos 0 out to the shape's edge at pos 1.
export function gradFillXml(g) {
  const pct = (v) => Math.round(Math.min(1, Math.max(0, v)) * 100000);
  const stops = g.stops.map((s) => `<a:gs pos="${pct(s.pos)}">${colorXml(s.color, s.alpha)}</a:gs>`).join("");
  const shade =
    g.type === "radial"
      ? `<a:path path="circle"><a:fillToRect l="${pct(g.cx)}" t="${pct(g.cy)}" r="${pct(1 - g.cx)}" b="${pct(1 - g.cy)}"/></a:path>`
      : `<a:lin ang="${Math.round(((((g.angle ?? 90) % 360) + 360) % 360) * 60000) % 21600000}" scaled="0"/>`;
  return `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst>${shade}</a:gradFill>`;
}

// A shape fill: a color (with optional alpha) or a gradient object.
export function fillXml(fill, alpha) {
  return fill && typeof fill === "object" ? gradFillXml(fill) : solidFill(fill, alpha);
}
```

In `src/ooxml/shapes.mjs`:

1. Change the xml import to `import { XML_DECL, NS_PML, esc, emu, emuSize, solidFill, colorXml, fillXml, gradFillXml } from "./xml.mjs";`
2. In `shapeXml`, replace `solidFill(shape.fill, shape.fillAlpha)` with `fillXml(shape.fill, shape.fillAlpha)`.
3. In `backgroundXml`, replace the `fill` assignment with:

```js
  const fill = rId
    ? `<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="${rId}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`
    : bg.gradient ? gradFillXml(bg.gradient) : solidFill(bg.color, bg.alpha);
```

In `src/ooxml/index.mjs`, change `else if (spec.background?.color) bg = backgroundXml(spec.background);` to:

```js
    else if (spec.background?.color || spec.background?.gradient) bg = backgroundXml(spec.background);
```

- [ ] **Step 4: Implement the extractor** — `src/render/extract.js`

1. Replace `parseColor` with this version (a `keepTransparent` flag for gradient stops):

```js
  // Any CSS color -> { hex: "RRGGBB", alpha }. Fully transparent -> null, unless
  // keepTransparent (gradient stops need their position even when invisible). Computed styles
  // are usually rgb()/rgba(); anything else (oklch(), color()) goes through a 1px canvas.
  function parseColor(css, keepTransparent = false) {
    if (!css || css === "none") return null;
    let r, g, b, a;
    const m = /^rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*(?:,\s*([\d.]+)\s*)?\)$/.exec(String(css).trim());
    if (m) {
      r = +m[1]; g = +m[2]; b = +m[3]; a = m[4] === undefined ? 1 : +m[4];
    } else {
      probe.clearRect(0, 0, 1, 1);
      probe.fillStyle = "rgba(0, 0, 0, 0)";
      probe.fillStyle = css;
      probe.fillRect(0, 0, 1, 1);
      const d = probe.getImageData(0, 0, 1, 1).data;
      r = d[0]; g = d[1]; b = d[2]; a = d[3] / 255;
    }
    if (!(a > 0.001)) return keepTransparent ? { hex: "000000", alpha: 0 } : null;
    const hex = [r, g, b].map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, "0")).join("").toUpperCase();
    return { hex, alpha: Math.round(a * 1000) / 1000 };
  }
```

2. Add after `firstFamily`:

```js
  const SIDE_ANGLES = {
    top: 0, right: 90, bottom: 180, left: 270,
    "top right": 45, "right top": 45, "bottom right": 135, "right bottom": 135,
    "bottom left": 225, "left bottom": 225, "top left": 315, "left top": 315,
  };

  function toDeg(a) {
    const v = parseFloat(a);
    if (a.endsWith("turn")) return v * 360;
    if (a.endsWith("grad")) return v * 0.9;
    if (a.endsWith("rad")) return (v * 180) / Math.PI;
    return v;
  }

  // "rgba(26, 131, 59, 0.12) 0%" -> { color, pos: "0%" | null }; null if it isn't a color stop.
  function parseColorStop(p) {
    const m = /^((?:rgba?|hsla?|hwb|lab|lch|oklab|oklch|color)\([^)]*\)|#[0-9a-f]{3,8}|[a-z]+)(?:\s+(-?[\d.]+(?:%|px)))?/i.exec(String(p).trim());
    if (!m || !CSS.supports("color", m[1])) return null;
    return { color: parseColor(m[1], true), pos: m[2] ?? null };
  }

  // One CSS gradient layer, measured against the element's box (rc), ->
  // { type: "linear", angle } | { type: "radial", cx, cy, rx }, plus stops [{ pos, color }].
  // Returns null for repeating or unparseable gradients.
  function parseGradient(layer, rc) {
    const m = /^(repeating-)?(linear|radial)-gradient\((.*)\)$/s.exec(layer.trim());
    if (!m || m[1]) return null;
    const parts = splitTop(m[3]);
    const g = { type: m[2] };
    let length;
    if (g.type === "linear") {
      let angle = 180;
      if (/^-?[\d.]+(deg|turn|grad|rad)$/.test(parts[0])) angle = toDeg(parts.shift());
      else if (/^to\s/.test(parts[0])) angle = SIDE_ANGLES[parts.shift().slice(3).trim()] ?? 180;
      g.angle = r2(((angle % 360) + 360) % 360);
      const rad = (g.angle * Math.PI) / 180;
      length = Math.abs(rc.w * Math.sin(rad)) + Math.abs(rc.h * Math.cos(rad));
    } else {
      const shape = parseColorStop(parts[0]) ? "" : parts.shift();
      const [sizePart, atPart] = shape.split(/\s+at\s+/);
      let [px, py] = (atPart || "center").trim().split(/\s+/);
      if (px === "top" || px === "bottom") [px, py] = [py || "center", px];
      const frac = (v, size) => {
        if (v === undefined || v === "center") return 0.5;
        if (v === "left" || v === "top") return 0;
        if (v === "right" || v === "bottom") return 1;
        return v.endsWith("%") ? parseFloat(v) / 100 : parseFloat(v) / size;
      };
      g.cx = r2(frac(px, rc.w));
      g.cy = r2(frac(py, rc.h));
      const cx = g.cx * rc.w;
      const cy = g.cy * rc.h;
      const toSides = [cx, rc.w - cx, cy, rc.h - cy];
      const toCorners = [[0, 0], [rc.w, 0], [0, rc.h], [rc.w, rc.h]].map(([x, y]) => Math.hypot(x - cx, y - cy));
      const tokens = (sizePart || "").trim().split(/\s+/).filter((t) => t && t !== "circle" && t !== "ellipse");
      const size = tokens[0];
      let rx;
      if (!size || size === "farthest-corner") rx = Math.max(...toCorners);
      else if (size === "closest-side") rx = Math.min(...toSides);
      else if (size === "farthest-side") rx = Math.max(...toSides);
      else if (size === "closest-corner") rx = Math.min(...toCorners);
      else rx = size.endsWith("%") ? (parseFloat(size) / 100) * rc.w : parseFloat(size);
      g.rx = r2(rx);
      length = rx;
    }
    const stops = [];
    for (const p of parts) {
      const s = parseColorStop(p);
      if (!s) return null;
      stops.push({ color: s.color, pos: s.pos == null ? null : s.pos.endsWith("%") ? parseFloat(s.pos) / 100 : parseFloat(s.pos) / length });
    }
    if (stops.length < 2) return null;
    // Missing positions: first 0, last 1, the rest spread evenly between their neighbours.
    if (stops[0].pos == null) stops[0].pos = 0;
    if (stops[stops.length - 1].pos == null) stops[stops.length - 1].pos = 1;
    for (let i = 1; i < stops.length - 1; i++) {
      if (stops[i].pos != null) continue;
      let j = i;
      while (stops[j].pos == null) j++;
      const a = stops[i - 1].pos;
      const b = stops[j].pos;
      for (let k = i; k < j; k++) stops[k].pos = a + ((b - a) * (k - i + 1)) / (j - i + 1);
    }
    g.stops = stops.map((s) => ({ pos: Math.round(s.pos * 10000) / 10000, color: s.color }));
    return g;
  }
```

3. Replace `backgroundImageOf` with:

```js
    // A single background layer -> { gradient } | { src, format, fit } | { unsupported }.
    function backgroundImageOf(cs, rc) {
      const bg = cs.backgroundImage;
      if (!bg || bg === "none") return null;
      // A `background` shorthand ending in a color reports an extra "none" image layer.
      const layers = splitTop(bg).filter((l) => l !== "none");
      if (layers.length > 1) return { unsupported: "background layers" };
      if (/gradient\(/.test(layers[0])) {
        const gradient = parseGradient(layers[0], rc);
        return gradient ? { gradient } : { unsupported: "gradient" };
      }
      const m = /^url\("?(.*?)"?\)$/.exec(layers[0]);
      if (!m) return { unsupported: "background image" };
      const src = new URL(m[1], location.href).href;
      const format = formatOf(src);
      if (!format || format === "svg") return { unsupported: "background image" };
      const size = cs.backgroundSize;
      return { src, format, fit: size === "cover" ? "cover" : size === "contain" ? "contain" : "fill" };
    }
```

4. In `rasterReason`, change its signature to `function rasterReason(el, cs, tag, op, rc)`, and change `const bg = backgroundImageOf(cs);` to `const bg = backgroundImageOf(cs, rc);`. In `walk`, change the call to `rasterReason(el, cs, tag, op, rc)`.

5. Replace `boxNode` with:

```js
    function boxNode(el, cs, rc, rotate, op) {
      if (el === pageEl || !onCanvas(rc)) return;
      const d = decoration(cs, rc);
      const bg = backgroundImageOf(cs, rc);
      // A gradient fills the box, over its background color.
      if (bg && bg.gradient) d.fill = { gradient: bg.gradient, base: d.fill };
      if (d.fill || d.border || d.sides || d.shadow) nodes.push({ kind: "box", rect: rc, ...d, opacity: op, rotate });
      if (bg && bg.src) {
        nodes.push({ kind: "image", id: mark(el), rect: rc, src: bg.src, format: bg.format, fit: bg.fit, natural: null, alt: "", border: null, shadow: null, opacity: op, rotate });
      }
    }
```

6. In `backgroundOf`, replace the block that starts `if (cs.backgroundImage && cs.backgroundImage !== "none") {` (through its closing brace) with:

```js
        if (cs.backgroundImage && cs.backgroundImage !== "none") {
          const bg = backgroundImageOf(cs, rectOf(e.getBoundingClientRect()));
          if (bg && bg.gradient) return { gradient: bg.gradient, base: parseColor(cs.backgroundColor) };
          return { raster: true, reason: bg && bg.unsupported === "gradient" ? "gradient background" : "background image" };
        }
```

- [ ] **Step 5: Implement the mapping** — `src/render/scene.mjs`

1. Add above `boxShapes`:

```js
// A CSS gradient (and the color under it) -> a DrawingML gradient. Stops are composited over an
// opaque base color; transparent stops borrow their neighbour's hue (a fade to "transparent"
// would otherwise pass through black); radial stop positions are rescaled from the CSS radius
// to DrawingML's focus-to-farthest-corner radius; angles move from CSS (0 = up) to DrawingML
// (0 = left to right).
function gradientOf(g, base, rect, style, opacity = 1) {
  let stops = g.stops.map((s) => ({ pos: s.pos, color: s.color }));
  stops = stops.map((s, i) => {
    if (s.color.alpha > 0) return s;
    const near = stops.slice(i + 1).concat(stops.slice(0, i).reverse()).find((x) => x.color.alpha > 0);
    return near ? { pos: s.pos, color: { hex: near.color.hex, alpha: 0 } } : s;
  });
  if (base && (base.alpha ?? 1) >= 1) {
    const over = (c) => {
      const a = c.alpha ?? 1;
      const ch = (i) => Math.round(parseInt(c.hex.slice(i, i + 2), 16) * a + parseInt(base.hex.slice(i, i + 2), 16) * (1 - a));
      return { hex: [0, 2, 4].map((i) => ch(i).toString(16).padStart(2, "0")).join("").toUpperCase(), alpha: 1 };
    };
    stops = stops.map((s) => ({ pos: s.pos, color: over(s.color) }));
  }
  let out;
  if (g.type === "radial") {
    const cx = g.cx * rect.w;
    const cy = g.cy * rect.h;
    const far = Math.max(...[[0, 0], [rect.w, 0], [0, rect.h], [rect.w, rect.h]].map(([x, y]) => Math.hypot(x - cx, y - cy)));
    const k = far > 0 ? g.rx / far : 1;
    stops = stops.map((s) => ({ ...s, pos: Math.min(1, s.pos * k) }));
    out = { type: "radial", cx: g.cx, cy: g.cy };
  } else {
    out = { type: "linear", angle: (((g.angle - 90) % 360) + 360) % 360 };
  }
  const last = stops[stops.length - 1];
  if (last.pos < 1) stops.push({ pos: 1, color: last.color });
  out.stops = stops.map((s) => {
    const stop = { pos: round(s.pos), color: colorOf(s.color, style) };
    const a = alphaOf(s.color, opacity);
    if (a !== undefined) stop.alpha = a;
    return stop;
  });
  return out;
}
```

2. In `boxShapes`, replace

```js
    if (node.fill) {
      shape.fill = colorOf(node.fill, style);
```

with

```js
    if (node.fill?.gradient) shape.fill = gradientOf(node.fill.gradient, node.fill.base, node.rect, style, op);
    else if (node.fill) {
      shape.fill = colorOf(node.fill, style);
```

3. Change `backgroundOf(bg, style)` to `backgroundOf(bg, style, canvas)`. Add as its third line, after `if (bg.png) ...`:

```js
  if (bg.gradient) return { gradient: gradientOf(bg.gradient, bg.base, { x: 0, y: 0, w: canvas.w, h: canvas.h }, style) };
```

In `sceneToSpec`, change the call to `backgroundOf(scene.background, style, scene.canvas)`.

- [ ] **Step 6: Run both suites**

Run: `npm test && npm run test:render`
Expected: all pass.

Then look at the cover. Export the innotek deck once with `render: true` (as `exportInnotek` in `test/render/deck.test.mjs` does) to `out/render-check.pptx`, convert it with `soffice --headless --convert-to pdf --outdir out out/render-check.pptx`, and look at page 1. The glow must be a soft light-green ellipse behind the title, fading to white towards the edges. LibreOffice draws stop 0 at the focus point; this was verified while writing this plan.

- [ ] **Step 7: Commit**

```bash
git add src/ooxml/xml.mjs src/ooxml/shapes.mjs src/ooxml/index.mjs src/render/extract.js src/render/scene.mjs test/shapes.test.mjs test/writer.test.mjs test/render/extract.test.mjs test/render-scene.test.mjs test/render/deck.test.mjs
git commit -m "Render mode: CSS gradients become native gradient fills and slide backgrounds"
```

---

### Task 11: SVG images, vector in PowerPoint

**Files:**
- Modify: `src/ooxml/package.mjs`, `src/ooxml/shapes.mjs`, `src/ooxml/index.mjs`, `src/render/extract.js`, `src/render/capture.mjs`, `src/render/scene.mjs`
- Test: `test/writer.test.mjs`, `test/render/extract.test.mjs`, `test/render-scene.test.mjs`, `test/render/deck.test.mjs`

**Interfaces:**
- Produces:
  - An image shape may carry `svg: Buffer` next to its PNG `data`. The writer embeds both, using the `asvg:svgBlip` extension.
  - A scene image node with `format: "svg"` gets `svg` (the file bytes) and `data` (a 2× PNG rendered by the browser at the element's size), from capture.

- [ ] **Step 1: Write the failing tests**

Append to `test/writer.test.mjs`:

```js
test("SVG pictures carry the SVG next to a PNG fallback", async () => {
  const svg = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="4" height="2"><rect width="4" height="2" fill="#1A833B"/></svg>');
  const buffer = await writePptx([{ layout: "blank", shapes: [{ kind: "image", x: 1, y: 1, w: 2, h: 1, data: readFileSync(PIC), ext: "png", svg, alt: "Logo" }] }], { ...DECK, logo: null });
  assert.deepEqual(await lintPackage(buffer), []);
  const z = await unzip(buffer);
  assert.ok(z.names.includes("ppt/media/image2.svg"));
  assert.match(z.text("[Content_Types].xml"), /<Default Extension="svg" ContentType="image\/svg\+xml"\/>/);
  assert.match(z.text("ppt/slides/slide1.xml"), /<a:blip r:embed="rId2"><a:extLst><a:ext uri="\{96DAC541-7B7A-43D3-8B79-37D633B846F1\}"><asvg:svgBlip xmlns:asvg="http:\/\/schemas.microsoft.com\/office\/drawing\/2016\/SVG\/main" r:embed="rId3"\/><\/a:ext><\/a:extLst><\/a:blip>/);
});
```

Append to `test/render/extract.test.mjs`:

```js
test("an SVG image is an image node, not a picture of an element", opts, async () => {
  const svg = `data:image/svg+xml;charset=utf-8,${encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20"><rect width="40" height="20" fill="#1a833b"/></svg>')}`;
  const s = await extractHtml(browser, `<img src="${svg}" alt="Logo" style="display:block;width:80px;height:40px">`);
  assert.deepEqual(s.nodes.map((n) => [n.kind, n.format, n.alt]), [["image", "svg", "Logo"]]);
});
```

Append to `test/render-scene.test.mjs`:

```js
test("an SVG image keeps its SVG bytes next to the PNG", () => {
  const svg = Buffer.from("<svg/>");
  const png = Buffer.from("png");
  const node = { kind: "image", id: 1, rect: { x: 0, y: 0, w: 80, h: 40 }, src: "/logo.svg", format: "svg", fit: "fill", natural: null, alt: "Logo", border: null, shadow: null, opacity: 1, data: png, ext: "png", svg };
  const [shape] = sceneToSpec(scene([node]), STYLE).shapes;
  assert.equal(shape.svg, svg);
  assert.equal(shape.data, png);
});
```

- [ ] **Step 2: Run them to see them fail**

Run: `npm test`, then `npm run test:render`
Expected: the new tests FAIL.

- [ ] **Step 3: Implement the writer**

1. In `src/ooxml/package.mjs`, add `svg: "image/svg+xml",` to `MEDIA_TYPES`.
2. In `src/ooxml/shapes.mjs`, add above `pictureXml`:

```js
// PowerPoint 2016+ draws the SVG; older PowerPoint, Google Slides, Keynote and LibreOffice use the PNG.
const SVG_EXT = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
```

Then in `pictureXml`, add `svgRId` to the destructured parameters, and replace `<a:blip r:embed="${rId}"/>` with `${blip}`. Define `blip` above the `return`:

```js
  const blip = svgRId
    ? `<a:blip r:embed="${rId}"><a:extLst><a:ext uri="${SVG_EXT}"><asvg:svgBlip xmlns:asvg="${ASVG}" r:embed="${svgRId}"/></a:ext></a:extLst></a:blip>`
    : `<a:blip r:embed="${rId}"/>`;
```

3. In `src/ooxml/index.mjs`, replace the `case "image":` block with:

```js
    case "image": {
      const rId = pkg.rel(part, media.add(shape.path ?? { data: shape.data, ext: shape.ext }), REL.image);
      const svgRId = shape.svg ? pkg.rel(part, media.add({ data: shape.svg, ext: "svg" }), REL.image) : null;
      return pictureXml({ id, name: `Picture ${id}`, descr: shape.alt, rId, svgRId, box: shape, crop: shape.crop, line: shape.line, shadow: shape.shadow });
    }
```

- [ ] **Step 4: Implement render mode**

In `src/render/extract.js`:

1. In `rasterReason`, delete the line `if (fmt === "svg") return "SVG image";`.
2. In `imageNode`, replace `if (!el.complete || !el.naturalWidth) {` with:

```js
      const svg = formatOf(src) === "svg";
      // An SVG without width/height has no natural size but still draws.
      if (!el.complete || (!el.naturalWidth && !svg)) {
```

In `src/render/capture.mjs`:

1. Add above `loadImage`:

```js
// An SVG drawn at the element's size (2x) into a transparent PNG, the fallback next to the SVG.
async function svgPng(page, src, rect) {
  const dataUrl = await page.evaluate(
    async ({ url, w, h }) => {
      const img = new Image();
      img.src = url;
      await img.decode();
      const c = document.createElement("canvas");
      c.width = Math.max(1, Math.round(w * 2));
      c.height = Math.max(1, Math.round(h * 2));
      c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
      return c.toDataURL("image/png");
    },
    { url: src, w: rect.w, h: rect.h },
  );
  return Buffer.from(dataUrl.slice(dataUrl.indexOf(",") + 1), "base64");
}
```

2. In `loadImage`, directly after `node.data = await fetchBytes(page, node.src);`, add:

```js
    if (node.format === "svg") {
      node.svg = node.data;
      node.data = await svgPng(page, node.src, node.rect);
      node.ext = "png";
      node.natural = null;
      node.fit = "fill";
      return;
    }
```

In `src/render/scene.mjs`, in `imageShape`, add after the `if (crop) shape.crop = crop;` line:

```js
  if (node.svg) shape.svg = node.svg;
```

- [ ] **Step 5: Add the deck assertion** — append to `test/render/deck.test.mjs`

```js
test("the SVG logo is exported as an SVG picture with a PNG fallback", deckOpts, async () => {
  const { z } = await exportInnotek("innotek-svg.pptx");
  assert.ok(z.names.some((n) => /^ppt\/media\/image\d+\.svg$/.test(n)));
  assert.match(z.text("ppt/slides/slide2.xml"), /<asvg:svgBlip /);
});
```

- [ ] **Step 6: Run both suites**

Run: `npm test && npm run test:render`
Expected: all pass.

- [ ] **Step 7: Commit**

```bash
git add src/ooxml/package.mjs src/ooxml/shapes.mjs src/ooxml/index.mjs src/render/extract.js src/render/capture.mjs src/render/scene.mjs test/writer.test.mjs test/render/extract.test.mjs test/render-scene.test.mjs test/render/deck.test.mjs
git commit -m "Render mode: SVG images stay vector in PowerPoint, with a PNG fallback"
```

---

### Task 12: Repeated elements go on the slide master

**Files:**
- Modify: `src/ooxml/master.mjs`, `src/ooxml/index.mjs`, `src/render/index.mjs`, `src/index.mjs`
- Test: `test/writer.test.mjs`, `test/render-index.test.mjs`, `test/render/deck.test.mjs`

**Interfaces:**
- Produces:
  - `deck.masterShapes`: extra text or image shapes drawn on the slide master.
  - `renderSpecs` also returns `masterShapes`.
  - Exports from `src/render/index.mjs`: `nodeSignature(node) → string | null` and `repeatedSignatures(scenes) → Set<string>`.
- A rendered element is "repeated" when a node with the same kind, the same rect (to 0.5 px) and the same content appears on every successfully rendered slide, and there are at least two such slides. Content means the bytes (image or raster), the runs (text), or the style (box). Title text never counts.
- Notice for a repeated picture: `Notice: slide master: <label> exported as a picture (<reason>).`

- [ ] **Step 1: Write the failing tests**

Append to `test/writer.test.mjs`:

```js
test("master shapes are drawn on the slide master, with their media related from it", async () => {
  const masterShapes = [
    { kind: "image", x: 0.5, y: 6.9, w: 1, h: 0.3, data: readFileSync(PIC), ext: "png", alt: "Logo" },
    { kind: "text", x: 11, y: 7, w: 2, h: 0.3, geometry: "rect", anchor: "t", inset: 0, paras: [{ runs: [{ text: "Confidential" }] }] },
  ];
  const buffer = await writePptx([{ layout: "blank", shapes: [] }], { ...DECK, logo: null, masterShapes });
  assert.deepEqual(await lintPackage(buffer), []);
  const z = await unzip(buffer);
  const master = z.text("ppt/slideMasters/slideMaster1.xml");
  assert.match(master, /<p:cNvPr id="5" name="Picture 5" descr="Logo"\/>/);
  assert.match(master, /<a:t>Confidential<\/a:t>/);
  assert.match(z.text("ppt/slideMasters/_rels/slideMaster1.xml.rels"), /Target="\.\.\/media\/image1\.png"/);
  assert.doesNotMatch(z.text("ppt/slides/slide1.xml"), /Confidential/);
});
```

Append to `test/render-index.test.mjs`:

```js
test("an element drawn identically on every slide moves to the slide master", async () => {
  const logo = () => ({ kind: "image", id: 9, rect: { x: 48, y: 498, w: 95, h: 26 }, src: "/logo.png", format: "png", fit: "fill", natural: null, alt: "Logo", border: null, shadow: null, opacity: 1, data: Buffer.from("same"), ext: "png" });
  const only = { ...logo(), rect: { x: 300, y: 300, w: 20, h: 20 } };
  const { args } = setup("# A\n\n---\n\n# B\n", { slides: [scene(1, 0, "A", [logo(), only]), scene(2, 1, "B", [logo()])] });
  const { result } = await withWarnings(() => renderSpecs(args));
  assert.equal(result.masterShapes.length, 1);
  assert.equal(result.masterShapes[0].alt, "Logo");
  assert.deepEqual(result.specs.map((s) => s.shapes.length), [1, 0]);
  const single = setup("# A\n", { slides: [scene(1, 0, "A", [logo()])] });
  assert.deepEqual((await withWarnings(() => renderSpecs(single.args))).result.masterShapes, []);
});
```

Append to `test/render/deck.test.mjs`:

```js
test("the footer logo is drawn once, on the slide master", deckOpts, async () => {
  const { z } = await exportInnotek("innotek-master.pptx");
  assert.match(z.text("ppt/slideMasters/slideMaster1.xml"), /<asvg:svgBlip /);
  for (let i = 1; i <= 7; i++) assert.doesNotMatch(z.text(`ppt/slides/slide${i}.xml`), /<asvg:svgBlip /, `slide ${i}`);
});
```

In the test `"the SVG logo is exported as an SVG picture with a PNG fallback"` (Task 11), replace `z.text("ppt/slides/slide2.xml")` with `z.text("ppt/slideMasters/slideMaster1.xml")`.

- [ ] **Step 2: Run them to see them fail**

Run: `npm test`, then `npm run test:render`
Expected: the new tests FAIL.

- [ ] **Step 3: Implement the writer**

1. In `src/ooxml/master.mjs`, change the signature to `export function masterXml({ geometry: g, ctx, layoutIds, logo, extra = "" })`, and change `${title}${body}${pic}` to `${title}${body}${pic}${extra}`.
2. In `src/ooxml/index.mjs`, replace the line `pkg.addXml(MASTER, CT.slideMaster, masterXml({ geometry: deck.geometry, ctx, layoutIds, logo }));` with:

```js
  // Render mode can put elements every slide shares (a footer logo) on the master. Their ids
  // follow the master's own title (2), body (3) and logo (4).
  const masterEnv = { pkg, part: MASTER, media, ctx, counters };
  let extra = "";
  let masterShapeId = 5;
  for (const shape of deck.masterShapes || []) extra += await shapeToXml(shape, masterShapeId++, masterEnv);
  pkg.addXml(MASTER, CT.slideMaster, masterXml({ geometry: deck.geometry, ctx, layoutIds, logo, extra }));
```

- [ ] **Step 4: Implement render mode** — `src/render/index.mjs`

1. Add imports: `import { createHash } from "node:crypto";`, and change the scene import to `import { sceneToSpec, shapesOf, unitsFor } from "./scene.mjs";`.
2. Add above `renderSpecs`:

```js
const sha1 = (b) => createHash("sha1").update(b).digest("hex");

// What makes two nodes "the same element": kind, position (to 0.5 px) and content. Titles
// never count - they differ per slide even when they sit in the same place.
export function nodeSignature(n) {
  if (!n.rect) return null;
  const r = [n.rect.x, n.rect.y, n.rect.w, n.rect.h].map((v) => Math.round(v * 2) / 2).join(",");
  if (n.kind === "image" || n.kind === "raster") {
    const bytes = n.png || n.data;
    return bytes ? `${n.kind}|${r}|${sha1(bytes)}|${n.svg ? sha1(n.svg) : ""}` : null;
  }
  if (n.kind === "text" && n.role !== "title") return `text|${r}|${JSON.stringify(n.runs)}`;
  if (n.kind === "box") return `box|${r}|${JSON.stringify([n.fill, n.border, n.sides, n.radius, n.shadow])}`;
  return null;
}

// Signatures present on every scene (when there are at least two).
export function repeatedSignatures(scenes) {
  if (scenes.length < 2) return new Set();
  const counts = new Map();
  for (const s of scenes) {
    for (const sig of new Set(s.nodes.map(nodeSignature).filter(Boolean))) counts.set(sig, (counts.get(sig) || 0) + 1);
  }
  return new Set([...counts].filter(([, n]) => n === scenes.length).map(([sig]) => sig));
}
```

3. In `renderSpecs`, directly before `const specs = [];`, add:

```js
  // Elements every slide shares go on the slide master once.
  const rendered = captured.slides.filter((s) => !s.error);
  const repeated = repeatedSignatures(rendered);
  let masterShapes = [];
  if (repeated.size) {
    const first = rendered[0];
    const shared = first.nodes.filter((n) => repeated.has(nodeSignature(n)));
    for (const n of shared) if (n.kind === "raster") console.warn(`Notice: slide master: ${n.label} exported as a picture (${n.reason}).`);
    masterShapes = shared.flatMap((n) => shapesOf(n, unitsFor(first, style.geometry), sceneStyle));
    for (const s of rendered) s.nodes = s.nodes.filter((n) => !repeated.has(nodeSignature(n)));
  }
```

4. Change the `return` to `return { specs, style, screenshots: shots, translatedCharts, masterShapes };`.

- [ ] **Step 5: Pass master shapes to the writer** — `src/index.mjs`

In the `writePptx(specs, { ... })` options, add after the `logo:` line:

```js
    masterShapes: rendered ? rendered.masterShapes : [],
```

- [ ] **Step 6: Run both suites**

Run: `npm test && npm run test:render`
Expected: all pass.

- [ ] **Step 7: Commit**

```bash
git add src/ooxml/master.mjs src/ooxml/index.mjs src/render/index.mjs src/index.mjs test/writer.test.mjs test/render-index.test.mjs test/render/deck.test.mjs
git commit -m "Render mode: elements drawn identically on every slide go on the slide master"
```

---

### Task 13: Visual check, docs, version

**Files:**
- Create: `src/render/compare.mjs`, `test/render/compare.test.mjs`
- Modify: `scripts/render.mjs`, `README.md`, `package.json`, `test/render/deck.test.mjs`

**Interfaces:**
- Consumes: `convert({ render: { screenshots: true } })` (Task 6), `loadChromium`, `slidevCli`.
- Produces:
  - `similarity(page, pngA, pngB) → Promise<number>`: 1 minus the mean absolute difference of 160×90 grayscale thumbnails, in 0..1.
  - `npm run render -- --render`.

- [ ] **Step 1: Write the failing tests**

`test/render/compare.test.mjs`:

```js
import { test, after } from "node:test";
import assert from "node:assert/strict";
import { usableChromium } from "./helpers.mjs";
import { similarity } from "../../src/render/compare.mjs";

const chromium = await usableChromium();
const browser = chromium ? await chromium.launch() : null;
after(() => browser?.close());
const opts = { skip: browser ? false : "Playwright/Chromium not available" };

test("identical images score 1; black against white scores 0", opts, async () => {
  const page = await browser.newPage({ viewport: { width: 320, height: 180 } });
  const shot = async (color) => {
    await page.setContent(`<body style="margin:0;background:${color}"></body>`);
    return page.screenshot();
  };
  const white = await shot("#ffffff");
  const black = await shot("#000000");
  assert.equal(await similarity(page, white, white), 1);
  assert.ok(Math.abs(await similarity(page, white, black)) < 0.001);
  await page.close();
});
```

Append to `test/render/deck.test.mjs`:

```js
test("the innotek deck needs no picture fallbacks", deckOpts, async () => {
  const notices = [];
  const orig = console.warn;
  console.warn = (...a) => notices.push(a.join(" "));
  try {
    await exportInnotek("innotek-no-pictures.pptx");
  } finally {
    console.warn = orig;
  }
  assert.deepEqual(notices.filter((n) => / exported as a picture /.test(n)), []);
});
```

- [ ] **Step 2: Run them to see them fail**

Run: `npm run test:render`
Expected: `compare.test.mjs` FAILS (the module is missing). The "no picture fallbacks" test may already pass, or may name an element. If it names one, find out why that element isn't native, and fix it in the task that owns it (extract or scene) before going on.

- [ ] **Step 3: Implement** — `src/render/compare.mjs`

```js
// Visual check for render mode: how alike two slide images are, as 1 minus the mean absolute
// difference of their 160x90 grayscale thumbnails (1 = identical, 0 = black vs white). Decoded
// and compared inside a Playwright page, so no image library is needed.

export async function similarity(page, a, b) {
  return page.evaluate(
    async ({ a: pngA, b: pngB }) => {
      const load = async (b64) => {
        const img = new Image();
        img.src = `data:image/png;base64,${b64}`;
        await img.decode();
        return img;
      };
      const gray = (img) => {
        const c = document.createElement("canvas");
        c.width = 160;
        c.height = 90;
        const x = c.getContext("2d");
        x.fillStyle = "#ffffff";
        x.fillRect(0, 0, 160, 90);
        x.drawImage(img, 0, 0, 160, 90);
        const d = x.getImageData(0, 0, 160, 90).data;
        const out = new Float64Array(160 * 90);
        for (let i = 0; i < out.length; i++) out[i] = 0.299 * d[i * 4] + 0.587 * d[i * 4 + 1] + 0.114 * d[i * 4 + 2];
        return out;
      };
      const ga = gray(await load(pngA));
      const gb = gray(await load(pngB));
      let sum = 0;
      for (let i = 0; i < ga.length; i++) sum += Math.abs(ga[i] - gb[i]);
      return 1 - sum / ga.length / 255;
    },
    { a: a.toString("base64"), b: b.toString("base64") },
  );
}
```

- [ ] **Step 4: Add `--render` to the render check** — `scripts/render.mjs`

1. Update the header comment's usage line to `// isn't installed. Usage: node scripts/render.mjs [--render] [path/to/slides.md ...]` and add:

```js
// --render exports in render mode and adds a visual check: each slide's LibreOffice render is
// compared with Slidev's own screenshot of it, and a slide scoring below 0.90 fails.
```

2. Change the fs import to `import { execFileSync, spawnSync } from "node:child_process";` (unchanged) and `import { existsSync, mkdirSync, rmSync, readFileSync, readdirSync, writeFileSync } from "node:fs";`, and add:

```js
import { loadChromium } from "../src/render/browser.mjs";
import { slidevCli } from "../src/render/server.mjs";
import { similarity } from "../src/render/compare.mjs";

const VISUAL_THRESHOLD = 0.9;

// LibreOffice's slide-N.png pages against Slidev's screenshots (null = a slide that fell back
// to the markdown export, not compared).
async function visualScores(chromium, outDir, screenshots) {
  const pages = readdirSync(outDir).filter((f) => /^slide-\d+\.png$/.test(f)).sort((a, b) => parseInt(a.slice(6), 10) - parseInt(b.slice(6), 10));
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage();
    const scores = [];
    for (const [i, file] of pages.entries()) {
      if (!screenshots[i]) continue;
      writeFileSync(join(outDir, `slidev-${i + 1}.png`), screenshots[i]);
      scores.push({ slide: i + 1, score: await similarity(page, screenshots[i], readFileSync(join(outDir, file))) });
    }
    return scores;
  } finally {
    await browser.close();
  }
}
```

3. In `main()`, replace the line `const args = process.argv.slice(2).map((p) => resolve(p));` with:

```js
  const argv = process.argv.slice(2);
  const render = argv.includes("--render");
  const args = argv.filter((a) => a !== "--render").map((p) => resolve(p));
  let chromium = null;
  if (render) {
    try {
      chromium = await loadChromium(ROOT);
    } catch (err) {
      console.log(`${err.message} - skipping the render-mode check.`);
      return;
    }
  }
```

4. Inside the `for (const deck of decks)` loop, at its top, add:

```js
    if (render && !slidevCli(dirname(deck))) {
      console.log(`skip ${basename(dirname(deck))}: not a Slidev deck with @slidev/cli installed`);
      continue;
    }
```

5. Change the `convert(...)` call to pass `render: render ? { screenshots: true } : undefined`.

6. Replace from `const ok = pages === result.slideCount && missing.length === 0;` to the `console.log(...)` line with:

```js
    let ok = pages === result.slideCount && missing.length === 0;
    let detail = missing.length ? `, titles missing from the render: ${missing.join(" | ")}` : "";
    if (render) {
      const scores = await visualScores(chromium, outDir, result.screenshots);
      const low = scores.filter((s) => s.score < VISUAL_THRESHOLD);
      ok &&= low.length === 0;
      detail += `, visual ${scores.map((s) => s.score.toFixed(2)).join(" ")}`;
      if (low.length) detail += ` (below ${VISUAL_THRESHOLD}: slide ${low.map((s) => s.slide).join(", ")})`;
    }
    failed ||= !ok;
    console.log(`${ok ? "ok  " : "FAIL"} ${name}: ${pages}/${result.slideCount} pages${detail} -> ${outDir}`);
```

- [ ] **Step 5: Run the visual check**

Run: `npm run render -- --render`
Expected: the kitchen-sink fixture is skipped (it isn't a Slidev deck), and `ok   innotek-SlideDevBase: 7/7 pages, visual …`, with every score ≥ 0.90.

If a slide scores below the threshold, compare `out/render/innotek-SlideDevBase/slidev-N.png` with `slide-N.png`. Fix the cause in the extractor or the mapping; don't lower the threshold. The threshold may only be tuned once, on this deck, with the reason written next to `VISUAL_THRESHOLD`.

- [ ] **Step 6: Docs and version**

1. `package.json`: set `"version": "0.3.0"`.
2. `README.md`, in the "Render mode" section, replace the bullet list with:

```markdown
- Text keeps its rendered font, size, weight, color, letter spacing and position; slide
  titles still go into title placeholders.
- Boxes with a background, border, rounded corners, shadow or gradient (cards, callouts)
  become native shapes; lists get native bullets in their rendered colors; tables become
  native tables with each cell's own fill and borders; code blocks keep their syntax colors.
- Chart.js bar charts (e.g. through vue-chartjs) become native, editable charts with the
  page's data, colors, axis range and number format. A `chartResolvers` entry for the
  component still wins.
- Images become native pictures; SVG images stay vector in PowerPoint 2016+ (with a PNG
  fallback elsewhere). Elements drawn identically on every slide, like a footer logo, go on
  the slide master once.
- Anything PowerPoint can't express natively (blur and other CSS filters, masks, inline SVG
  drawings, other canvases) is drawn as a picture of just that element, with a notice
  naming the slide.
- `colors` and `fonts` in `pptx.config.mjs` still build the theme (rendered colors that match
  become theme colors); `logo`, `eyebrowClass` and `cssVarColors` aren't needed.
  `render: { timeout }` sets the per-slide wait (default 15000 ms).
- A slide that can't be rendered falls back to the markdown export, with a warning.
```

3. `README.md`, inside the code block of the "Development" section, add these two lines:

```bash
npm run test:render         # browser tests (needs playwright-chromium; skips otherwise)
npm run render -- --render  # render-mode export + visual check against Slidev's own screenshots
```

- [ ] **Step 7: Run everything**

Run: `npm test && npm run test:render && npm run render && npm run render -- --render`
Expected: all pass. The static render check is unchanged (`ok` for both decks), and the render-mode check prints `ok` with the visual scores.

- [ ] **Step 8: Commit**

```bash
git add src/render/compare.mjs test/render/compare.test.mjs test/render/deck.test.mjs scripts/render.mjs README.md package.json
git commit -m "Render mode: visual check against Slidev's own screenshots; docs; 0.3.0"
```
