# Own OOXML Writer Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace pptxgenjs with an in-repo OOXML writer that emits PowerPoint-native decks (real placeholders, theme, notes, table styles, editable charts) that open cleanly in PowerPoint desktop/Online, Google Slides, Keynote and LibreOffice.

**Architecture:** `parse.mjs` (markdown → blocks, unchanged contract) → `layout.mjs` (blocks → plain `SlideSpec` objects, today's geometry) → `src/ooxml/` (SlideSpec → .pptx bytes; knows nothing about Slidev). A `Package` class owns content types and relationships so they cannot drift.

**Tech Stack:** Node ≥ 18 ESM, `jszip`, `js-yaml`, `image-size@2` (runtime); `@xmldom/xmldom@0.9` (dev); `node --test`; LibreOffice + poppler (`pdfinfo`, `pdftotext`, `pdftoppm`) for the optional render check.

**Spec:** `docs/superpowers/specs/2026-10-13-ooxml-writer-design.md`

## Global Constraints

- No `pptxgenjs` anywhere after Task 13; `src/fix-pptx.mjs` and `src/render.mjs` are deleted.
- `convert(options)` signature/return value, CLI, `pptx.config.mjs` shape and resolver kinds (`statcards`, `barchart-panels`, `barchart`) unchanged. New optional config keys: `lang` (default `"en-US"`), `timestamps` (default `false`).
- Same input ⇒ byte-identical `.pptx` (fixed zip dates, deterministic ids/GUIDs, no dates in `core.xml` unless `timestamps: true`).
- Exactly one `<a:pPr>` per `<a:p>`, always the first child. Every element sequence in schema order.
- Every relationship target exists; every content-type Override names an existing part.
- Slide ids from 256; slide master + layout ids share one counter from 2147483648; shape ids from 2, unique per part.
- Theme slots: `dk1`=text, `lt1`=FFFFFF, `dk2`=accentDark, `lt2`=panel, `accent1`=accent, `accent2`=accent2, `accent3`=muted, `accent4`=border, `accent5`=9CA3AF, `accent6`=F59E0B, `hlink`=accent, `folHlink`=accentDark.
- In slide/chart XML, theme slots `dk1/lt1/dk2/lt2` are written as `schemeClr` `tx1/bg1/tx2/bg2` (what PowerPoint itself writes); accents as-is.
- Code style: ESM `.mjs`, 2-space indent, double quotes, explanatory `//` comments on non-obvious decisions (match `src/parse.mjs`).
- Commit messages end with `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.

## File Structure

```
src/
  index.mjs              MODIFY  convert(): parse → layout → writePptx
  parse.mjs              MODIFY  notes extraction; mdInlineToRuns → mdInlineTokens
  config.mjs             (unchanged)
  default-config.mjs     MODIFY  lang, timestamps
  colors.mjs             CREATE  theme slot mapping, toColor, blendWithWhite
  geometry.mjs           CREATE  slide geometry shared by layout and master
  layout.mjs             CREATE  blocks → SlideSpec (port of render.mjs)
  render.mjs             DELETE
  fix-pptx.mjs           DELETE
  ooxml/
    xml.mjs              CREATE  esc, emu, guid, colorXml, namespaces
    package.mjs          CREATE  Package (parts, content types, rels, zip)
    theme.mjs            CREATE  themeXml
    text.mjs             CREATE  txBodyXml (bodyPr/pPr/rPr)
    shapes.mjs           CREATE  spTree header, placeholder/text/picture/graphicFrame XML
    master.mjs           CREATE  layout definitions, slideMaster + slideLayout XML
    table.mjs            CREATE  deck table styles, a:tbl XML
    xlsx.mjs             CREATE  embedded chart workbook
    chart.mjs            CREATE  c:chartSpace XML
    notes.mjs            CREATE  notesMaster + notesSlide XML
    presentation.mjs     CREATE  presentation/presProps/viewProps/app/core XML
    index.mjs            CREATE  writePptx(specs, deck) → Buffer
test/
  helpers.mjs            CREATE  unzip + XML helpers for tests
  lint.mjs               CREATE  lintPackage(buffer) → string[]
  *.test.mjs             CREATE  one per module
  fixtures/kitchen-sink/ CREATE  slides.md, pptx.config.mjs, public/pic.png, public/data/chart.json
scripts/render.mjs       CREATE  LibreOffice render check
```

Note on module naming: the spec's `slide.mjs` row is implemented as `shapes.mjs` (shape XML builders) plus slide assembly inside `ooxml/index.mjs`, to keep each file focused.

---

### Task 1: Test harness and XML primitives

**Files:**
- Modify: `package.json`
- Create: `src/ooxml/xml.mjs`
- Test: `test/xml.test.mjs`

**Interfaces:**
- Produces:
  - `esc(value: any): string` — strips XML-1.0-invalid chars (incl. lone surrogates, U+FFFE/FFFF), escapes `& < > " '`.
  - `emu(inches: number): number` — rounded EMU, `0` for non-finite.
  - `emuSize(inches)` — like `emu` but never negative.
  - `pt100(pt: number): number` — `Math.round(pt * 100)`.
  - `guid(seed: string): string` — deterministic `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}` uppercase.
  - `colorXml(color: Color, alpha?: number): string`, `solidFill(color, alpha?)` — `Color` is a slot name (`dk1 lt1 dk2 lt2 accent1..accent6`) or `#RRGGBB`.
  - `XML_DECL`, `NS` (`a`, `r`, `p`, `c`, `rels`, `ct`, `ssml`), `NS_PML` (`xmlns:a/r/p` attribute string).

- [ ] **Step 1: Add dev dependency and test script**

Run: `npm install --save-dev @xmldom/xmldom@0.9`

Then edit `package.json` `"scripts"` (add the key; it doesn't exist yet) so it contains:

```json
  "scripts": {
    "test": "node --test test/*.test.mjs",
    "render": "node scripts/render.mjs"
  },
```

- [ ] **Step 2: Write the failing test** — `test/xml.test.mjs`

```js
import { test } from "node:test";
import assert from "node:assert/strict";
import { esc, emu, emuSize, pt100, guid, colorXml, solidFill } from "../src/ooxml/xml.mjs";

test("esc escapes markup characters", () => {
  assert.equal(esc(`a<b>&"c"'`), "a&lt;b&gt;&amp;&quot;c&quot;&apos;");
});

test("esc strips characters XML 1.0 forbids", () => {
  assert.equal(esc("a\u0001b\u0008c\u000Bd"), "abcd");
  assert.equal(esc("x￾y￿z"), "xyz");
  assert.equal(esc("lone\uD800surrogate"), "lonesurrogate");
  assert.equal(esc("keep\t\n\r"), "keep\t\n\r");
  assert.equal(esc("emoji 😀"), "emoji 😀");
});

test("esc handles null/undefined/numbers", () => {
  assert.equal(esc(undefined), "");
  assert.equal(esc(null), "");
  assert.equal(esc(42), "42");
});

test("emu converts inches and guards bad numbers", () => {
  assert.equal(emu(1), 914400);
  assert.equal(emu(0.5), 457200);
  assert.equal(emu(NaN), 0);
  assert.equal(emuSize(-1), 0);
  assert.equal(pt100(13), 1300);
});

test("guid is deterministic and well-formed", () => {
  const a = guid("table-default");
  assert.equal(a, guid("table-default"));
  assert.notEqual(a, guid("table-deflist"));
  assert.match(a, /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/);
});

test("colorXml writes PowerPoint's mapped scheme names", () => {
  assert.equal(colorXml("dk1"), '<a:schemeClr val="tx1"/>');
  assert.equal(colorXml("lt2"), '<a:schemeClr val="bg2"/>');
  assert.equal(colorXml("accent3"), '<a:schemeClr val="accent3"/>');
  assert.equal(colorXml("#1a833b"), '<a:srgbClr val="1A833B"/>');
  assert.equal(colorXml("dk2", 0.16), '<a:schemeClr val="tx2"><a:alpha val="16000"/></a:schemeClr>');
  assert.equal(solidFill("accent1"), '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>');
  assert.throws(() => colorXml("purple"), /Unknown color/);
});
```

- [ ] **Step 3: Run test to verify it fails**

Run: `npm test`
Expected: FAIL — `Cannot find module '.../src/ooxml/xml.mjs'`

- [ ] **Step 4: Implement** — `src/ooxml/xml.mjs`

```js
// Shared XML primitives for every OOXML part writer: escaping, unit conversion, deterministic
// GUIDs and DrawingML color elements. Nothing here knows about slides or Slidev.

export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export const NS = {
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
  c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
  rels: "http://schemas.openxmlformats.org/package/2006/relationships",
  ct: "http://schemas.openxmlformats.org/package/2006/content-types",
  ssml: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
};

export const NS_PML = `xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"`;

// XML 1.0 allows TAB, LF, CR, U+0020-U+D7FF, U+E000-U+FFFD and U+10000-U+10FFFF. With the `u`
// flag a lone surrogate is its own code point, so it falls outside every allowed range and is
// stripped too - JS strings can carry those (Rust strings can't), and one is enough to make
// PowerPoint reject the whole part.
const INVALID_XML_CHARS = /[^\t\n\r -퟿-�\u{10000}-\u{10FFFF}]/gu;

export function esc(value) {
  return String(value ?? "")
    .replace(INVALID_XML_CHARS, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

export const EMU_PER_INCH = 914400;

export function emu(inches) {
  const n = Math.round(Number(inches) * EMU_PER_INCH);
  return Number.isFinite(n) ? n : 0;
}

// Extents (a:ext cx/cy) must not be negative - PowerPoint repairs the file if they are.
export function emuSize(inches) {
  return Math.max(0, emu(inches));
}

export function pt100(pt) {
  return Math.round(Number(pt) * 100);
}

function fnv1a(str) {
  let h = 0x811c9dc5;
  for (const ch of str) {
    h ^= ch.codePointAt(0);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h >>> 0;
}

// Content-derived GUID (four salted 32-bit FNV-1a hashes) so table style ids etc. are stable
// across exports - same input, byte-identical file.
export function guid(seed) {
  const hex = [0, 1, 2, 3]
    .map((i) => fnv1a(`${i}:${seed}`).toString(16).padStart(8, "0"))
    .join("")
    .toUpperCase();
  return `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20, 32)}}`;
}

// Theme slots dk1/lt1/dk2/lt2 are referenced from slide content through the master's color
// map (tx1/bg1/tx2/bg2) - that's what PowerPoint itself writes, and what Google Slides and
// Keynote expect. Accents are referenced directly.
const SCHEME_REF = { dk1: "tx1", lt1: "bg1", dk2: "tx2", lt2: "bg2" };
const ACCENTS = new Set(["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]);

export function colorXml(color, alpha) {
  const mod = alpha != null && alpha < 1 ? `<a:alpha val="${Math.round(alpha * 100000)}"/>` : "";
  let open;
  if (typeof color === "string" && /^#[0-9a-fA-F]{6}$/.test(color)) {
    open = `<a:srgbClr val="${color.slice(1).toUpperCase()}"`;
    return mod ? `${open}>${mod}</a:srgbClr>` : `${open}/>`;
  }
  const val = SCHEME_REF[color] || (ACCENTS.has(color) ? color : null);
  if (!val) throw new Error(`Unknown color ${JSON.stringify(color)} (expected a theme slot or #RRGGBB)`);
  open = `<a:schemeClr val="${val}"`;
  return mod ? `${open}>${mod}</a:schemeClr>` : `${open}/>`;
}

export function solidFill(color, alpha) {
  return `<a:solidFill>${colorXml(color, alpha)}</a:solidFill>`;
}
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `npm test`
Expected: PASS (6 tests)

- [ ] **Step 6: Commit**

```bash
git add package.json package-lock.json src/ooxml/xml.mjs test/xml.test.mjs
git commit -m "Add XML primitives for the OOXML writer and a node --test harness

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 2: Package writer (parts, content types, relationships)

**Files:**
- Create: `src/ooxml/package.mjs`
- Create: `test/helpers.mjs`
- Test: `test/package.test.mjs`

**Interfaces:**
- Consumes: `XML_DECL`, `NS`, `esc` from Task 1.
- Produces:
  - `REL` (relationship type URIs: `officeDocument coreProps extProps slideMaster slideLayout slide theme notesMaster notesSlide presProps viewProps tableStyles image chart package hyperlink worksheet styles sharedStrings`), `CT` (content types: `presentation slideMaster slideLayout slide notesMaster notesSlide presProps viewProps tableStyles theme chart coreProps extProps workbook worksheet styles sharedStrings`).
  - `class Package { addXml(partName, contentType, xml); addMedia(partName, buffer); has(partName): boolean; rel(fromPart, target, type, { external }?): string /* rId */; async toBuffer(): Promise<Buffer> }`. Part names are absolute (`/ppt/slides/slide1.xml`); `fromPart = ""` is the package root. `toBuffer` throws if an internal relationship targets a missing part.
  - `test/helpers.mjs`: `async unzip(buffer) → { names: string[], text(name): string, xml(name): Document, buffer(name): Buffer }`, `elements(node)` generator (all descendant elements), `byLocal(node, localName): Element[]`.

- [ ] **Step 1: Write test helpers** — `test/helpers.mjs`

```js
// Small helpers for asserting on generated packages: unzip into memory and parse parts.
import JSZip from "jszip";
import { DOMParser } from "@xmldom/xmldom";

export function parseXml(text, name = "part") {
  const errors = [];
  let doc;
  try {
    doc = new DOMParser({
      onError: (level, msg) => {
        if (level !== "warning") errors.push(msg);
      },
    }).parseFromString(text, "text/xml");
  } catch (err) {
    errors.push(err.message);
  }
  if (errors.length) throw new Error(`${name} is not well-formed XML: ${errors[0]}`);
  return doc;
}

export async function unzip(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
  const texts = new Map();
  const buffers = new Map();
  for (const n of names) {
    buffers.set(n, await zip.file(n).async("nodebuffer"));
    if (/\.(xml|rels)$/.test(n)) texts.set(n, buffers.get(n).toString("utf8"));
  }
  return {
    names,
    text: (n) => {
      if (!texts.has(n)) throw new Error(`No XML part ${n}`);
      return texts.get(n);
    },
    xml: (n) => parseXml(texts.get(n) ?? (() => { throw new Error(`No XML part ${n}`); })(), n),
    buffer: (n) => buffers.get(n),
  };
}

export function* elements(node) {
  for (let c = node.firstChild; c; c = c.nextSibling) {
    if (c.nodeType === 1) {
      yield c;
      yield* elements(c);
    }
  }
}

export function byLocal(node, localName) {
  return [...elements(node)].filter((e) => e.localName === localName);
}
```

- [ ] **Step 2: Write the failing test** — `test/package.test.mjs`

```js
import { test } from "node:test";
import assert from "node:assert/strict";
import { Package, REL, CT } from "../src/ooxml/package.mjs";
import { unzip } from "./helpers.mjs";

function samplePackage() {
  const pkg = new Package();
  pkg.addXml("/ppt/presentation.xml", CT.presentation, "<x/>");
  pkg.addXml("/ppt/slides/slide1.xml", CT.slide, "<s/>");
  pkg.addMedia("/ppt/media/image1.png", Buffer.from([1, 2, 3]));
  pkg.rel("", "/ppt/presentation.xml", REL.officeDocument);
  pkg.rel("/ppt/presentation.xml", "/ppt/slides/slide1.xml", REL.slide);
  pkg.rel("/ppt/slides/slide1.xml", "/ppt/media/image1.png", REL.image);
  return pkg;
}

test("content types list exactly the parts that were added", async () => {
  const z = await unzip(await samplePackage().toBuffer());
  const ct = z.text("[Content_Types].xml");
  assert.match(ct, /<Override PartName="\/ppt\/presentation.xml" ContentType="application\/vnd.openxmlformats-officedocument.presentationml.presentation.main\+xml"\/>/);
  assert.match(ct, /<Override PartName="\/ppt\/slides\/slide1.xml"/);
  assert.match(ct, /<Default Extension="png" ContentType="image\/png"\/>/);
  assert.equal((ct.match(/<Override /g) || []).length, 2);
  assert.equal(z.names[0], "[Content_Types].xml");
});

test("relationship targets are relative to the source part", async () => {
  const z = await unzip(await samplePackage().toBuffer());
  assert.match(z.text("_rels/.rels"), /Target="ppt\/presentation.xml"/);
  assert.match(z.text("ppt/_rels/presentation.xml.rels"), /Id="rId1" Type="[^"]+\/slide" Target="slides\/slide1.xml"/);
  assert.match(z.text("ppt/slides/_rels/slide1.xml.rels"), /Target="..\/media\/image1.png"/);
});

test("rel() reuses identical relationships and numbers new ones per part", () => {
  const pkg = new Package();
  assert.equal(pkg.rel("/a.xml", "/b.xml", REL.slide), "rId1");
  assert.equal(pkg.rel("/a.xml", "/b.xml", REL.slide), "rId1");
  assert.equal(pkg.rel("/a.xml", "https://example.com", REL.hyperlink, { external: true }), "rId2");
  assert.equal(pkg.rel("/c.xml", "/b.xml", REL.slide), "rId1");
});

test("external relationships carry TargetMode and are not resolved", async () => {
  const pkg = samplePackage();
  pkg.rel("/ppt/slides/slide1.xml", "https://example.com/?a=1&b=2", REL.hyperlink, { external: true });
  const z = await unzip(await pkg.toBuffer());
  assert.match(z.text("ppt/slides/_rels/slide1.xml.rels"), /Target="https:\/\/example.com\/\?a=1&amp;b=2" TargetMode="External"/);
});

test("toBuffer refuses a relationship to a missing part", async () => {
  const pkg = samplePackage();
  pkg.rel("/ppt/presentation.xml", "/ppt/slides/slide2.xml", REL.slide);
  await assert.rejects(pkg.toBuffer(), /missing part \/ppt\/slides\/slide2.xml/);
});

test("duplicate parts and unknown media types are rejected", () => {
  const pkg = samplePackage();
  assert.throws(() => pkg.addXml("/ppt/slides/slide1.xml", CT.slide, "<s/>"), /Duplicate part/);
  assert.throws(() => pkg.addMedia("/ppt/media/x.bmp", Buffer.from([])), /Unsupported media/);
});

test("output is byte-identical across runs", async () => {
  const a = await samplePackage().toBuffer();
  const b = await samplePackage().toBuffer();
  assert.ok(a.equals(b));
});
```

- [ ] **Step 3: Run test to verify it fails**

Run: `npm test`
Expected: FAIL — `Cannot find module '.../src/ooxml/package.mjs'`

- [ ] **Step 4: Implement** — `src/ooxml/package.mjs`

```js
// OPC package writer. The only way to add a part is through addXml/addMedia, which also
// registers its content type, and [Content_Types].xml and every _rels/*.rels are generated
// from what was actually added - so a phantom Override (pptxgenjs#1449) or a relationship to a
// part that doesn't exist can't be written. Same calls in the same order give the same bytes.

import JSZip from "jszip";
import { XML_DECL, NS, esc } from "./xml.mjs";

const OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export const REL = {
  officeDocument: `${OFFICE_REL}/officeDocument`,
  coreProps: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
  extProps: `${OFFICE_REL}/extended-properties`,
  slideMaster: `${OFFICE_REL}/slideMaster`,
  slideLayout: `${OFFICE_REL}/slideLayout`,
  slide: `${OFFICE_REL}/slide`,
  theme: `${OFFICE_REL}/theme`,
  notesMaster: `${OFFICE_REL}/notesMaster`,
  notesSlide: `${OFFICE_REL}/notesSlide`,
  presProps: `${OFFICE_REL}/presProps`,
  viewProps: `${OFFICE_REL}/viewProps`,
  tableStyles: `${OFFICE_REL}/tableStyles`,
  image: `${OFFICE_REL}/image`,
  chart: `${OFFICE_REL}/chart`,
  package: `${OFFICE_REL}/package`,
  hyperlink: `${OFFICE_REL}/hyperlink`,
  worksheet: `${OFFICE_REL}/worksheet`,
  styles: `${OFFICE_REL}/styles`,
  sharedStrings: `${OFFICE_REL}/sharedStrings`,
};

const PML = "application/vnd.openxmlformats-officedocument.presentationml";
const SML = "application/vnd.openxmlformats-officedocument.spreadsheetml";

export const CT = {
  presentation: `${PML}.presentation.main+xml`,
  slideMaster: `${PML}.slideMaster+xml`,
  slideLayout: `${PML}.slideLayout+xml`,
  slide: `${PML}.slide+xml`,
  notesMaster: `${PML}.notesMaster+xml`,
  notesSlide: `${PML}.notesSlide+xml`,
  presProps: `${PML}.presProps+xml`,
  viewProps: `${PML}.viewProps+xml`,
  tableStyles: `${PML}.tableStyles+xml`,
  theme: "application/vnd.openxmlformats-officedocument.theme+xml",
  chart: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
  coreProps: "application/vnd.openxmlformats-package.core-properties+xml",
  extProps: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
  workbook: `${SML}.sheet.main+xml`,
  worksheet: `${SML}.worksheet+xml`,
  styles: `${SML}.styles+xml`,
  sharedStrings: `${SML}.sharedStrings+xml`,
};

const MEDIA_TYPES = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  xlsx: `${SML}.sheet`,
};

// DOS epoch: the earliest date a zip entry can carry. Fixed so output is reproducible.
const FIXED_DATE = new Date(Date.UTC(1980, 0, 1, 0, 0, 0));

function relsPartFor(sourcePart) {
  if (sourcePart === "") return "/_rels/.rels";
  const i = sourcePart.lastIndexOf("/");
  return `${sourcePart.slice(0, i)}/_rels/${sourcePart.slice(i + 1)}.rels`;
}

function relativeTarget(fromPart, toPart) {
  const fromDir = fromPart.split("/").slice(1, -1);
  const to = toPart.split("/").slice(1);
  let i = 0;
  while (i < fromDir.length && i < to.length - 1 && fromDir[i] === to[i]) i++;
  return "../".repeat(fromDir.length - i) + to.slice(i).join("/");
}

function assertPartName(name) {
  if (!/^\/[^\s]+[^/]$/.test(name)) throw new Error(`Invalid part name ${JSON.stringify(name)}`);
}

export class Package {
  constructor() {
    this.parts = new Map(); // partName -> { data, compress }
    this.overrides = new Map(); // partName -> content type
    this.defaults = new Map([
      ["rels", "application/vnd.openxmlformats-package.relationships+xml"],
      ["xml", "application/xml"],
    ]);
    this.rels = new Map(); // sourcePart ("" = package root) -> [{ id, type, target, external }]
  }

  addXml(partName, contentType, xml) {
    assertPartName(partName);
    if (this.parts.has(partName)) throw new Error(`Duplicate part ${partName}`);
    this.parts.set(partName, { data: xml, compress: true });
    this.overrides.set(partName, contentType);
  }

  addMedia(partName, buffer) {
    assertPartName(partName);
    if (this.parts.has(partName)) throw new Error(`Duplicate part ${partName}`);
    const ext = partName.split(".").pop().toLowerCase();
    const type = MEDIA_TYPES[ext];
    if (!type) throw new Error(`Unsupported media type for ${partName}`);
    // Already-compressed binaries are stored, not deflated again.
    this.parts.set(partName, { data: buffer, compress: false });
    this.defaults.set(ext, type);
  }

  has(partName) {
    return this.parts.has(partName);
  }

  // Returns the rId of a relationship from `fromPart` to `target`, reusing an identical one.
  rel(fromPart, target, type, { external = false } = {}) {
    if (!this.rels.has(fromPart)) this.rels.set(fromPart, []);
    const list = this.rels.get(fromPart);
    const found = list.find((r) => r.type === type && r.target === target && r.external === external);
    if (found) return found.id;
    const id = `rId${list.length + 1}`;
    list.push({ id, type, target, external });
    return id;
  }

  contentTypesXml() {
    const defaults = [...this.defaults]
      .map(([ext, type]) => `<Default Extension="${esc(ext)}" ContentType="${esc(type)}"/>`)
      .join("");
    const overrides = [...this.overrides]
      .map(([name, type]) => `<Override PartName="${esc(name)}" ContentType="${esc(type)}"/>`)
      .join("");
    return `${XML_DECL}<Types xmlns="${NS.ct}">${defaults}${overrides}</Types>`;
  }

  relsXml(sourcePart, list) {
    const items = list
      .map((r) => {
        const target = r.external ? r.target : relativeTarget(sourcePart, r.target);
        const mode = r.external ? ' TargetMode="External"' : "";
        return `<Relationship Id="${r.id}" Type="${esc(r.type)}" Target="${esc(target)}"${mode}/>`;
      })
      .join("");
    return `${XML_DECL}<Relationships xmlns="${NS.rels}">${items}</Relationships>`;
  }

  async toBuffer() {
    for (const [source, list] of this.rels) {
      for (const r of list) {
        if (!r.external && !this.parts.has(r.target)) {
          throw new Error(`Relationship ${r.id} from ${source || "/"} points at missing part ${r.target}`);
        }
      }
    }
    const zip = new JSZip();
    const add = (name, data, compress) =>
      zip.file(name.slice(1), data, { date: FIXED_DATE, createFolders: false, compression: compress ? "DEFLATE" : "STORE" });

    add("/[Content_Types].xml", this.contentTypesXml(), true);
    if (this.rels.has("")) add(relsPartFor(""), this.relsXml("", this.rels.get("")), true);
    for (const [name, part] of this.parts) {
      add(name, part.data, part.compress);
      const list = this.rels.get(name);
      if (list && list.length) add(relsPartFor(name), this.relsXml(name, list), true);
    }
    return zip.generateAsync({ type: "nodebuffer", compressionOptions: { level: 6 } });
  }
}
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `npm test`
Expected: PASS (all tests in `xml.test.mjs` and `package.test.mjs`)

- [ ] **Step 6: Commit**

```bash
git add src/ooxml/package.mjs test/helpers.mjs test/package.test.mjs
git commit -m "Add OPC package writer that derives content types and rels from added parts

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 3: Package lint

**Files:**
- Create: `test/lint.mjs`
- Test: `test/lint.test.mjs`

**Interfaces:**
- Consumes: `Package`, `REL`, `CT` (Task 2); `parseXml`, `elements` (Task 2 helpers).
- Produces: `async lintPackage(buffer: Buffer, label = ""): Promise<string[]>` — empty array means clean. Recurses into embedded `.xlsx` parts. Checks: well-formed XML; Overrides name existing parts; every part has a content type; rel targets exist; rel ids unique per part; every `r:*` attribute used in a part resolves in its rels; `cNvPr` ids unique per part; `sldId` ≥ 256 and unique; master/layout ids ≥ 2147483648 and unique; ≤ 1 `a:pPr` per `a:p` and it's first; every `txBody`/`txPr`/`rich` has ≥ 1 `a:p`; every `c:barChart` has exactly 2 `c:axId`.

- [ ] **Step 1: Write the failing test** — `test/lint.test.mjs`

```js
import { test } from "node:test";
import assert from "node:assert/strict";
import { Package, REL, CT } from "../src/ooxml/package.mjs";
import { lintPackage } from "./lint.mjs";

const P = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

function slide(body) {
  return `<p:sld ${P}><p:cSld><p:spTree>${body}</p:spTree></p:cSld></p:sld>`;
}

async function lintSlide(body, extra = (pkg) => {}) {
  const pkg = new Package();
  pkg.addXml("/ppt/slides/slide1.xml", CT.slide, slide(body));
  extra(pkg);
  return lintPackage(await pkg.toBuffer());
}

test("a clean package has no problems", async () => {
  assert.deepEqual(await lintSlide('<p:sp><p:nvSpPr><p:cNvPr id="2" name="a"/></p:nvSpPr><p:txBody><a:p/></p:txBody></p:sp>'), []);
});

test("flags two pPr in one paragraph (the pptxgenjs bullet bug)", async () => {
  const problems = await lintSlide('<p:sp><p:txBody><a:p><a:pPr/><a:r><a:t>x</a:t></a:r><a:pPr/></a:p></p:txBody></p:sp>');
  assert.ok(problems.some((p) => /more than one a:pPr/.test(p)), problems.join("\n"));
});

test("flags pPr that isn't the first child", async () => {
  const problems = await lintSlide('<p:sp><p:txBody><a:p><a:r><a:t>x</a:t></a:r><a:pPr/></a:p></p:txBody></p:sp>');
  assert.ok(problems.some((p) => /a:pPr is not the first child/.test(p)), problems.join("\n"));
});

test("flags duplicate shape ids and empty text bodies", async () => {
  const problems = await lintSlide('<p:cNvPr id="2" name="a"/><p:cNvPr id="2" name="b"/><p:txBody></p:txBody>');
  assert.ok(problems.some((p) => /duplicate cNvPr id 2/.test(p)), problems.join("\n"));
  assert.ok(problems.some((p) => /txBody without a:p/.test(p)), problems.join("\n"));
});

test("flags an r:id that is not in the part's rels", async () => {
  const problems = await lintSlide('<p:pic><a:blip r:embed="rId9"/></p:pic>');
  assert.ok(problems.some((p) => /r:embed="rId9" not defined/.test(p)), problems.join("\n"));
});

test("flags a bar chart with three axis ids", async () => {
  const problems = await lintSlide("", (pkg) =>
    pkg.addXml("/ppt/charts/chart1.xml", CT.chart, '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:axId val="1"/><c:axId val="2"/><c:axId val="3"/></c:barChart></c:plotArea></c:chart></c:chartSpace>')
  );
  assert.ok(problems.some((p) => /barChart has 3 c:axId/.test(p)), problems.join("\n"));
});

test("flags malformed XML", async () => {
  const problems = await lintSlide("<p:sp>");
  assert.ok(problems.some((p) => /not well-formed/.test(p)), problems.join("\n"));
});

test("flags an Override for a part that is missing (pptxgenjs#1449)", async () => {
  const JSZip = (await import("jszip")).default;
  const zip = new JSZip();
  zip.file("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/slideMasters/slideMaster2.xml" ContentType="x"/></Types>');
  const problems = await lintPackage(await zip.generateAsync({ type: "nodebuffer" }));
  assert.ok(problems.some((p) => /Override for missing part \/ppt\/slideMasters\/slideMaster2.xml/.test(p)), problems.join("\n"));
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test`
Expected: FAIL — `Cannot find module '.../test/lint.mjs'`

- [ ] **Step 3: Implement** — `test/lint.mjs`

```js
// Package lint: structural checks for the things that make PowerPoint (especially Online)
// show a repair prompt. Not a schema validator - it checks the invariants we've been bitten
// by or that DeckCraft's writer guarantees by construction.

import JSZip from "jszip";
import { parseXml, elements } from "./helpers.mjs";

const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

function sourceOfRels(relsName) {
  // "ppt/slides/_rels/slide1.xml.rels" -> "ppt/slides/slide1.xml"; "_rels/.rels" -> ""
  const m = /^(.*?)_rels\/(.*)\.rels$/.exec(relsName);
  return m ? `${m[1]}${m[2]}` : null;
}

function resolveTarget(sourcePart, target) {
  if (target.startsWith("/")) return target.slice(1);
  const parts = sourcePart.split("/").slice(0, -1);
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== ".") parts.push(seg);
  }
  return parts.join("/");
}

export async function lintPackage(buffer, label = "") {
  const problems = [];
  const report = (msg) => problems.push(label ? `${label}: ${msg}` : msg);
  const zip = await JSZip.loadAsync(buffer);
  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
  const nameSet = new Set(names);

  const docs = new Map();
  for (const n of names.filter((n) => /\.(xml|rels)$/.test(n))) {
    try {
      docs.set(n, parseXml(await zip.file(n).async("string"), n));
    } catch (err) {
      report(err.message);
    }
  }

  // Content types.
  const ct = docs.get("[Content_Types].xml");
  if (!ct) {
    report("missing or unparseable [Content_Types].xml");
  } else {
    const defaults = new Set();
    const overrides = new Set();
    for (const el of elements(ct)) {
      if (el.localName === "Default") defaults.add(el.getAttribute("Extension").toLowerCase());
      if (el.localName === "Override") {
        const part = el.getAttribute("PartName").slice(1);
        overrides.add(part);
        if (!nameSet.has(part)) report(`Override for missing part /${part}`);
      }
    }
    for (const n of names) {
      if (n === "[Content_Types].xml") continue;
      const ext = n.split(".").pop().toLowerCase();
      if (!overrides.has(n) && !defaults.has(ext)) report(`no content type for /${n}`);
    }
  }

  // Relationships.
  const relIds = new Map(); // source part -> Set(ids)
  for (const [n, doc] of docs) {
    if (!n.endsWith(".rels")) continue;
    const source = sourceOfRels(n);
    const ids = new Set();
    for (const el of elements(doc)) {
      if (el.localName !== "Relationship") continue;
      const id = el.getAttribute("Id");
      if (ids.has(id)) report(`${n}: duplicate relationship id ${id}`);
      ids.add(id);
      if (el.getAttribute("TargetMode") === "External") continue;
      const target = resolveTarget(source, el.getAttribute("Target"));
      if (!nameSet.has(target)) report(`${n}: ${id} targets missing part /${target}`);
    }
    relIds.set(source, ids);
  }

  // Per-part checks.
  const masterLayoutIds = new Set();
  for (const [n, doc] of docs) {
    if (n.endsWith(".rels") || n === "[Content_Types].xml") continue;
    const cNvPrIds = new Set();
    for (const el of elements(doc)) {
      for (let i = 0; i < el.attributes.length; i++) {
        const attr = el.attributes[i];
        if (attr.namespaceURI === R_NS && attr.value !== "") {
          if (!(relIds.get(n) || new Set()).has(attr.value)) report(`${n}: ${attr.name}="${attr.value}" not defined in its rels`);
        }
      }
      if (el.localName === "cNvPr") {
        const id = el.getAttribute("id");
        if (cNvPrIds.has(id)) report(`${n}: duplicate cNvPr id ${id}`);
        cNvPrIds.add(id);
      }
      if (el.localName === "p" && el.namespaceURI === A_NS) {
        const kids = [...elements(el)].filter((k) => k.parentNode === el);
        const pPrs = kids.filter((k) => k.localName === "pPr");
        if (pPrs.length > 1) report(`${n}: a:p with more than one a:pPr`);
        if (pPrs.length === 1 && kids[0] !== pPrs[0]) report(`${n}: a:pPr is not the first child of a:p`);
      }
      if (["txBody", "txPr", "rich"].includes(el.localName)) {
        const hasP = [...elements(el)].some((k) => k.parentNode === el && k.localName === "p");
        if (!hasP) report(`${n}: ${el.localName === "txBody" ? "txBody" : el.localName} without a:p`);
      }
      if (el.localName === "barChart") {
        const axIds = [...elements(el)].filter((k) => k.parentNode === el && k.localName === "axId").length;
        if (axIds !== 2) report(`${n}: barChart has ${axIds} c:axId (expected 2)`);
      }
      if (el.localName === "sldId") {
        const id = Number(el.getAttribute("id"));
        if (!(id >= 256)) report(`${n}: sldId ${id} < 256`);
      }
      if (el.localName === "sldMasterId" || el.localName === "sldLayoutId") {
        const id = Number(el.getAttribute("id"));
        if (!(id >= 2147483648)) report(`${n}: ${el.localName} ${id} < 2147483648`);
        if (masterLayoutIds.has(id)) report(`${n}: master/layout id ${id} used twice`);
        masterLayoutIds.add(id);
      }
    }
  }

  // Embedded workbooks are packages too.
  for (const n of names.filter((n) => n.endsWith(".xlsx"))) {
    problems.push(...(await lintPackage(await zip.file(n).async("nodebuffer"), `${label ? `${label} > ` : ""}${n}`)));
  }

  return problems;
}
```

Note: the spec's "sldId unique" is enforced because `sldId` duplicates also produce duplicate `r:id` rels in practice; ids come from a counter in Task 8, so a separate uniqueness check isn't needed.

- [ ] **Step 4: Run tests to verify they pass**

Run: `npm test`
Expected: PASS (all lint tests plus earlier tasks)

- [ ] **Step 5: Commit**

```bash
git add test/lint.mjs test/lint.test.mjs
git commit -m "Add package lint for repair-prompt causes (pPr order, rels, ids, axIds)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 4: Theme colors and theme part

**Files:**
- Create: `src/colors.mjs`
- Create: `src/ooxml/theme.mjs`
- Test: `test/theme.test.mjs`

**Interfaces:**
- Consumes: `esc`, `XML_DECL`, `NS` (Task 1).
- Produces:
  - `src/colors.mjs`: `normalizeHex(hex): "RRGGBB"` (accepts `#`, 3-digit; throws on garbage), `themeColors(configColors): ThemeSlots` (object with the 12 slots, hex without `#`), `toColor(hex, theme): Color` (slot name if any of `dk1 lt1 dk2 lt2 accent1..6` matches, else `#RRGGBB`), `hexOf(color, theme): "RRGGBB"`, `blendWithWhite(hex, alpha): "RRGGBB"`.
  - `src/ooxml/theme.mjs`: `themeXml({ name, colors: ThemeSlots, fonts: { heading, body } }): string`.

- [ ] **Step 1: Write the failing test** — `test/theme.test.mjs`

```js
import { test } from "node:test";
import assert from "node:assert/strict";
import { normalizeHex, themeColors, toColor, hexOf, blendWithWhite } from "../src/colors.mjs";
import { themeXml } from "../src/ooxml/theme.mjs";
import { parseXml, byLocal } from "./helpers.mjs";

const COLORS = { accent: "1A833B", accentDark: "14532D", accent2: "1A589F", panel: "EEF4F0", border: "D9E3DB", text: "1F2937", muted: "6B7280" };

test("normalizeHex accepts #, lowercase and 3-digit forms", () => {
  assert.equal(normalizeHex("#1a833b"), "1A833B");
  assert.equal(normalizeHex("fff"), "FFFFFF");
  assert.throws(() => normalizeHex("green"), /Invalid color/);
});

test("themeColors maps config colors onto the 12 slots", () => {
  assert.deepEqual(themeColors(COLORS), {
    dk1: "1F2937", lt1: "FFFFFF", dk2: "14532D", lt2: "EEF4F0",
    accent1: "1A833B", accent2: "1A589F", accent3: "6B7280", accent4: "D9E3DB",
    accent5: "9CA3AF", accent6: "F59E0B", hlink: "1A833B", folHlink: "14532D",
  });
});

test("toColor resolves theme colors to slots and others to hex", () => {
  const theme = themeColors(COLORS);
  assert.equal(toColor("1F2937", theme), "dk1");
  assert.equal(toColor("#1a833b", theme), "accent1");
  assert.equal(toColor("FFFFFF", theme), "lt1");
  assert.equal(toColor("9CA3AF", theme), "accent5");
  assert.equal(toColor("123456", theme), "#123456");
  assert.equal(hexOf("dk2", theme), "14532D");
  assert.equal(hexOf("#abcdef", theme), "ABCDEF");
});

test("blendWithWhite reproduces CSS opacity over white", () => {
  assert.equal(blendWithWhite("000000", 0.5), "808080");
  assert.equal(blendWithWhite("1F2937", 1), "1F2937");
});

test("themeXml writes slots in schema order, fonts and 3-entry style lists", () => {
  const xml = themeXml({ name: "Deck", colors: themeColors(COLORS), fonts: { heading: "Plus Jakarta Sans", body: "Inter" } });
  const doc = parseXml(xml);
  const scheme = byLocal(doc, "clrScheme")[0];
  const slots = [];
  for (let c = scheme.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) slots.push(c.localName);
  assert.deepEqual(slots, ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]);
  assert.match(xml, /<a:majorFont><a:latin typeface="Plus Jakarta Sans"\/>/);
  assert.match(xml, /<a:minorFont><a:latin typeface="Inter"\/>/);
  for (const list of ["fillStyleLst", "lnStyleLst", "effectStyleLst", "bgFillStyleLst"]) {
    const el = byLocal(doc, list)[0];
    let count = 0;
    for (let c = el.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) count++;
    assert.equal(count, 3, list);
  }
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test`
Expected: FAIL — `Cannot find module '.../src/colors.mjs'`

- [ ] **Step 3: Implement** — `src/colors.mjs`

```js
// Theme color mapping. The config's named colors become the deck's theme (so PowerPoint's
// Design > Variants can recolor everything), and any hex color layout uses is resolved back to
// a theme slot when it matches one - otherwise it stays a fixed hex color.

const SLOT_ORDER = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];

export function normalizeHex(hex) {
  let h = String(hex ?? "").trim().replace(/^#/, "").toUpperCase();
  if (/^[0-9A-F]{3}$/.test(h)) h = h.split("").map((c) => c + c).join("");
  if (!/^[0-9A-F]{6}$/.test(h)) throw new Error(`Invalid color ${JSON.stringify(hex)} (expected RRGGBB)`);
  return h;
}

export function themeColors(colors) {
  const n = normalizeHex;
  return {
    dk1: n(colors.text),
    lt1: "FFFFFF",
    dk2: n(colors.accentDark),
    lt2: n(colors.panel),
    accent1: n(colors.accent),
    accent2: n(colors.accent2),
    accent3: n(colors.muted),
    accent4: n(colors.border),
    // The config has no fifth/sixth accent; neutral gray and amber keep charts with extra
    // series readable.
    accent5: "9CA3AF",
    accent6: "F59E0B",
    hlink: n(colors.accent),
    folHlink: n(colors.accentDark),
  };
}

export function toColor(hex, theme) {
  const h = normalizeHex(hex);
  return SLOT_ORDER.find((slot) => theme[slot] === h) || `#${h}`;
}

export function hexOf(color, theme) {
  return color.startsWith("#") ? normalizeHex(color) : theme[color];
}

// CSS opacity composited against the white slide background - the faded shade a Tailwind
// `opacity-NN` wrapper actually produces on screen.
export function blendWithWhite(hex, alpha) {
  const n = parseInt(normalizeHex(hex), 16);
  const mix = (c) => Math.round(c * alpha + 255 * (1 - alpha));
  return [mix((n >> 16) & 0xff), mix((n >> 8) & 0xff), mix(n & 0xff)]
    .map((v) => v.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();
}
```

- [ ] **Step 4: Implement** — `src/ooxml/theme.mjs`

```js
// theme1.xml from the deck's theme slots and fonts. The format scheme is the minimal valid one
// (each list needs exactly three entries); shapes here set their own fills/lines, so it only
// matters for shapes users insert later in PowerPoint.

import { XML_DECL, NS, esc } from "./xml.mjs";

const SLOTS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

const LINE = (w) =>
  `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;

const FMT_SCHEME =
  '<a:fmtScheme name="Deck">' +
  "<a:fillStyleLst>" +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:solidFill><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:solidFill>' +
  '<a:solidFill><a:schemeClr val="phClr"><a:shade val="80000"/></a:schemeClr></a:solidFill>' +
  "</a:fillStyleLst>" +
  `<a:lnStyleLst>${LINE(6350)}${LINE(12700)}${LINE(19050)}</a:lnStyleLst>` +
  "<a:effectStyleLst>" +
  "<a:effectStyle><a:effectLst/></a:effectStyle>".repeat(3) +
  "</a:effectStyleLst>" +
  "<a:bgFillStyleLst>" +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/></a:schemeClr></a:solidFill>' +
  '<a:solidFill><a:schemeClr val="phClr"><a:shade val="90000"/></a:schemeClr></a:solidFill>' +
  "</a:bgFillStyleLst>" +
  "</a:fmtScheme>";

export function themeXml({ name = "Deck", colors, fonts }) {
  const clr = SLOTS.map((s) => `<a:${s}><a:srgbClr val="${colors[s]}"/></a:${s}>`).join("");
  const font = (typeface) => `<a:latin typeface="${esc(typeface)}"/><a:ea typeface=""/><a:cs typeface=""/>`;
  return (
    `${XML_DECL}<a:theme xmlns:a="${NS.a}" name="${esc(name)}"><a:themeElements>` +
    `<a:clrScheme name="${esc(name)}">${clr}</a:clrScheme>` +
    `<a:fontScheme name="${esc(name)}"><a:majorFont>${font(fonts.heading)}</a:majorFont><a:minorFont>${font(fonts.body)}</a:minorFont></a:fontScheme>` +
    FMT_SCHEME +
    "</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>"
  );
}
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `npm test`
Expected: PASS

- [ ] **Step 6: Commit**

```bash
git add src/colors.mjs src/ooxml/theme.mjs test/theme.test.mjs
git commit -m "Map config colors onto theme slots and write theme1.xml

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 5: Text bodies (bodyPr / pPr / rPr)

**Files:**
- Create: `src/ooxml/text.mjs`
- Test: `test/text.test.mjs`

**Interfaces:**
- Consumes: `esc`, `emu`, `pt100`, `solidFill` (Task 1).
- Produces (types used by every later task):
  - `Run = { text: string, b?: boolean, i?: boolean, font?: "heading"|"body"|"mono", size?: number /*pt*/, color?: Color, link?: string, spacing?: number /*pt*/ }`
  - `Para = { level?: 0..8, bullet?: "char"|"number"|"none", number?: number, align?: "l"|"ctr"|"r", spaceBefore?: number /*pt*/, runs: Run[] }`
  - `TextCtx = { lang: string, fonts: { mono: string }, link?: (url) => rId }`
  - `BodyOpts = { placeholder?: boolean, anchor?: "t"|"ctr"|"b", inset?: number /*in*/, wrap?: boolean, autofit?: "none"|"norm"|"shape", fontScale?: number, lstStyle?: string /*raw XML*/ }`
  - `txBodyXml(paras: Para[], ctx: TextCtx, body?: BodyOpts, tag = "p:txBody"): string`
  - `rPrXml(run, ctx, tag = "a:rPr"): string`, `bodyPrXml(body): string`, `pPrXml(para, ctx, { placeholder }): string`.
- Rules: in placeholders, `bullet:"char"` writes no bullet XML (inherits the master); `"none"` writes `marL="0" indent="0"` + `<a:buNone/>`; `"number"` writes `<a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/>`. In free text boxes, `"char"`/`"number"` also write `marL = 0.25"×(level+1)`, `indent = -0.2"`, and `"char"` writes `buClr accent1`, `buFont Arial`, `buChar •`.

- [ ] **Step 1: Write the failing test** — `test/text.test.mjs`

```js
import { test } from "node:test";
import assert from "node:assert/strict";
import { txBodyXml, bodyPrXml } from "../src/ooxml/text.mjs";
import { parseXml, byLocal } from "./helpers.mjs";

const ctx = { lang: "en-US", fonts: { mono: "JetBrains Mono" }, link: (url) => (url === "https://x.test" ? "rId7" : "rId8") };
const wrap = (xml) => `<p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">${xml}</p:sp>`;

test("a bullet with mixed formatting is ONE paragraph with at most one pPr, first", () => {
  const xml = txBodyXml(
    [{ bullet: "char", runs: [{ text: "Bold", b: true, color: "dk2" }, { text: " for emphasis, " }, { text: "code", font: "mono" }] }],
    ctx,
    {}
  );
  const doc = parseXml(wrap(xml));
  const ps = byLocal(doc, "p");
  assert.equal(ps.length, 1);
  const kids = [];
  for (let c = ps[0].firstChild; c; c = c.nextSibling) if (c.nodeType === 1) kids.push(c.localName);
  assert.equal(kids.filter((k) => k === "pPr").length, 1);
  assert.equal(kids[0], "pPr");
  assert.deepEqual(kids.slice(1), ["r", "r", "r", "endParaRPr"]);
});

test("free text box bullets write indent and bullet children in schema order", () => {
  const xml = txBodyXml([{ bullet: "char", runs: [{ text: "x" }] }], ctx, {});
  assert.match(xml, /<a:pPr marL="228600" indent="-182880"><a:buClr><a:schemeClr val="accent1"\/><\/a:buClr><a:buFont typeface="Arial"\/><a:buChar char="•"\/><\/a:pPr>/);
});

test("placeholder bullets inherit; none and number are explicit", () => {
  const xml = txBodyXml(
    [
      { bullet: "char", runs: [{ text: "a" }] },
      { bullet: "char", level: 1, runs: [{ text: "b" }] },
      { bullet: "none", runs: [{ text: "c" }] },
      { bullet: "number", number: 3, runs: [{ text: "d" }] },
    ],
    ctx,
    { placeholder: true }
  );
  assert.match(xml, /<a:p><a:r><a:rPr lang="en-US" dirty="0"\/><a:t>a<\/a:t>/);
  assert.match(xml, /<a:pPr lvl="1"\/><a:r>/);
  assert.match(xml, /<a:pPr marL="0" indent="0"><a:buNone\/><\/a:pPr>/);
  assert.match(xml, /<a:buFont typeface="\+mj-lt"\/><a:buAutoNum type="arabicPeriod" startAt="3"\/>/);
});

test("spaceBefore and alignment go on pPr", () => {
  const xml = txBodyXml([{ align: "ctr", spaceBefore: 12, runs: [{ text: "x" }] }], ctx, {});
  assert.match(xml, /<a:pPr algn="ctr"><a:spcBef><a:spcPts val="1200"\/><\/a:spcBef><\/a:pPr>/);
});

test("runs: size, bold, italic, spacing, color, theme fonts, mono, link", () => {
  const xml = txBodyXml(
    [{ runs: [
      { text: "T", size: 26, b: true, i: true, spacing: 1.5, color: "accent1", font: "heading" },
      { text: "c", font: "mono" },
      { text: "l", font: "body", link: "https://x.test" },
    ] }],
    ctx,
    {}
  );
  assert.match(xml, /<a:rPr lang="en-US" sz="2600" b="1" i="1" spc="150" dirty="0"><a:solidFill><a:schemeClr val="accent1"\/><\/a:solidFill><a:latin typeface="\+mj-lt"\/><a:ea typeface="\+mj-ea"\/><a:cs typeface="\+mj-cs"\/><\/a:rPr><a:t>T<\/a:t>/);
  assert.match(xml, /<a:latin typeface="JetBrains Mono"\/><a:cs typeface="JetBrains Mono"\/><\/a:rPr><a:t>c<\/a:t>/);
  assert.match(xml, /<a:latin typeface="\+mn-lt"\/><a:ea typeface="\+mn-ea"\/><a:cs typeface="\+mn-cs"\/><a:hlinkClick r:id="rId7"\/><\/a:rPr><a:t>l<\/a:t>/);
});

test("newlines become a:br and text is escaped", () => {
  const xml = txBodyXml([{ runs: [{ text: "a < b\nc & d" }] }], ctx, {});
  assert.match(xml, /<a:t>a &lt; b<\/a:t><\/a:r><a:br><a:rPr lang="en-US" dirty="0"\/><\/a:br><a:r><a:rPr lang="en-US" dirty="0"\/><a:t>c &amp; d<\/a:t>/);
});

test("empty bodies still contain one paragraph", () => {
  assert.match(txBodyXml([], ctx, {}), /<a:lstStyle\/><a:p><a:endParaRPr lang="en-US" dirty="0"\/><\/a:p><\/p:txBody>$/);
});

test("bodyPr: attributes only when set, autofit children", () => {
  assert.equal(bodyPrXml({ placeholder: true }), "<a:bodyPr/>");
  assert.equal(bodyPrXml({ placeholder: true, fontScale: 0.85 }), '<a:bodyPr><a:normAutofit fontScale="85000" lnSpcReduction="10000"/></a:bodyPr>');
  assert.equal(bodyPrXml({ anchor: "ctr", inset: 0.1 }), '<a:bodyPr wrap="square" lIns="91440" tIns="91440" rIns="91440" bIns="91440" anchor="ctr" rtlCol="0"/>');
  assert.equal(bodyPrXml({ placeholder: true, autofit: "norm", anchor: "b" }), '<a:bodyPr anchor="b"><a:normAutofit/></a:bodyPr>');
});

test("a run with a link but no link callback throws", () => {
  assert.throws(() => txBodyXml([{ runs: [{ text: "x", link: "https://x" }] }], { lang: "en-US", fonts: { mono: "M" } }, {}), /link/);
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test`
Expected: FAIL — `Cannot find module '.../src/ooxml/text.mjs'`

- [ ] **Step 3: Implement** — `src/ooxml/text.mjs`

```js
// DrawingML text bodies. Every paragraph gets at most one <a:pPr>, always first, and every
// pPr/rPr child is written in schema order - pptxgenjs 4.0.1 writes a pPr per *run*, which
// is what dropped bullets in LibreOffice and is a likely PowerPoint Online repair trigger.

import { esc, emu, pt100, solidFill } from "./xml.mjs";

const THEME_FONTS = {
  heading: ["+mj-lt", "+mj-ea", "+mj-cs"],
  body: ["+mn-lt", "+mn-ea", "+mn-cs"],
};

function fontXml(font, ctx) {
  if (!font) return "";
  if (font === "mono") {
    const f = esc(ctx.fonts.mono);
    return `<a:latin typeface="${f}"/><a:cs typeface="${f}"/>`;
  }
  const refs = THEME_FONTS[font];
  if (!refs) throw new Error(`Unknown font role ${JSON.stringify(font)}`);
  return `<a:latin typeface="${refs[0]}"/><a:ea typeface="${refs[1]}"/><a:cs typeface="${refs[2]}"/>`;
}

export function rPrXml(run, ctx, tag = "a:rPr") {
  let attrs = ` lang="${esc(ctx.lang)}"`;
  if (run.size) attrs += ` sz="${Math.min(400000, Math.max(100, pt100(run.size)))}"`;
  if (run.b) attrs += ' b="1"';
  if (run.i) attrs += ' i="1"';
  if (run.spacing) attrs += ` spc="${pt100(run.spacing)}"`;
  attrs += ' dirty="0"';
  // Child order per CT_TextCharacterProperties: fill, ..., latin, ea, cs, sym, hlinkClick.
  let kids = "";
  if (run.color) kids += solidFill(run.color);
  kids += fontXml(run.font, ctx);
  if (run.link) {
    if (!ctx.link) throw new Error("Run has a link but the text context has no link() callback");
    kids += `<a:hlinkClick r:id="${ctx.link(run.link)}"/>`;
  }
  return kids ? `<${tag}${attrs}>${kids}</${tag}>` : `<${tag}${attrs}/>`;
}

function runsXml(runs, ctx) {
  let out = "";
  for (const run of runs) {
    const pieces = String(run.text ?? "").split("\n");
    pieces.forEach((piece, i) => {
      if (i > 0) out += `<a:br>${rPrXml({ ...run, link: undefined }, ctx)}</a:br>`;
      if (piece !== "") out += `<a:r>${rPrXml(run, ctx)}<a:t>${esc(piece)}</a:t></a:r>`;
    });
  }
  return out;
}

export function pPrXml(para, ctx, { placeholder = false } = {}) {
  const level = Math.max(0, Math.min(8, para.level || 0));
  const attrs = [];
  const kids = [];
  const bullet = para.bullet;

  if (bullet === "none" && placeholder) {
    // Body placeholders inherit bullets from the master's bodyStyle, so a plain paragraph has
    // to switch them off and drop the hanging indent explicitly.
    attrs.push('marL="0"', 'indent="0"');
  } else if ((bullet === "char" || bullet === "number") && !placeholder) {
    // Free text boxes inherit no list style, so they carry their own indent.
    attrs.push(`marL="${emu(0.25 * (level + 1))}"`, `indent="${-emu(0.2)}"`);
  }
  if (level > 0) attrs.push(`lvl="${level}"`);
  if (para.align && para.align !== "l") attrs.push(`algn="${para.align}"`);

  // Child order per CT_TextParagraphProperties: lnSpc, spcBef, spcAft, buClr, buSz, buFont, bullet.
  if (para.spaceBefore != null) kids.push(`<a:spcBef><a:spcPts val="${pt100(para.spaceBefore)}"/></a:spcBef>`);
  if (bullet === "char" && !placeholder) {
    kids.push('<a:buClr><a:schemeClr val="accent1"/></a:buClr>', '<a:buFont typeface="Arial"/>', '<a:buChar char="•"/>');
  } else if (bullet === "number") {
    const startAt = para.number && para.number !== 1 ? ` startAt="${para.number}"` : "";
    kids.push('<a:buFont typeface="+mj-lt"/>', `<a:buAutoNum type="arabicPeriod"${startAt}/>`);
  } else if (bullet === "none" && placeholder) {
    kids.push("<a:buNone/>");
  }

  if (!attrs.length && !kids.length) return "";
  const a = attrs.length ? ` ${attrs.join(" ")}` : "";
  return kids.length ? `<a:pPr${a}>${kids.join("")}</a:pPr>` : `<a:pPr${a}/>`;
}

function paraXml(para, ctx, opts) {
  const runs = para.runs || [];
  const last = runs[runs.length - 1] || {};
  // endParaRPr keeps an empty trailing line (and an empty paragraph) at the run's size/font
  // instead of PowerPoint's 18pt default.
  const end = rPrXml({ size: last.size, font: last.font }, ctx, "a:endParaRPr");
  return `<a:p>${pPrXml(para, ctx, opts)}${runsXml(runs, ctx)}${end}</a:p>`;
}

export function bodyPrXml(body = {}) {
  const attrs = [];
  if (!body.placeholder && body.wrap !== false) attrs.push('wrap="square"');
  if (body.inset != null) {
    const i = emu(body.inset);
    attrs.push(`lIns="${i}" tIns="${i}" rIns="${i}" bIns="${i}"`);
  }
  if (body.anchor) attrs.push(`anchor="${body.anchor}"`);
  if (!body.placeholder) attrs.push('rtlCol="0"');
  let kid = "";
  if (body.fontScale != null && body.fontScale < 1) {
    kid = `<a:normAutofit fontScale="${Math.round(body.fontScale * 100000)}" lnSpcReduction="10000"/>`;
  } else if (body.autofit === "norm") kid = "<a:normAutofit/>";
  else if (body.autofit === "shape") kid = "<a:spAutoFit/>";
  else if (body.autofit === "none") kid = "<a:noAutofit/>";
  const a = attrs.length ? ` ${attrs.join(" ")}` : "";
  return kid ? `<a:bodyPr${a}>${kid}</a:bodyPr>` : `<a:bodyPr${a}/>`;
}

export function txBodyXml(paras, ctx, body = {}, tag = "p:txBody") {
  const list = paras && paras.length ? paras : [{ runs: [] }];
  const ps = list.map((p) => paraXml(p, ctx, { placeholder: !!body.placeholder })).join("");
  const lst = body.lstStyle ? `<a:lstStyle>${body.lstStyle}</a:lstStyle>` : "<a:lstStyle/>";
  return `<${tag}>${bodyPrXml(body)}${lst}${ps}</${tag}>`;
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `npm test`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/ooxml/text.mjs test/text.test.mjs
git commit -m "Add schema-ordered text body writer with one pPr per paragraph

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 6: Shape builders

**Files:**
- Create: `src/ooxml/shapes.mjs`
- Test: `test/shapes.test.mjs`

**Interfaces:**
- Consumes: `esc`, `emu`, `emuSize`, `solidFill`, `colorXml`, `XML_DECL`, `NS_PML` (Task 1); `txBodyXml` (Task 5).
- Produces:
  - `Box = { x, y, w, h }` (inches).
  - `TextShape = { kind: "text", x, y, w, h, geometry?: "rect"|"roundRect"|"ellipse", radius?: number /*in*/, fill?: Color, line?: { color: Color, width: number /*pt*/, dash?: "dash" } | null, shadow?: boolean, paras?: Para[], anchor?: "t"|"ctr"|"b", inset?: number /*in*/ }`
  - `SP_TREE_HEADER: string`, `xfrmXml(box, tag = "a:xfrm")`, `phXml({ type?, idx?, sz? })`
  - `placeholderXml({ id, name, ph, box?, paras, ctx, body? }): string` — `box` omitted ⇒ `<p:spPr/>` (inherit layout geometry).
  - `shapeXml(shape: TextShape, id, ctx): string` — rect with no fill and no line ⇒ `txBox="1"`; missing line ⇒ `<a:ln><a:noFill/></a:ln>`.
  - `pictureXml({ id, name, descr, rId, box }): string`
  - `graphicFrameXml({ id, name, box, uri, inner }): string`
  - `slideXml(shapesXml): string` — full `p:sld` part with `p:clrMapOvr`.

- [ ] **Step 1: Write the failing test** — `test/shapes.test.mjs`

```js
import { test } from "node:test";
import assert from "node:assert/strict";
import { placeholderXml, shapeXml, pictureXml, graphicFrameXml, slideXml, phXml } from "../src/ooxml/shapes.mjs";
import { parseXml } from "./helpers.mjs";

const ctx = { lang: "en-US", fonts: { mono: "M" } };
const box = { x: 1, y: 0.5, w: 2, h: 1 };

test("placeholder without a box inherits geometry from the layout", () => {
  const xml = placeholderXml({ id: 2, name: "Title 1", ph: { type: "title" }, paras: [{ runs: [{ text: "Hi" }] }], ctx });
  assert.match(xml, /^<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"\/><p:cNvSpPr><a:spLocks noGrp="1"\/><\/p:cNvSpPr><p:nvPr><p:ph type="title"\/><\/p:nvPr><\/p:nvSpPr><p:spPr\/><p:txBody><a:bodyPr\/>/);
});

test("placeholder with a box writes a slide-level xfrm", () => {
  const xml = placeholderXml({ id: 3, name: "Body", ph: { idx: 1 }, box, paras: [], ctx, body: { fontScale: 0.9 } });
  assert.match(xml, /<p:ph idx="1"\/>/);
  assert.match(xml, /<p:spPr><a:xfrm><a:off x="914400" y="457200"\/><a:ext cx="1828800" cy="914400"\/><\/a:xfrm><\/p:spPr>/);
  assert.match(xml, /<a:normAutofit fontScale="90000"/);
});

test("phXml writes only the attributes that are set", () => {
  assert.equal(phXml({ type: "body", idx: 1, sz: "half" }), '<p:ph type="body" idx="1" sz="half"/>');
  assert.equal(phXml({ type: "ctrTitle" }), '<p:ph type="ctrTitle"/>');
});

test("plain text shape is a txBox with no fill and no line", () => {
  const xml = shapeXml({ kind: "text", ...box, paras: [{ runs: [{ text: "x" }] }] }, 4, ctx);
  assert.match(xml, /<p:cNvSpPr txBox="1"\/>/);
  assert.match(xml, /<a:prstGeom prst="rect"><a:avLst\/><\/a:prstGeom><a:noFill\/><a:ln><a:noFill\/><\/a:ln><\/p:spPr>/);
  assert.match(xml, /<a:bodyPr wrap="square" anchor="t" rtlCol="0"\/>/);
});

test("card shape: roundRect radius, fill, dashed line, shadow, in spPr schema order", () => {
  const xml = shapeXml({ kind: "text", x: 0, y: 0, w: 2, h: 1, geometry: "roundRect", radius: 0.1, fill: "lt2", line: { color: "accent4", width: 1, dash: "dash" }, shadow: true }, 5, ctx);
  assert.match(xml, /<p:cNvSpPr\/>/);
  assert.match(xml, /<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"\/><\/a:avLst><\/a:prstGeom><a:solidFill><a:schemeClr val="bg2"\/><\/a:solidFill><a:ln w="12700"><a:solidFill><a:schemeClr val="accent4"\/><\/a:solidFill><a:prstDash val="dash"\/><\/a:ln><a:effectLst><a:outerShdw /);
  assert.match(xml, /<a:outerShdw blurRad="114300" dist="38100" dir="5400000" algn="t" rotWithShape="0"><a:schemeClr val="tx2"><a:alpha val="16000"\/><\/a:schemeClr><\/a:outerShdw>/);
});

test("picture carries alt text as descr", () => {
  const xml = pictureXml({ id: 6, name: "Picture 6", descr: 'A "quoted" <alt>', rId: "rId3", box });
  assert.match(xml, /<p:cNvPr id="6" name="Picture 6" descr="A &quot;quoted&quot; &lt;alt&gt;"\/>/);
  assert.match(xml, /<a:blip r:embed="rId3"\/><a:stretch><a:fillRect\/><\/a:stretch>/);
});

test("graphic frame wraps graphicData with the given uri", () => {
  const xml = graphicFrameXml({ id: 7, name: "Chart 7", box, uri: "urn:x", inner: "<x/>" });
  assert.match(xml, /<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"\/><\/p:cNvGraphicFramePr>/);
  assert.match(xml, /<p:xfrm><a:off x="914400" y="457200"\/><a:ext cx="1828800" cy="914400"\/><\/p:xfrm><a:graphic><a:graphicData uri="urn:x"><x\/><\/a:graphicData><\/a:graphic>/);
});

test("slideXml is a well-formed slide with clrMapOvr", () => {
  const xml = slideXml(shapeXml({ kind: "text", ...box, paras: [] }, 2, ctx));
  parseXml(xml);
  assert.match(xml, /<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""\/>/);
  assert.match(xml, /<\/p:cSld><p:clrMapOvr><a:masterClrMapping\/><\/p:clrMapOvr><\/p:sld>$/);
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test`
Expected: FAIL — `Cannot find module '.../src/ooxml/shapes.mjs'`

- [ ] **Step 3: Implement** — `src/ooxml/shapes.mjs`

```js
// PresentationML shape builders: placeholders, free text/geometry shapes, pictures and
// graphic frames (tables, charts). Geometry arrives in inches; ids come from the caller's
// per-part counter so they stay unique within a slide.

import { XML_DECL, NS_PML, esc, emu, emuSize, solidFill, colorXml } from "./xml.mjs";
import { txBodyXml } from "./text.mjs";

export const SP_TREE_HEADER =
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

export function xfrmXml(box, tag = "a:xfrm") {
  return `<${tag}><a:off x="${emu(box.x)}" y="${emu(box.y)}"/><a:ext cx="${emuSize(box.w)}" cy="${emuSize(box.h)}"/></${tag}>`;
}

export function phXml(ph) {
  const type = ph.type ? ` type="${ph.type}"` : "";
  const idx = ph.idx != null ? ` idx="${ph.idx}"` : "";
  const sz = ph.sz ? ` sz="${ph.sz}"` : "";
  return `<p:ph${type}${idx}${sz}/>`;
}

export function placeholderXml({ id, name, ph, box, paras, ctx, body = {} }) {
  // No xfrm means "inherit position and size from the layout" - that's what makes Reset Slide
  // and layout switching work.
  const spPr = box ? `<p:spPr>${xfrmXml(box)}</p:spPr>` : "<p:spPr/>";
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
    `<p:nvPr>${phXml(ph)}</p:nvPr></p:nvSpPr>${spPr}${txBodyXml(paras, ctx, { ...body, placeholder: true })}</p:sp>`
  );
}

function geometryXml(shape) {
  if (shape.geometry === "roundRect") {
    // roundRect's adj is the corner radius as a fraction (x100000) of the shorter side.
    const side = Math.max(0.0001, Math.min(shape.w, shape.h));
    const adj = Math.round(Math.min(50000, Math.max(0, ((shape.radius ?? 0.1) / side) * 100000)));
    return `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val ${adj}"/></a:avLst></a:prstGeom>`;
  }
  return `<a:prstGeom prst="${shape.geometry === "ellipse" ? "ellipse" : "rect"}"><a:avLst/></a:prstGeom>`;
}

function lineXml(line) {
  // Explicit "no line": without a p:style there is no line anyway in PowerPoint, but Keynote
  // and Google Slides apply their own defaults to a shape with no a:ln at all.
  if (!line) return "<a:ln><a:noFill/></a:ln>";
  const dash = line.dash ? `<a:prstDash val="${line.dash}"/>` : "";
  return `<a:ln w="${Math.round(line.width * 12700)}">${solidFill(line.color)}${dash}</a:ln>`;
}

const SHADOW =
  '<a:effectLst><a:outerShdw blurRad="114300" dist="38100" dir="5400000" algn="t" rotWithShape="0">' +
  `${colorXml("dk2", 0.16)}</a:outerShdw></a:effectLst>`;

export function shapeXml(shape, id, ctx) {
  const geometry = shape.geometry || "rect";
  const isTextBox = geometry === "rect" && !shape.fill && !shape.line;
  const nvSpPr =
    `<p:nvSpPr><p:cNvPr id="${id}" name="${isTextBox ? "TextBox" : "Shape"} ${id}"/>` +
    `<p:cNvSpPr${isTextBox ? ' txBox="1"' : ""}/><p:nvPr/></p:nvSpPr>`;
  // spPr child order: xfrm, geometry, fill, ln, effectLst.
  const spPr =
    `<p:spPr>${xfrmXml(shape)}${geometryXml({ ...shape, geometry })}` +
    `${shape.fill ? solidFill(shape.fill) : "<a:noFill/>"}${lineXml(shape.line)}${shape.shadow ? SHADOW : ""}</p:spPr>`;
  const body = txBodyXml(shape.paras || [], ctx, { anchor: shape.anchor || "t", inset: shape.inset });
  return `<p:sp>${nvSpPr}${spPr}${body}</p:sp>`;
}

export function pictureXml({ id, name, descr, rId, box }) {
  const alt = descr ? ` descr="${esc(descr)}"` : "";
  return (
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${esc(name)}"${alt}/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr>${xfrmXml(box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
  );
}

export function graphicFrameXml({ id, name, box, uri, inner }) {
  return (
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${esc(name)}"/>` +
    `<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
    `${xfrmXml(box, "p:xfrm")}<a:graphic><a:graphicData uri="${uri}">${inner}</a:graphicData></a:graphic></p:graphicFrame>`
  );
}

export function slideXml(shapesXml) {
  return (
    `${XML_DECL}<p:sld ${NS_PML}><p:cSld><p:spTree>${SP_TREE_HEADER}${shapesXml}</p:spTree></p:cSld>` +
    "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
  );
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `npm test`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/ooxml/shapes.mjs test/shapes.test.mjs
git commit -m "Add PresentationML shape builders (placeholders, shapes, pictures, frames)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 7: Slide geometry, master and layouts

**Files:**
- Create: `src/geometry.mjs`
- Create: `src/ooxml/master.mjs`
- Test: `test/master.test.mjs`

**Interfaces:**
- Consumes: `placeholderXml`, `pictureXml`, `SP_TREE_HEADER` (Task 6); `solidFill`, `esc`, `XML_DECL`, `NS_PML` (Task 1).
- Produces:
  - `slideGeometry({ width, height, marginX }): Geometry` with `width, height, marginX, contentW, bottomMargin (0.35)` and boxes `eyebrow, title, body, coverTitle, coverSubtitle, sectionTitle, sectionBody, centeredTitle, centeredBody` (all `Box`).
  - `layoutDefs(geometry): LayoutDef[]` in this order with these keys: `title, content, section, twoContent, titleOnly, centered, blank`. `LayoutDef = { key, name, type, placeholders: [{ role: "title"|"subtitle"|"body"|"body2", ph, name, box, anchor, lstStyle?, prompt }] }`.
  - `layoutXml(def, ctx): string`
  - `masterXml({ geometry, ctx, layoutIds: [{ id, rId }], logo: { rId, box } | null }): string`

- [ ] **Step 1: Write the failing test** — `test/master.test.mjs`

```js
import { test } from "node:test";
import assert from "node:assert/strict";
import { slideGeometry } from "../src/geometry.mjs";
import { layoutDefs, layoutXml, masterXml } from "../src/ooxml/master.mjs";
import { parseXml, byLocal } from "./helpers.mjs";

const g = slideGeometry({ width: 13.333, height: 7.5, marginX: 0.6 });
const ctx = { lang: "en-US", fonts: { mono: "M" } };

test("geometry reserves the eyebrow band above the title", () => {
  assert.equal(g.eyebrow.x, 0.6);
  assert.equal(g.eyebrow.y, 0.45);
  assert.equal(g.eyebrow.h, 0.3);
  assert.ok(Math.abs(g.eyebrow.w - 12.133) < 1e-9);
  assert.equal(g.title.y, 0.8);
  assert.equal(g.body.y, 1.6);
  assert.ok(Math.abs(g.body.h - 5.55) < 1e-9);
  assert.ok(Math.abs(g.coverTitle.y - 2.7) < 1e-9);
  assert.ok(Math.abs(g.coverSubtitle.y - 4.1) < 1e-9);
});

test("layout catalogue has the expected keys, types and placeholders", () => {
  const defs = layoutDefs(g);
  assert.deepEqual(defs.map((d) => [d.key, d.type]), [
    ["title", "title"], ["content", "obj"], ["section", "secHead"], ["twoContent", "twoObj"],
    ["titleOnly", "titleOnly"], ["centered", "cust"], ["blank", "blank"],
  ]);
  const byKey = Object.fromEntries(defs.map((d) => [d.key, d]));
  assert.deepEqual(byKey.title.placeholders.map((p) => p.ph), [{ type: "ctrTitle" }, { type: "subTitle", idx: 1 }]);
  assert.deepEqual(byKey.content.placeholders.map((p) => p.role), ["title", "body"]);
  assert.deepEqual(byKey.blank.placeholders, []);
});

test("every layout part is well-formed and typed", () => {
  for (const def of layoutDefs(g)) {
    const xml = layoutXml(def, ctx);
    const doc = parseXml(xml, def.name);
    assert.equal(doc.documentElement.getAttribute("type"), def.type);
    assert.equal(byLocal(doc, "cSld")[0].getAttribute("name"), def.name);
    assert.equal(byLocal(doc, "ph").length, def.placeholders.length);
  }
});

test("cover title style lives in the layout's lstStyle", () => {
  const def = layoutDefs(g).find((d) => d.key === "title");
  const xml = layoutXml(def, ctx);
  assert.match(xml, /<a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="4000" b="1"><a:solidFill><a:schemeClr val="accent2"\/><\/a:solidFill><\/a:defRPr><\/a:lvl1pPr><\/a:lstStyle>/);
});

test("master: clrMap, layout id list, txStyles with theme fonts and accent bullets, logo", () => {
  const xml = masterXml({ geometry: g, ctx, layoutIds: [{ id: 2147483649, rId: "rId2" }], logo: { rId: "rId9", box: { x: 0.5, y: 6.94, w: 0.98, h: 0.27 } } });
  parseXml(xml);
  assert.match(xml, /<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1"/);
  assert.match(xml, /<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId2"\/><\/p:sldLayoutIdLst>/);
  assert.match(xml, /<p:titleStyle><a:lvl1pPr[^>]*>.*<a:defRPr sz="2600" b="1" kern="1200">.*<a:latin typeface="\+mj-lt"\/>/);
  assert.match(xml, /<p:bodyStyle><a:lvl1pPr marL="228600" indent="-182880"[^>]*>.*<a:buClr><a:schemeClr val="accent1"\/><\/a:buClr><a:buFont typeface="Arial"\/><a:buChar char="•"\/><a:defRPr sz="1300"/);
  assert.match(xml, /<a:blip r:embed="rId9"\/>/);
  // cSld, clrMap, sldLayoutIdLst, txStyles - in schema order.
  const order = ["<p:cSld>", "<p:clrMap ", "<p:sldLayoutIdLst>", "<p:txStyles>"].map((s) => xml.indexOf(s));
  assert.deepEqual([...order].sort((a, b) => a - b), order);
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test`
Expected: FAIL — `Cannot find module '.../src/geometry.mjs'`

- [ ] **Step 3: Implement** — `src/geometry.mjs`

```js
// Slide geometry in inches, shared by the layout engine (where free shapes go) and the slide
// master (where placeholders sit), so the two always line up. Layouts always reserve the
// eyebrow band, like a real template, so titles sit at the same height on every slide.

export function slideGeometry({ width, height, marginX }) {
  const contentW = width - marginX * 2;
  const box = (y, h) => ({ x: marginX, y, w: contentW, h });
  const bottomMargin = 0.35;
  return {
    width,
    height,
    marginX,
    contentW,
    bottomMargin,
    eyebrow: box(0.45, 0.3),
    title: box(0.8, 0.6),
    body: box(1.6, height - bottomMargin - 1.6),
    coverTitle: box(height * 0.36, 1.3),
    coverSubtitle: box(height * 0.36 + 1.4, 0.8),
    sectionTitle: box(height * 0.36, 1.0),
    sectionBody: box(height * 0.36 + 1.1, 1.0),
    centeredTitle: box(height / 2 - 1.0, 0.6),
    centeredBody: box(height / 2 - 0.2, 2.0),
  };
}
```

- [ ] **Step 4: Implement** — `src/ooxml/master.mjs`

```js
// Slide master + the layout catalogue. Layouts carry real types (title, obj, secHead, ...) so
// PowerPoint, Google Slides and Keynote all recognize them, and their placeholders carry the
// geometry, anchors and autofit that slide placeholders inherit.

import { XML_DECL, NS_PML, esc, solidFill } from "./xml.mjs";
import { SP_TREE_HEADER, placeholderXml, pictureXml } from "./shapes.mjs";

// Level-1 list style override for one layout placeholder (e.g. the cover's big centered title).
function lvl1({ algn, sz, b, color, buNone }) {
  const attrs = [buNone ? 'marL="0" indent="0"' : "", algn ? `algn="${algn}"` : ""].filter(Boolean).join(" ");
  const rAttrs = `${sz ? ` sz="${sz * 100}"` : ""}${b ? ' b="1"' : ""}`;
  const defRPr = color ? `<a:defRPr${rAttrs}>${solidFill(color)}</a:defRPr>` : `<a:defRPr${rAttrs}/>`;
  return `<a:lvl1pPr${attrs ? ` ${attrs}` : ""}>${buNone ? "<a:buNone/>" : ""}${defRPr}</a:lvl1pPr>`;
}

const TITLE_PROMPT = "Click to edit Master title style";
const BODY_PROMPT = "Click to edit Master text styles";

export function layoutDefs(g) {
  const title = { role: "title", ph: { type: "title" }, name: "Title 1", box: g.title, anchor: "b", prompt: TITLE_PROMPT };
  const body = { role: "body", ph: { idx: 1 }, name: "Content Placeholder 2", box: g.body, anchor: "t", prompt: BODY_PROMPT };
  const half = (g.contentW - 0.3) / 2;
  return [
    {
      key: "title", name: "Title Slide", type: "title",
      placeholders: [
        { role: "title", ph: { type: "ctrTitle" }, name: "Title 1", box: g.coverTitle, anchor: "ctr", prompt: TITLE_PROMPT,
          lstStyle: lvl1({ algn: "ctr", sz: 40, b: true, color: "accent2" }) },
        { role: "subtitle", ph: { type: "subTitle", idx: 1 }, name: "Subtitle 2", box: g.coverSubtitle, anchor: "t", prompt: "Click to edit Master subtitle style",
          lstStyle: lvl1({ algn: "ctr", sz: 18, buNone: true }) },
      ],
    },
    { key: "content", name: "Title and Content", type: "obj", placeholders: [title, body] },
    {
      key: "section", name: "Section Header", type: "secHead",
      placeholders: [
        { ...title, box: g.sectionTitle, lstStyle: lvl1({ sz: 34, b: true }) },
        { role: "body", ph: { type: "body", idx: 1 }, name: "Text Placeholder 2", box: g.sectionBody, anchor: "t", prompt: BODY_PROMPT,
          lstStyle: lvl1({ sz: 16, color: "accent3", buNone: true }) },
      ],
    },
    {
      key: "twoContent", name: "Two Content", type: "twoObj",
      placeholders: [
        title,
        { ...body, ph: { sz: "half", idx: 1 }, box: { ...g.body, w: half } },
        { ...body, role: "body2", ph: { sz: "half", idx: 2 }, name: "Content Placeholder 3", box: { ...g.body, x: g.marginX + half + 0.3, w: half } },
      ],
    },
    { key: "titleOnly", name: "Title Only", type: "titleOnly", placeholders: [title] },
    {
      key: "centered", name: "Centered", type: "cust",
      placeholders: [{ ...title, box: g.centeredTitle }, { ...body, box: g.centeredBody }],
    },
    { key: "blank", name: "Blank", type: "blank", placeholders: [] },
  ];
}

export function layoutXml(def, ctx) {
  const sps = def.placeholders
    .map((p, i) =>
      placeholderXml({
        id: i + 2, name: p.name, ph: p.ph, box: p.box, ctx,
        paras: [{ runs: [{ text: p.prompt }] }],
        body: { anchor: p.anchor, autofit: "norm", lstStyle: p.lstStyle },
      })
    )
    .join("");
  return (
    `${XML_DECL}<p:sldLayout ${NS_PML} type="${def.type}" preserve="1"><p:cSld name="${esc(def.name)}">` +
    `<p:spTree>${SP_TREE_HEADER}${sps}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`
  );
}

const CLR_MAP =
  '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" ' +
  'accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';

const P_ATTRS = 'algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"';

function defRPr(sz, { b = false, font = "mn" } = {}) {
  return (
    `<a:defRPr sz="${sz * 100}"${b ? ' b="1"' : ""} kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>` +
    `<a:latin typeface="+${font}-lt"/><a:ea typeface="+${font}-ea"/><a:cs typeface="+${font}-cs"/></a:defRPr>`
  );
}

const TITLE_STYLE =
  `<p:titleStyle><a:lvl1pPr ${P_ATTRS}><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef>` +
  `<a:buNone/>${defRPr(26, { b: true, font: "mj" })}</a:lvl1pPr></p:titleStyle>`;

// Body levels 1-3: 0.25"/0.5"/0.75" left margin with a 0.2" hanging bullet, 13/12/11 pt,
// accent1 bullets - the deck's brand color comes from the master, not from each slide.
const BODY_LEVELS = [[228600, 13], [457200, 12], [685800, 11]];
const BODY_STYLE = `<p:bodyStyle>${BODY_LEVELS.map(
  ([marL, sz], i) =>
    `<a:lvl${i + 1}pPr marL="${marL}" indent="-182880" ${P_ATTRS}><a:lnSpc><a:spcPct val="100000"/></a:lnSpc>` +
    '<a:spcBef><a:spcPts val="600"/></a:spcBef><a:buClr><a:schemeClr val="accent1"/></a:buClr><a:buFont typeface="Arial"/><a:buChar char="•"/>' +
    `${defRPr(sz)}</a:lvl${i + 1}pPr>`
).join("")}</p:bodyStyle>`;

const OTHER_STYLE =
  `<p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr marL="0" ${P_ATTRS}>${defRPr(18)}</a:lvl1pPr></p:otherStyle>`;

export function masterXml({ geometry: g, ctx, layoutIds, logo }) {
  const title = placeholderXml({
    id: 2, name: "Title Placeholder 1", ph: { type: "title" }, box: g.title, ctx,
    paras: [{ runs: [{ text: TITLE_PROMPT }] }], body: { anchor: "b", autofit: "norm" },
  });
  const body = placeholderXml({
    id: 3, name: "Text Placeholder 2", ph: { type: "body", idx: 1 }, box: g.body, ctx,
    paras: [
      { bullet: "char", runs: [{ text: BODY_PROMPT }] },
      { bullet: "char", level: 1, runs: [{ text: "Second level" }] },
      { bullet: "char", level: 2, runs: [{ text: "Third level" }] },
    ],
    body: { anchor: "t", autofit: "norm" },
  });
  const pic = logo ? pictureXml({ id: 4, name: "Logo", descr: "Logo", rId: logo.rId, box: logo.box }) : "";
  const ids = layoutIds.map(({ id, rId }) => `<p:sldLayoutId id="${id}" r:id="${rId}"/>`).join("");
  return (
    `${XML_DECL}<p:sldMaster ${NS_PML}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>` +
    `<p:spTree>${SP_TREE_HEADER}${title}${body}${pic}</p:spTree></p:cSld>${CLR_MAP}` +
    `<p:sldLayoutIdLst>${ids}</p:sldLayoutIdLst><p:txStyles>${TITLE_STYLE}${BODY_STYLE}${OTHER_STYLE}</p:txStyles></p:sldMaster>`
  );
}
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `npm test`
Expected: PASS

- [ ] **Step 6: Commit**

```bash
git add src/geometry.mjs src/ooxml/master.mjs test/master.test.mjs
git commit -m "Add slide geometry, slide master and typed layout catalogue

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 8: Presentation parts and `writePptx` (text, placeholders, images)

**Files:**
- Create: `src/ooxml/presentation.mjs`
- Create: `src/ooxml/index.mjs`
- Create: `test/fixtures/kitchen-sink/public/pic.png` (4×2 px PNG)
- Create: `test/deck.mjs` (shared `DeckOptions` fixture for writer tests)
- Test: `test/writer.test.mjs`

**Interfaces:**
- Consumes: everything from Tasks 1–7.
- Produces:
  - `presentationXml({ masterId, masterRId, notesMasterRId, slides: [{ id, rId }], width, height })`, `presPropsXml()`, `viewPropsXml()`, `EMPTY_TABLE_STYLES` (replaced in Task 9), `appXml({ slides, notes })`, `coreXml({ title, author, date: Date | null })`.
  - `SlideSpec = { layout: "title"|"content"|"section"|"twoContent"|"titleOnly"|"centered"|"blank", placeholders?: { title?, subtitle?, body?, body2?: Para[], titleBox?, bodyBox?: Box, bodyFontScale?: number }, shapes?: Shape[], notes?: string }` where `Shape` is `TextShape` (Task 6) or `{ kind: "image", x, y, w, h, path, alt }` (tables and charts added in Tasks 9–10).
  - `DeckOptions = { geometry, theme: ThemeSlots, fonts: { heading, body, mono }, lang, title, author, timestamps, logo: { path, box } | null }`.
  - `async writePptx(slides: SlideSpec[], deck: DeckOptions): Promise<Buffer>`.
  - Internal (extended by Tasks 9–11): `shapeToXml(shape, id, env)` with `env = { pkg, part, media, ctx, counters: { chart: 0 } }`.

- [ ] **Step 1: Create the fixture picture**

```bash
mkdir -p test/fixtures/kitchen-sink/public
node -e "require('node:fs').writeFileSync('test/fixtures/kitchen-sink/public/pic.png', Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAIAAADwyuo0AAAAEklEQVR4nGOUarZmgAEmBiQAABRtANyhctckAAAAAElFTkSuQmCC', 'base64'))"
```

- [ ] **Step 2: Write the shared deck fixture** — `test/deck.mjs`

Kept out of `*.test.mjs` so other test files can import it without re-running the writer tests.

```js
// DeckOptions shared by the writer tests (Tasks 8-11).
import { fileURLToPath } from "node:url";
import { slideGeometry } from "../src/geometry.mjs";
import { themeColors } from "../src/colors.mjs";

export const PIC = fileURLToPath(new URL("./fixtures/kitchen-sink/public/pic.png", import.meta.url));
export const geometry = slideGeometry({ width: 13.333, height: 7.5, marginX: 0.6 });

export const DECK = {
  geometry,
  theme: themeColors({ accent: "1A833B", accentDark: "14532D", accent2: "1A589F", panel: "EEF4F0", border: "D9E3DB", text: "1F2937", muted: "6B7280" }),
  fonts: { heading: "Plus Jakarta Sans", body: "Inter", mono: "JetBrains Mono" },
  lang: "en-US",
  title: "Writer test",
  author: "Tester",
  timestamps: false,
  logo: { path: PIC, box: { x: 0.5, y: 6.94, w: 0.54, h: 0.27 } },
};
```

- [ ] **Step 3: Write the failing test** — `test/writer.test.mjs`

```js
import { test } from "node:test";
import assert from "node:assert/strict";
import { writePptx } from "../src/ooxml/index.mjs";
import { DECK, PIC, geometry } from "./deck.mjs";
import { lintPackage } from "./lint.mjs";
import { unzip, byLocal } from "./helpers.mjs";

const t = (text) => [{ runs: [{ text, font: "heading" }] }];

const SLIDES = [
  { layout: "title", placeholders: { title: t("Cover"), subtitle: [{ runs: [{ text: "Sub" }] }] }, shapes: [] },
  {
    layout: "content",
    placeholders: {
      title: t("Bullets"),
      body: [
        { bullet: "char", runs: [{ text: "One " }, { text: "bold", b: true, color: "dk2" }] },
        { bullet: "number", runs: [{ text: "Two" }] },
      ],
      bodyFontScale: 0.85,
    },
    shapes: [{ kind: "text", ...geometry.eyebrow, paras: [{ runs: [{ text: "EYEBROW", link: "https://example.com/?a=1&b=2" }] }] }],
  },
  {
    layout: "titleOnly",
    placeholders: { title: t("Pictures") },
    shapes: [
      { kind: "image", x: 1, y: 2, w: 2, h: 1, path: PIC, alt: "A picture" },
      { kind: "image", x: 4, y: 2, w: 2, h: 1, path: PIC, alt: "Same picture again" },
    ],
  },
  {
    layout: "centered",
    placeholders: { title: t("Centered"), titleBox: { x: 0.6, y: 3, w: 12.133, h: 0.6 }, body: [{ bullet: "none", align: "ctr", runs: [{ text: "Questions?" }] }], bodyBox: { x: 0.6, y: 3.8, w: 12.133, h: 0.5 } },
    shapes: [],
  },
  { layout: "blank", shapes: [{ kind: "text", x: 1, y: 1, w: 3, h: 1, geometry: "roundRect", fill: "lt2", shadow: true, paras: [] }] },
];

async function build(slides = SLIDES, deck = DECK) {
  const buffer = await writePptx(slides, deck);
  return { buffer, z: await unzip(buffer) };
}

function layoutNameOf(z, slideNo) {
  const rels = z.text(`ppt/slides/_rels/slide${slideNo}.xml.rels`);
  const target = /Type="[^"]+\/slideLayout" Target="..\/slideLayouts\/(slideLayout\d+\.xml)"/.exec(rels)[1];
  return byLocal(z.xml(`ppt/slideLayouts/${target}`), "cSld")[0].getAttribute("name");
}

test("package passes lint", async () => {
  const { buffer } = await build();
  assert.deepEqual(await lintPackage(buffer), []);
});

test("each slide uses the requested layout", async () => {
  const { z } = await build();
  assert.deepEqual([1, 2, 3, 4, 5].map((n) => layoutNameOf(z, n)), ["Title Slide", "Title and Content", "Title Only", "Centered", "Blank"]);
});

test("titles live in real title placeholders that inherit geometry", async () => {
  const { z } = await build();
  const s1 = z.text("ppt/slides/slide1.xml");
  assert.match(s1, /<p:ph type="ctrTitle"\/><\/p:nvPr><\/p:nvSpPr><p:spPr\/>.*<a:t>Cover<\/a:t>/);
  assert.match(s1, /<p:ph type="subTitle" idx="1"\/>.*<a:t>Sub<\/a:t>/);
  const s2 = z.text("ppt/slides/slide2.xml");
  assert.match(s2, /<p:ph type="title"\/>.*<a:t>Bullets<\/a:t>/);
  assert.match(s2, /<p:ph idx="1"\/>.*<a:normAutofit fontScale="85000"/);
});

test("centered slide overrides title and body position on the slide", async () => {
  const { z } = await build();
  assert.match(z.text("ppt/slides/slide4.xml"), /<p:ph type="title"\/><\/p:nvPr><\/p:nvSpPr><p:spPr><a:xfrm><a:off x="548640" y="2743200"\/>/);
});

test("hyperlinks become external relationships", async () => {
  const { z } = await build();
  assert.match(z.text("ppt/slides/_rels/slide2.xml.rels"), /Target="https:\/\/example.com\/\?a=1&amp;b=2" TargetMode="External"/);
  assert.match(z.text("ppt/slides/slide2.xml"), /<a:hlinkClick r:id="rId2"\/>/);
});

test("identical images are stored once and carry alt text", async () => {
  const { z } = await build();
  assert.deepEqual(z.names.filter((n) => n.startsWith("ppt/media/")), ["ppt/media/image1.png"]);
  assert.match(z.text("ppt/slides/slide3.xml"), /descr="A picture"/);
  assert.match(z.text("ppt/slideMasters/slideMaster1.xml"), /name="Logo" descr="Logo"/);
});

test("presentation ids, size and theme", async () => {
  const { z } = await build();
  const pres = z.text("ppt/presentation.xml");
  assert.match(pres, /<p:sldMasterId id="2147483648" r:id="rId1"\/>/);
  assert.deepEqual([...pres.matchAll(/<p:sldId id="(\d+)"/g)].map((m) => Number(m[1])), [256, 257, 258, 259, 260]);
  assert.match(pres, /<p:sldSz cx="12191695" cy="6858000"\/><p:notesSz cx="6858000" cy="9144000"\/>/);
  assert.match(z.text("ppt/slideMasters/slideMaster1.xml"), /<p:sldLayoutId id="2147483649" /);
  assert.match(z.text("ppt/theme/theme1.xml"), /<a:accent1><a:srgbClr val="1A833B"\/><\/a:accent1>/);
});

test("core.xml has no dates unless timestamps are on", async () => {
  const off = await build();
  assert.doesNotMatch(off.z.text("docProps/core.xml"), /dcterms:created/);
  assert.match(off.z.text("docProps/core.xml"), /<dc:title>Writer test<\/dc:title><dc:creator>Tester<\/dc:creator>/);
  const on = await build(SLIDES, { ...DECK, timestamps: true });
  assert.match(on.z.text("docProps/core.xml"), /<dcterms:created xsi:type="dcterms:W3CDTF">\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ<\/dcterms:created>/);
});

test("same input gives byte-identical output", async () => {
  const a = await writePptx(SLIDES, DECK);
  const b = await writePptx(SLIDES, DECK);
  assert.ok(a.equals(b));
});

test("unknown layouts and shape kinds are rejected", async () => {
  await assert.rejects(writePptx([{ layout: "nope" }], DECK), /unknown layout "nope"/);
  await assert.rejects(writePptx([{ layout: "blank", shapes: [{ kind: "video" }] }], DECK), /Unknown shape kind "video"/);
});
```

- [ ] **Step 4: Run test to verify it fails**

Run: `npm test`
Expected: FAIL — `Cannot find module '.../src/ooxml/index.mjs'`

- [ ] **Step 5: Implement** — `src/ooxml/presentation.mjs`

```js
// presentation.xml plus the small package-level parts PowerPoint expects beside it.

import { XML_DECL, NS_PML, esc, emuSize } from "./xml.mjs";

const DEFAULT_TEXT_STYLE =
  '<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>' +
  '<a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">' +
  '<a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>' +
  '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:defaultTextStyle>';

// PowerPoint only accepts slide sizes between 1" and 56".
const clampSize = (inches) => Math.min(51206400, Math.max(914400, emuSize(inches)));

export function presentationXml({ masterId, masterRId, notesMasterRId, slides, width, height }) {
  // Child order per CT_Presentation: sldMasterIdLst, notesMasterIdLst, sldIdLst, sldSz,
  // notesSz, defaultTextStyle. (pptxgenjs writes sldIdLst before notesMasterIdLst.)
  const notes = notesMasterRId ? `<p:notesMasterIdLst><p:notesMasterId r:id="${notesMasterRId}"/></p:notesMasterIdLst>` : "";
  const sldIds = slides.length ? `<p:sldIdLst>${slides.map((s) => `<p:sldId id="${s.id}" r:id="${s.rId}"/>`).join("")}</p:sldIdLst>` : "";
  return (
    `${XML_DECL}<p:presentation ${NS_PML}><p:sldMasterIdLst><p:sldMasterId id="${masterId}" r:id="${masterRId}"/></p:sldMasterIdLst>` +
    `${notes}${sldIds}<p:sldSz cx="${clampSize(width)}" cy="${clampSize(height)}"/><p:notesSz cx="6858000" cy="9144000"/>` +
    `${DEFAULT_TEXT_STYLE}</p:presentation>`
  );
}

export function presPropsXml() {
  return `${XML_DECL}<p:presentationPr ${NS_PML}/>`;
}

export function viewPropsXml() {
  return (
    `${XML_DECL}<p:viewPr ${NS_PML}><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr>` +
    '<p:gridSpacing cx="76200" cy="76200"/></p:viewPr>'
  );
}

// Replaced by table.mjs's tableStylesXml() once deck table styles exist (Task 9).
export const EMPTY_TABLE_STYLES = `${XML_DECL}<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`;

export function appXml({ slides, notes }) {
  return (
    `${XML_DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ` +
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    `<Application>slidev-pptx-export</Application><PresentationFormat>Custom</PresentationFormat>` +
    `<Slides>${slides}</Slides><Notes>${notes}</Notes></Properties>`
  );
}

// Dates are left out unless asked for, so two exports of the same deck are byte-identical.
export function coreXml({ title, author, date }) {
  const iso = date ? date.toISOString().replace(/\.\d{3}Z$/, "Z") : null;
  const dates = iso
    ? `<dcterms:created xsi:type="dcterms:W3CDTF">${iso}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${iso}</dcterms:modified>`
    : "";
  return (
    `${XML_DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
    'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    `<dc:title>${esc(title)}</dc:title><dc:creator>${esc(author)}</dc:creator><cp:lastModifiedBy>${esc(author)}</cp:lastModifiedBy>` +
    `<cp:revision>1</cp:revision>${dates}</cp:coreProperties>`
  );
}
```

- [ ] **Step 6: Implement** — `src/ooxml/index.mjs`

```js
// writePptx(): SlideSpec[] -> .pptx bytes. This file owns part naming, relationship wiring
// and id allocation; the XML of each part comes from the modules next to it. It knows nothing
// about Slidev or markdown - see ../layout.mjs for that.

import { readFileSync } from "node:fs";
import { extname } from "node:path";
import { createHash } from "node:crypto";
import { Package, REL, CT } from "./package.mjs";
import { themeXml } from "./theme.mjs";
import { layoutDefs, layoutXml, masterXml } from "./master.mjs";
import { placeholderXml, shapeXml, pictureXml, slideXml } from "./shapes.mjs";
import { presentationXml, presPropsXml, viewPropsXml, EMPTY_TABLE_STYLES, appXml, coreXml } from "./presentation.mjs";

const PRES = "/ppt/presentation.xml";
const MASTER = "/ppt/slideMasters/slideMaster1.xml";
const THEME = "/ppt/theme/theme1.xml";

// Each distinct image is stored once, however many slides (or the master) use it.
class MediaStore {
  constructor(pkg) {
    this.pkg = pkg;
    this.parts = new Map();
  }

  add(path) {
    const buf = readFileSync(path);
    const ext = extname(path).slice(1).toLowerCase().replace("jpeg", "jpg");
    const key = `${createHash("sha1").update(buf).digest("hex")}.${ext}`;
    if (!this.parts.has(key)) {
      const part = `/ppt/media/image${this.parts.size + 1}.${ext}`;
      this.pkg.addMedia(part, buf);
      this.parts.set(key, part);
    }
    return this.parts.get(key);
  }
}

async function shapeToXml(shape, id, env) {
  const { pkg, part, media, ctx } = env;
  switch (shape.kind) {
    case "text":
      return shapeXml(shape, id, ctx);
    case "image": {
      const rId = pkg.rel(part, media.add(shape.path), REL.image);
      return pictureXml({ id, name: `Picture ${id}`, descr: shape.alt, rId, box: shape });
    }
    default:
      throw new Error(`Unknown shape kind ${JSON.stringify(shape.kind)}`);
  }
}

export async function writePptx(slides, deck) {
  const pkg = new Package();
  const media = new MediaStore(pkg);
  const ctx = { lang: deck.lang || "en-US", fonts: deck.fonts };
  const counters = { chart: 0 };
  // Slide masters and layouts share one id space that starts at 2^31.
  let nextMasterId = 2147483648;

  pkg.rel("", PRES, REL.officeDocument);
  pkg.rel("", "/docProps/core.xml", REL.coreProps);
  pkg.rel("", "/docProps/app.xml", REL.extProps);

  pkg.addXml(THEME, CT.theme, themeXml({ name: deck.title || "Deck", colors: deck.theme, fonts: deck.fonts }));
  const masterRId = pkg.rel(PRES, MASTER, REL.slideMaster);
  const masterId = nextMasterId++;
  pkg.rel(MASTER, THEME, REL.theme);
  const layouts = new Map();
  const layoutIds = [];
  layoutDefs(deck.geometry).forEach((def, i) => {
    const part = `/ppt/slideLayouts/slideLayout${i + 1}.xml`;
    pkg.addXml(part, CT.slideLayout, layoutXml(def, ctx));
    pkg.rel(part, MASTER, REL.slideMaster);
    layoutIds.push({ id: nextMasterId++, rId: pkg.rel(MASTER, part, REL.slideLayout) });
    layouts.set(def.key, { part, def });
  });
  const logo = deck.logo ? { rId: pkg.rel(MASTER, media.add(deck.logo.path), REL.image), box: deck.logo.box } : null;
  pkg.addXml(MASTER, CT.slideMaster, masterXml({ geometry: deck.geometry, ctx, layoutIds, logo }));

  const slideRefs = [];
  for (const [i, spec] of slides.entries()) {
    const part = `/ppt/slides/slide${i + 1}.xml`;
    const layout = layouts.get(spec.layout);
    if (!layout) throw new Error(`Slide ${i + 1}: unknown layout ${JSON.stringify(spec.layout)}`);
    pkg.rel(part, layout.part, REL.slideLayout);
    const sctx = { ...ctx, link: (url) => pkg.rel(part, url, REL.hyperlink, { external: true }) };
    const env = { pkg, part, media, ctx: sctx, counters };

    let nextId = 2;
    let xml = "";
    // Only placeholders that have content are put on the slide; empty ones would show
    // "Click to add text" prompts in the editor.
    const ph = spec.placeholders || {};
    for (const p of layout.def.placeholders) {
      const paras = ph[p.role];
      if (!paras || !paras.length) continue;
      const box = p.role === "title" ? ph.titleBox : p.role === "body" ? ph.bodyBox : undefined;
      const body = p.role === "body" && ph.bodyFontScale ? { fontScale: ph.bodyFontScale } : {};
      xml += placeholderXml({ id: nextId++, name: p.name, ph: p.ph, box, paras, ctx: sctx, body });
    }
    for (const shape of spec.shapes || []) xml += await shapeToXml(shape, nextId++, env);

    pkg.addXml(part, CT.slide, slideXml(xml));
    slideRefs.push({ id: 256 + i, rId: pkg.rel(PRES, part, REL.slide) });
  }

  pkg.rel(PRES, THEME, REL.theme);
  for (const [name, type, rel, xml] of [
    ["/ppt/presProps.xml", CT.presProps, REL.presProps, presPropsXml()],
    ["/ppt/viewProps.xml", CT.viewProps, REL.viewProps, viewPropsXml()],
    ["/ppt/tableStyles.xml", CT.tableStyles, REL.tableStyles, EMPTY_TABLE_STYLES],
  ]) {
    pkg.addXml(name, type, xml);
    pkg.rel(PRES, name, rel);
  }
  pkg.addXml(PRES, CT.presentation, presentationXml({
    masterId, masterRId, notesMasterRId: null, slides: slideRefs,
    width: deck.geometry.width, height: deck.geometry.height,
  }));
  pkg.addXml("/docProps/core.xml", CT.coreProps, coreXml({ title: deck.title, author: deck.author, date: deck.timestamps ? new Date() : null }));
  pkg.addXml("/docProps/app.xml", CT.extProps, appXml({ slides: slides.length, notes: 0 }));
  return pkg.toBuffer();
}
```

- [ ] **Step 7: Run tests to verify they pass**

Run: `npm test`
Expected: PASS. If `layoutNameOf` or a regex fails, open the generated XML (`node -e` + `unzip -p`) and fix the writer, not the test.

- [ ] **Step 8: Commit**

```bash
git add src/ooxml/presentation.mjs src/ooxml/index.mjs test/deck.mjs test/writer.test.mjs test/fixtures/kitchen-sink/public/pic.png
git commit -m "Add writePptx with master, layouts, placeholders, text shapes and images

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 9: Tables with deck table styles

**Files:**
- Create: `src/ooxml/table.mjs`
- Modify: `src/ooxml/index.mjs` (add `table` case; use `tableStylesXml()`)
- Modify: `src/ooxml/presentation.mjs` (delete `EMPTY_TABLE_STYLES`)
- Test: `test/table.test.mjs`

**Interfaces:**
- Consumes: `txBodyXml` (Task 5), `graphicFrameXml` (Task 6), `guid`, `emuSize`, `XML_DECL`, `NS` (Task 1).
- Produces:
  - `TableShape = { kind: "table", x, y, w, h, rows: Para[][][] /* row → cell → paragraphs */, colW?: number[], header: boolean, variant?: "default"|"deflist" }`
  - `TABLE_STYLES = { default: guid("deck-table-default"), deflist: guid("deck-table-deflist") }`, `TABLE_URI`, `tableStylesXml(): string`, `tableXml(shape, ctx): string`.

- [ ] **Step 1: Write the failing test** — `test/table.test.mjs`

```js
import { test } from "node:test";
import assert from "node:assert/strict";
import { tableStylesXml, tableXml, TABLE_STYLES } from "../src/ooxml/table.mjs";
import { writePptx } from "../src/ooxml/index.mjs";
import { DECK } from "./deck.mjs";
import { lintPackage } from "./lint.mjs";
import { parseXml, unzip, byLocal } from "./helpers.mjs";

const ctx = { lang: "en-US", fonts: { mono: "M" } };
const cell = (text) => [{ runs: [{ text, size: 13 }] }];
const TABLE = { kind: "table", x: 1, y: 2, w: 6, h: 1.2, header: true, rows: [[cell("A"), cell("B"), cell("C")], [cell("1"), cell("2")], [cell("x"), cell("y"), cell("z")]] };

test("table styles part defines both deck styles with theme colors only", () => {
  const xml = tableStylesXml();
  const doc = parseXml(xml);
  assert.equal(doc.documentElement.getAttribute("def"), TABLE_STYLES.default);
  const ids = byLocal(doc, "tblStyle").map((s) => s.getAttribute("styleId"));
  assert.deepEqual(ids, [TABLE_STYLES.default, TABLE_STYLES.deflist]);
  assert.doesNotMatch(xml, /srgbClr/);
  assert.match(xml, /<a:firstRow><a:tcTxStyle b="on">/);
});

test("table: style id + header flag, padded rows, grid widths sum to frame width", () => {
  const xml = tableXml(TABLE, ctx);
  assert.match(xml, new RegExp(`<a:tblPr firstRow="1" bandRow="0"><a:tableStyleId>${TABLE_STYLES.default.replace(/[{}]/g, "\\$&")}</a:tableStyleId></a:tblPr>`));
  const doc = parseXml(`<x xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${xml}</x>`);
  const rows = byLocal(doc, "tr");
  assert.deepEqual(rows.map((r) => byLocal(r, "tc").length), [3, 3, 3]);
  const widths = byLocal(doc, "gridCol").map((c) => Number(c.getAttribute("w")));
  assert.ok(Math.abs(widths.reduce((a, b) => a + b, 0) - 6 * 914400) <= 3);
  assert.match(xml, /<a:tcPr marL="91440" marR="91440" marT="45720" marB="45720"\/>/);
  assert.doesNotMatch(xml, /<a:tcPr[^>]*>.*<a:solidFill/);
});

test("deflist variant: own style, no header, explicit column widths, centered cells", () => {
  const xml = tableXml({ ...TABLE, header: false, variant: "deflist", colW: [2.4, 3.6], rows: [[cell("t"), cell("d")]] }, ctx);
  assert.match(xml, /<a:tblPr bandRow="0"><a:tableStyleId>/);
  assert.ok(xml.includes(TABLE_STYLES.deflist));
  assert.match(xml, /<a:gridCol w="2194560"\/><a:gridCol w="3291840"\/>/);
  assert.match(xml, /anchor="ctr"/);
});

test("writer embeds tables in graphic frames and passes lint", async () => {
  const buffer = await writePptx([{ layout: "blank", shapes: [TABLE] }], DECK);
  assert.deepEqual(await lintPackage(buffer), []);
  const z = await unzip(buffer);
  assert.match(z.text("ppt/slides/slide1.xml"), /<a:graphicData uri="http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/table"><a:tbl>/);
  assert.ok(z.text("ppt/tableStyles.xml").includes(TABLE_STYLES.default));
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test`
Expected: FAIL — `Cannot find module '.../src/ooxml/table.mjs'`

- [ ] **Step 3: Implement** — `src/ooxml/table.mjs`

```js
// Native tables. Styling comes from deck table styles defined in tableStyles.xml with theme
// colors only, and cells carry no direct fills or borders - so PowerPoint's Table Design tab
// and theme recoloring both work. (If Keynote/Google Slides turn out to ignore custom table
// styles, the fallback is to add direct cell borders/fills here.)

import { XML_DECL, NS, guid, emuSize } from "./xml.mjs";
import { txBodyXml } from "./text.mjs";

export const TABLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/table";

export const TABLE_STYLES = {
  default: guid("deck-table-default"),
  deflist: guid("deck-table-deflist"),
};

function tcBdr(width) {
  const ln = `<a:ln w="${width}" cmpd="sng"><a:solidFill><a:schemeClr val="accent4"/></a:solidFill></a:ln>`;
  return `<a:tcBdr>${["left", "right", "top", "bottom", "insideH", "insideV"].map((s) => `<a:${s}>${ln}</a:${s}>`).join("")}</a:tcBdr>`;
}

function tcTxStyle(color, bold = false) {
  return `<a:tcTxStyle${bold ? ' b="on"' : ""}><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="${color}"/></a:tcTxStyle>`;
}

export function tableStylesXml() {
  // Header row: panel (bg2) fill with bold accentDark (tx2) text; borders: border color (accent4).
  const deck =
    `<a:tblStyle styleId="${TABLE_STYLES.default}" styleName="Deck Table">` +
    `<a:wholeTbl>${tcTxStyle("tx1")}<a:tcStyle>${tcBdr(12700)}<a:fill><a:noFill/></a:fill></a:tcStyle></a:wholeTbl>` +
    `<a:firstRow>${tcTxStyle("tx2", true)}<a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="bg2"/></a:solidFill></a:fill></a:tcStyle></a:firstRow>` +
    "</a:tblStyle>";
  const deflist =
    `<a:tblStyle styleId="${TABLE_STYLES.deflist}" styleName="Deck Definition List">` +
    `<a:wholeTbl>${tcTxStyle("tx1")}<a:tcStyle>${tcBdr(6350)}<a:fill><a:solidFill><a:schemeClr val="bg2"/></a:solidFill></a:fill></a:tcStyle></a:wholeTbl>` +
    "</a:tblStyle>";
  // def = the deck style, so tables users insert later in PowerPoint match the deck too.
  return `${XML_DECL}<a:tblStyleLst xmlns:a="${NS.a}" def="${TABLE_STYLES.default}">${deck}${deflist}</a:tblStyleLst>`;
}

export function tableXml(shape, ctx) {
  const cols = Math.max(1, ...shape.rows.map((r) => r.length));
  const widths = shape.colW && shape.colW.length === cols ? shape.colW : Array(cols).fill(shape.w / cols);
  const rowH = emuSize(shape.h / Math.max(1, shape.rows.length));
  const deflist = shape.variant === "deflist";
  const mar = deflist ? [76200, 76200, 50800, 50800] : [91440, 91440, 45720, 45720];
  const tcPr = `<a:tcPr marL="${mar[0]}" marR="${mar[1]}" marT="${mar[2]}" marB="${mar[3]}"${deflist ? ' anchor="ctr"' : ""}/>`;
  const rows = shape.rows
    .map((row) => {
      // Short rows are padded to the grid width - a row with fewer a:tc than gridCols is a
      // repair prompt.
      const cells = Array.from({ length: cols }, (_, i) => row[i] || []);
      // Cell bodies are written like placeholder bodies: a bare <a:bodyPr/>, no wrap/inset attrs.
      return `<a:tr h="${rowH}">${cells.map((paras) => `<a:tc>${txBodyXml(paras, ctx, { placeholder: true }, "a:txBody")}${tcPr}</a:tc>`).join("")}</a:tr>`;
    })
    .join("");
  const styleId = deflist ? TABLE_STYLES.deflist : TABLE_STYLES.default;
  const firstRow = shape.header && !deflist ? ' firstRow="1"' : "";
  const grid = widths.map((w) => `<a:gridCol w="${emuSize(w)}"/>`).join("");
  return `<a:tbl><a:tblPr${firstRow} bandRow="0"><a:tableStyleId>${styleId}</a:tableStyleId></a:tblPr><a:tblGrid>${grid}</a:tblGrid>${rows}</a:tbl>`;
}
```

- [ ] **Step 4: Wire into the writer** — `src/ooxml/index.mjs`

Add imports:

```js
import { graphicFrameXml } from "./shapes.mjs";
import { TABLE_URI, tableXml, tableStylesXml } from "./table.mjs";
```

(Merge `graphicFrameXml` into the existing `./shapes.mjs` import line rather than adding a second import of the same module.)

Add this case to `shapeToXml`, before `default:`:

```js
    case "table":
      return graphicFrameXml({ id, name: `Table ${id}`, box: shape, uri: TABLE_URI, inner: tableXml(shape, ctx) });
```

Replace `EMPTY_TABLE_STYLES` with `tableStylesXml()` in the presentation-level parts loop, remove `EMPTY_TABLE_STYLES` from the `./presentation.mjs` import, and delete the `EMPTY_TABLE_STYLES` export (and its comment) from `src/ooxml/presentation.mjs`.

- [ ] **Step 5: Run tests to verify they pass**

Run: `npm test`
Expected: PASS

- [ ] **Step 6: Commit**

```bash
git add src/ooxml/table.mjs src/ooxml/index.mjs src/ooxml/presentation.mjs test/table.test.mjs
git commit -m "Add native tables styled by theme-colored deck table styles

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 10: Charts with an embedded workbook

**Files:**
- Create: `src/ooxml/xlsx.mjs`
- Create: `src/ooxml/chart.mjs`
- Modify: `src/ooxml/index.mjs` (add `chart` case)
- Test: `test/chart.test.mjs`

**Interfaces:**
- Consumes: `Package`, `REL`, `CT` (Task 2); `esc`, `pt100`, `solidFill`, `NS`, `XML_DECL` (Task 1); `graphicFrameXml` (Task 6).
- Produces:
  - `ChartSpec = { type: "bar"|"column", categories: string[], series: [{ name: string, color: Color, values: number[] }], min?: number, max?: number, valueAxis: boolean, legend: boolean, gapWidth?: number, overlap?: number, dataLabels?: { color: Color, size: number }, background: Color, catLabelColor: Color, catLabelSize: number, valLabelColor?: Color, valLabelSize?: number, legendColor?: Color, legendSize?: number }`
  - `ChartShape = { kind: "chart", x, y, w, h, chart: ChartSpec }`
  - `colName(i): string` (0 → "A"), `async chartWorkbook({ categories, series }): Promise<Buffer>`.
  - `CHART_URI`, `chartXml(spec, { lang, workbookRId }): string`, `chartRefXml(rId): string`.

- [ ] **Step 1: Write the failing test** — `test/chart.test.mjs`

```js
import { test } from "node:test";
import assert from "node:assert/strict";
import { colName, chartWorkbook } from "../src/ooxml/xlsx.mjs";
import { chartXml } from "../src/ooxml/chart.mjs";
import { writePptx } from "../src/ooxml/index.mjs";
import { DECK } from "./deck.mjs";
import { lintPackage } from "./lint.mjs";
import { parseXml, unzip, byLocal } from "./helpers.mjs";

const SPEC = {
  type: "column", categories: ["Q1", "Q2", "Q3"],
  series: [{ name: "Before", color: "accent5", values: [22, 31, 28] }, { name: "After & now", color: "#123456", values: [41, NaN, 63] }],
  min: 0, max: 100, valueAxis: true, legend: true, gapWidth: 45, overlap: -6,
  dataLabels: { color: "dk1", size: 10 }, background: "lt1", catLabelColor: "dk1", catLabelSize: 11, valLabelColor: "accent3", valLabelSize: 9,
};

test("colName maps indexes to spreadsheet columns", () => {
  assert.deepEqual([0, 1, 25, 26, 27].map(colName), ["A", "B", "Z", "AA", "AB"]);
});

test("workbook holds categories, series names and finite values", async () => {
  const buffer = await chartWorkbook(SPEC);
  assert.deepEqual(await lintPackage(buffer), []);
  const z = await unzip(buffer);
  const sheet = z.text("xl/worksheets/sheet1.xml");
  assert.match(sheet, /<dimension ref="A1:C4"\/>/);
  assert.match(sheet, /<c r="B2"><v>22<\/v><\/c>/);
  assert.doesNotMatch(sheet, /r="C3"/); // NaN is left empty
  assert.match(z.text("xl/sharedStrings.xml"), /<t xml:space="preserve">After &amp; now<\/t>/);
  assert.match(z.text("xl/workbook.xml"), /<sheet name="Sheet1" sheetId="1" r:id="rId1"\/>/);
  assert.match(z.text("xl/_rels/workbook.xml.rels"), /Id="rId1" Type="[^"]+\/worksheet" Target="worksheets\/sheet1.xml"/);
});

test("chart XML: two axes, caches, formulas, colors, hidden value axis option", () => {
  const xml = chartXml(SPEC, { lang: "en-US", workbookRId: "rId1" });
  const doc = parseXml(xml);
  assert.equal(byLocal(byLocal(doc, "barChart")[0], "axId").length, 2);
  assert.match(xml, /<c:barDir val="col"\/><c:grouping val="clustered"\/><c:varyColors val="0"\/>/);
  assert.match(xml, /<c:f>Sheet1!\$B\$1<\/c:f><c:strCache><c:ptCount val="1"\/><c:pt idx="0"><c:v>Before<\/c:v>/);
  assert.match(xml, /<c:f>Sheet1!\$A\$2:\$A\$4<\/c:f>/);
  assert.match(xml, /<c:f>Sheet1!\$C\$2:\$C\$4<\/c:f><c:numCache><c:formatCode>General<\/c:formatCode><c:ptCount val="3"\/><c:pt idx="0"><c:v>41<\/c:v><\/c:pt><c:pt idx="2">/);
  assert.match(xml, /<c:spPr><a:solidFill><a:schemeClr val="accent5"\/><\/a:solidFill>/);
  assert.match(xml, /<a:srgbClr val="123456"\/>/);
  assert.match(xml, /<c:scaling><c:orientation val="minMax"\/><c:max val="100"\/><c:min val="0"\/><\/c:scaling><c:delete val="0"\/>/);
  assert.match(xml, /<c:gapWidth val="45"\/><c:overlap val="-6"\/><c:axId val="50001"\/><c:axId val="50002"\/><\/c:barChart>/);
  assert.match(xml, /<c:legend><c:legendPos val="b"\/>/);
  assert.match(xml, /<c:externalData r:id="rId1"><c:autoUpdate val="0"\/><\/c:externalData><\/c:chartSpace>$/);

  const bar = chartXml({ ...SPEC, type: "bar", valueAxis: false, legend: false }, { lang: "en-US", workbookRId: "rId1" });
  assert.match(bar, /<c:barDir val="bar"\/>/);
  assert.match(bar, /<c:valAx><c:axId val="50002"\/><c:scaling>.*<\/c:scaling><c:delete val="1"\/><c:axPos val="b"\/>/);
  assert.doesNotMatch(bar, /<c:legend>/);
});

test("writer adds chart part + embedded workbook and passes lint", async () => {
  const shape = { kind: "chart", x: 1, y: 2, w: 6, h: 3, chart: SPEC };
  const buffer = await writePptx([{ layout: "blank", shapes: [shape, { ...shape, x: 7 }] }], DECK);
  assert.deepEqual(await lintPackage(buffer), []);
  const z = await unzip(buffer);
  assert.ok(z.names.includes("ppt/charts/chart1.xml") && z.names.includes("ppt/charts/chart2.xml"));
  assert.match(z.text("ppt/charts/_rels/chart1.xml.rels"), /\/package" Target="..\/embeddings\/Microsoft_Excel_Worksheet1.xlsx"/);
  assert.match(z.text("[Content_Types].xml"), /<Default Extension="xlsx" ContentType="application\/vnd.openxmlformats-officedocument.spreadsheetml.sheet"\/>/);
  assert.match(z.text("ppt/slides/slide1.xml"), /<a:graphicData uri="http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/chart"><c:chart xmlns:c="[^"]+" r:id="rId2"\/>/);
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test`
Expected: FAIL — `Cannot find module '.../src/ooxml/xlsx.mjs'`

- [ ] **Step 3: Implement** — `src/ooxml/xlsx.mjs`

```js
// Minimal embedded workbook for one chart, so PowerPoint's "Edit Data" opens the numbers in
// Excel. Column A holds categories, row 1 series names, B2.. the values - matching the c:f
// formulas chart.mjs writes. Built with the same Package class, so it's lint-checked too.

import { Package, REL, CT } from "./package.mjs";
import { XML_DECL, NS, esc } from "./xml.mjs";

export function colName(i) {
  let n = i + 1;
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

const STYLES =
  `${XML_DECL}<styleSheet xmlns="${NS.ssml}">` +
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

export async function chartWorkbook({ categories, series }) {
  const strings = [];
  const index = new Map();
  const str = (text) => {
    const key = String(text);
    if (!index.has(key)) {
      index.set(key, strings.length);
      strings.push(key);
    }
    return index.get(key);
  };

  const rows = [`<row r="1">${series.map((s, j) => `<c r="${colName(j + 1)}1" t="s"><v>${str(s.name)}</v></c>`).join("")}</row>`];
  categories.forEach((cat, i) => {
    const r = i + 2;
    const cells = [`<c r="A${r}" t="s"><v>${str(cat)}</v></c>`];
    series.forEach((s, j) => {
      const v = s.values[i];
      if (Number.isFinite(v)) cells.push(`<c r="${colName(j + 1)}${r}"><v>${v}</v></c>`);
    });
    rows.push(`<row r="${r}">${cells.join("")}</row>`);
  });

  const sheet =
    `${XML_DECL}<worksheet xmlns="${NS.ssml}" xmlns:r="${NS.r}"><dimension ref="A1:${colName(series.length)}${categories.length + 1}"/>` +
    `<sheetData>${rows.join("")}</sheetData></worksheet>`;
  const sst =
    `${XML_DECL}<sst xmlns="${NS.ssml}" count="${strings.length}" uniqueCount="${strings.length}">` +
    `${strings.map((t) => `<si><t xml:space="preserve">${esc(t)}</t></si>`).join("")}</sst>`;
  const workbook =
    `${XML_DECL}<workbook xmlns="${NS.ssml}" xmlns:r="${NS.r}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`;

  const pkg = new Package();
  const WB = "/xl/workbook.xml";
  pkg.rel("", WB, REL.officeDocument);
  pkg.rel(WB, "/xl/worksheets/sheet1.xml", REL.worksheet); // rId1, referenced by <sheet r:id>
  pkg.rel(WB, "/xl/styles.xml", REL.styles);
  pkg.rel(WB, "/xl/sharedStrings.xml", REL.sharedStrings);
  pkg.addXml(WB, CT.workbook, workbook);
  pkg.addXml("/xl/worksheets/sheet1.xml", CT.worksheet, sheet);
  pkg.addXml("/xl/styles.xml", CT.styles, STYLES);
  pkg.addXml("/xl/sharedStrings.xml", CT.sharedStrings, sst);
  return pkg.toBuffer();
}
```

- [ ] **Step 4: Implement** — `src/ooxml/chart.mjs`

```js
// Clustered bar/column charts (c:chartSpace). Every series carries strCache/numCache - that's
// what Google Slides and Keynote render from - plus c:f formulas into the embedded workbook
// (xlsx.mjs) so "Edit Data" works in PowerPoint. Exactly two axes per bar chart (pptxgenjs
// writes three, which is schema-invalid).

import { XML_DECL, NS, esc, pt100, solidFill } from "./xml.mjs";
import { colName } from "./xlsx.mjs";

export const CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";

const CAT_AX = 50001;
const VAL_AX = 50002;
const NO_LINE = "<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>";

export function chartRefXml(rId) {
  return `<c:chart xmlns:c="${NS.c}" r:id="${rId}"/>`;
}

function txPr(size, color, lang) {
  return (
    `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${pt100(size)}" b="0">${solidFill(color)}` +
    `<a:latin typeface="+mn-lt"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="${esc(lang)}"/></a:p></c:txPr>`
  );
}

export function chartXml(spec, { lang, workbookRId }) {
  const horizontal = spec.type === "bar";
  const n = spec.categories.length;
  const catCache =
    `<c:strCache><c:ptCount val="${n}"/>` +
    `${spec.categories.map((c, i) => `<c:pt idx="${i}"><c:v>${esc(c)}</c:v></c:pt>`).join("")}</c:strCache>`;
  const dLbls = spec.dataLabels
    ? `<c:dLbls>${NO_LINE}${txPr(spec.dataLabels.size, spec.dataLabels.color, lang)}<c:dLblPos val="outEnd"/>` +
      '<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/>' +
      '<c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'
    : "";

  // CT_BarSer order: idx, order, tx, spPr, invertIfNegative, dLbls, cat, val.
  const series = spec.series
    .map((s, j) => {
      const col = colName(j + 1);
      const pts = s.values.map((v, i) => (Number.isFinite(v) ? `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>` : "")).join("");
      return (
        `<c:ser><c:idx val="${j}"/><c:order val="${j}"/>` +
        `<c:tx><c:strRef><c:f>Sheet1!$${col}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${esc(s.name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
        `<c:spPr>${solidFill(s.color)}<a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/>${dLbls}` +
        `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$${n + 1}</c:f>${catCache}</c:strRef></c:cat>` +
        `<c:val><c:numRef><c:f>Sheet1!$${col}$2:$${col}$${n + 1}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${n}"/>${pts}</c:numCache></c:numRef></c:val>` +
        "</c:ser>"
      );
    })
    .join("");

  const scaling =
    '<c:scaling><c:orientation val="minMax"/>' +
    `${spec.max != null ? `<c:max val="${spec.max}"/>` : ""}${spec.min != null ? `<c:min val="${spec.min}"/>` : ""}</c:scaling>`;
  const ticks = '<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>';
  const catAx =
    `<c:catAx><c:axId val="${CAT_AX}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>` +
    `<c:axPos val="${horizontal ? "l" : "b"}"/>${ticks}${NO_LINE}${txPr(spec.catLabelSize, spec.catLabelColor, lang)}` +
    `<c:crossAx val="${VAL_AX}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
  // No c:majorGridlines element = no gridlines.
  const valAx =
    `<c:valAx><c:axId val="${VAL_AX}"/>${scaling}<c:delete val="${spec.valueAxis ? 0 : 1}"/>` +
    `<c:axPos val="${horizontal ? "b" : "l"}"/>${ticks}${NO_LINE}${txPr(spec.valLabelSize ?? 9, spec.valLabelColor ?? "accent3", lang)}` +
    `<c:crossAx val="${CAT_AX}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  const legend = spec.legend
    ? `<c:legend><c:legendPos val="b"/><c:overlay val="0"/>${txPr(spec.legendSize ?? 11, spec.legendColor ?? "dk1", lang)}</c:legend>`
    : "";
  const background = `<c:spPr>${solidFill(spec.background)}<a:ln><a:noFill/></a:ln></c:spPr>`;
  const overlap = spec.overlap != null ? `<c:overlap val="${spec.overlap}"/>` : "";

  return (
    `${XML_DECL}<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">` +
    '<c:date1904 val="0"/><c:roundedCorners val="0"/>' +
    '<c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>' +
    `<c:barChart><c:barDir val="${horizontal ? "bar" : "col"}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series}` +
    `<c:gapWidth val="${spec.gapWidth ?? 100}"/>${overlap}<c:axId val="${CAT_AX}"/><c:axId val="${VAL_AX}"/></c:barChart>` +
    `${catAx}${valAx}${background}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `${background}<c:externalData r:id="${workbookRId}"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`
  );
}
```

- [ ] **Step 5: Wire into the writer** — `src/ooxml/index.mjs`

Add imports:

```js
import { CHART_URI, chartXml, chartRefXml } from "./chart.mjs";
import { chartWorkbook } from "./xlsx.mjs";
```

Add this case to `shapeToXml`, before `default:`:

```js
    case "chart": {
      const n = ++env.counters.chart;
      const chartPart = `/ppt/charts/chart${n}.xml`;
      const workbookPart = `/ppt/embeddings/Microsoft_Excel_Worksheet${n}.xlsx`;
      pkg.addMedia(workbookPart, await chartWorkbook(shape.chart));
      const workbookRId = pkg.rel(chartPart, workbookPart, REL.package);
      pkg.addXml(chartPart, CT.chart, chartXml(shape.chart, { lang: ctx.lang, workbookRId }));
      const rId = pkg.rel(part, chartPart, REL.chart);
      return graphicFrameXml({ id, name: `Chart ${id}`, box: shape, uri: CHART_URI, inner: chartRefXml(rId) });
    }
```

- [ ] **Step 6: Run tests to verify they pass**

Run: `npm test`
Expected: PASS

- [ ] **Step 7: Commit**

```bash
git add src/ooxml/xlsx.mjs src/ooxml/chart.mjs src/ooxml/index.mjs test/chart.test.mjs
git commit -m "Add bar/column charts with caches and an embedded workbook for Edit Data

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 11: Speaker notes

**Files:**
- Create: `src/ooxml/notes.mjs`
- Modify: `src/ooxml/index.mjs` (notes master + notes slides; `notesMasterRId`; app.xml notes count)
- Test: `test/notes.test.mjs`

**Interfaces:**
- Consumes: `txBodyXml` (Task 5), `SP_TREE_HEADER`, `placeholderXml` (Task 6), `themeXml` (Task 4).
- Produces: `notesParas(text): Para[]` (one paragraph per non-empty line; runs of blank lines collapse to one empty paragraph; leading/trailing blank lines dropped), `notesMasterXml(ctx): string`, `notesSlideXml(text, ctx): string`. `SlideSpec.notes` (string) now written.

- [ ] **Step 1: Write the failing test** — `test/notes.test.mjs`

```js
import { test } from "node:test";
import assert from "node:assert/strict";
import { notesParas } from "../src/ooxml/notes.mjs";
import { writePptx } from "../src/ooxml/index.mjs";
import { DECK } from "./deck.mjs";
import { lintPackage } from "./lint.mjs";
import { unzip } from "./helpers.mjs";

test("notesParas: one paragraph per line, blank runs collapse", () => {
  assert.deepEqual(notesParas("\n  first  \nsecond\n\n\n- third\n\n"), [
    { runs: [{ text: "first" }] },
    { runs: [{ text: "second" }] },
    { runs: [] },
    { runs: [{ text: "- third" }] },
  ]);
});

test("slides with notes get notes slides wired to the notes master", async () => {
  const slides = [
    { layout: "blank", shapes: [], notes: "Say hello <warmly> & smile" },
    { layout: "blank", shapes: [] },
  ];
  const buffer = await writePptx(slides, DECK);
  assert.deepEqual(await lintPackage(buffer), []);
  const z = await unzip(buffer);
  assert.ok(z.names.includes("ppt/notesSlides/notesSlide1.xml"));
  assert.ok(!z.names.some((n) => n.startsWith("ppt/notesSlides/notesSlide2")));
  assert.match(z.text("ppt/notesSlides/notesSlide1.xml"), /<p:ph type="body" idx="1"\/>.*<a:t>Say hello &lt;warmly&gt; &amp; smile<\/a:t>/);
  assert.match(z.text("ppt/notesSlides/_rels/notesSlide1.xml.rels"), /\/notesMaster" Target="..\/notesMasters\/notesMaster1.xml"/);
  assert.match(z.text("ppt/notesSlides/_rels/notesSlide1.xml.rels"), /\/slide" Target="..\/slides\/slide1.xml"/);
  assert.match(z.text("ppt/slides/_rels/slide1.xml.rels"), /\/notesSlide" Target="..\/notesSlides\/notesSlide1.xml"/);
  assert.match(z.text("ppt/notesMasters/_rels/notesMaster1.xml.rels"), /\/theme" Target="..\/theme\/theme2.xml"/);
  assert.match(z.text("ppt/presentation.xml"), /<\/p:sldMasterIdLst><p:notesMasterIdLst><p:notesMasterId r:id="rId\d+"\/><\/p:notesMasterIdLst><p:sldIdLst>/);
  assert.match(z.text("docProps/app.xml"), /<Notes>1<\/Notes>/);
});

test("decks without notes have no notes master", async () => {
  const z = await unzip(await writePptx([{ layout: "blank", shapes: [] }], DECK));
  assert.ok(!z.names.some((n) => n.includes("notesMaster")));
  assert.doesNotMatch(z.text("ppt/presentation.xml"), /notesMasterIdLst/);
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test`
Expected: FAIL — `Cannot find module '.../src/ooxml/notes.mjs'`

- [ ] **Step 3: Implement** — `src/ooxml/notes.mjs`

```js
// Speaker notes: a notes master (written only when some slide has notes) and one notes slide
// per slide that has them. Positions are PowerPoint's own defaults for a 7.5x10in notes page.

import { XML_DECL, NS_PML } from "./xml.mjs";
import { SP_TREE_HEADER, placeholderXml, xfrmXml } from "./shapes.mjs";

export function notesParas(text) {
  const paras = [];
  for (const raw of String(text ?? "").split("\n")) {
    const line = raw.trim();
    if (line === "") {
      if (paras.length && paras[paras.length - 1].runs.length) paras.push({ runs: [] });
      continue;
    }
    paras.push({ runs: [{ text: line }] });
  }
  while (paras.length && !paras[paras.length - 1].runs.length) paras.pop();
  return paras;
}

const CLR_MAP =
  '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" ' +
  'accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';

const SLIDE_IMAGE_LOCKS = '<a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/>';

export function notesMasterXml(ctx) {
  const image =
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr>${SLIDE_IMAGE_LOCKS}</p:cNvSpPr>` +
    '<p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr>' +
    `<p:spPr>${xfrmXml({ x: 0.75, y: 1.25, w: 6, h: 3.375 })}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>` +
    '<a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr></p:sp>';
  const body = placeholderXml({
    id: 3, name: "Notes Placeholder 2", ph: { type: "body", idx: 3, sz: "quarter" }, ctx,
    box: { x: 0.75, y: 4.8125, w: 6, h: 3.9375 },
    paras: [{ runs: [{ text: "Click to edit Master text styles" }] }],
  });
  const style =
    '<p:notesStyle><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">' +
    '<a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>' +
    '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:notesStyle>';
  return (
    `${XML_DECL}<p:notesMaster ${NS_PML}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>` +
    `<p:spTree>${SP_TREE_HEADER}${image}${body}</p:spTree></p:cSld>${CLR_MAP}${style}</p:notesMaster>`
  );
}

export function notesSlideXml(text, ctx) {
  const image =
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr>${SLIDE_IMAGE_LOCKS}</p:cNvSpPr>` +
    '<p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>';
  const body = placeholderXml({ id: 3, name: "Notes Placeholder 2", ph: { type: "body", idx: 1 }, ctx, paras: notesParas(text) });
  return (
    `${XML_DECL}<p:notes ${NS_PML}><p:cSld><p:spTree>${SP_TREE_HEADER}${image}${body}</p:spTree></p:cSld>` +
    "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>"
  );
}
```

- [ ] **Step 4: Wire into the writer** — `src/ooxml/index.mjs`

Add import:

```js
import { notesMasterXml, notesSlideXml } from "./notes.mjs";
```

Add a constant next to the others:

```js
const NOTES_MASTER = "/ppt/notesMasters/notesMaster1.xml";
```

Before the slide loop, add:

```js
  // The notes master (with its own theme - every master needs one) only exists if some slide
  // has notes.
  const hasNotes = slides.some((s) => s.notes);
  let notesMasterRId = null;
  if (hasNotes) {
    const NOTES_THEME = "/ppt/theme/theme2.xml";
    pkg.addXml(NOTES_THEME, CT.theme, themeXml({ name: "Notes", colors: deck.theme, fonts: deck.fonts }));
    pkg.addXml(NOTES_MASTER, CT.notesMaster, notesMasterXml(ctx));
    pkg.rel(NOTES_MASTER, NOTES_THEME, REL.theme);
    notesMasterRId = pkg.rel(PRES, NOTES_MASTER, REL.notesMaster);
  }
  let notesCount = 0;
```

Inside the slide loop, right after `pkg.addXml(part, CT.slide, slideXml(xml));`, add:

```js
    if (spec.notes) {
      const notesPart = `/ppt/notesSlides/notesSlide${i + 1}.xml`;
      pkg.addXml(notesPart, CT.notesSlide, notesSlideXml(spec.notes, ctx));
      pkg.rel(notesPart, NOTES_MASTER, REL.notesMaster);
      pkg.rel(notesPart, part, REL.slide);
      pkg.rel(part, notesPart, REL.notesSlide);
      notesCount++;
    }
```

In the `presentationXml({...})` call replace `notesMasterRId: null` with `notesMasterRId`, and in the `appXml` call replace `notes: 0` with `notes: notesCount`.

- [ ] **Step 5: Run tests to verify they pass**

Run: `npm test`
Expected: PASS

- [ ] **Step 6: Commit**

```bash
git add src/ooxml/notes.mjs src/ooxml/index.mjs test/notes.test.mjs
git commit -m "Write speaker notes as notes slides with a notes master

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 12: Parser — speaker notes and inline tokens

**Files:**
- Modify: `src/parse.mjs` (`mdInlineToRuns` at lines 136-164; `parseSlide` at lines 384-411)
- Test: `test/parse.test.mjs`

**Interfaces:**
- Produces:
  - `mdInlineTokens(text): Array<{ text: string, bold?: true, italic?: true, code?: true }>` — pure markdown, no styling.
  - `parseSlide(...)` now returns `{ eyebrow, title, blocks, notes: string | null }`. Notes = the slide's **last** HTML comment when nothing but whitespace follows it (Slidev's convention), trimmed. Found after code fences are tokenized, so comments inside code are never notes.
  - `mdInlineToRuns` keeps working (re-implemented on top of `mdInlineTokens`) until Task 14 deletes it with `render.mjs`.

- [ ] **Step 1: Write the failing test** — `test/parse.test.mjs`

```js
import { test } from "node:test";
import assert from "node:assert/strict";
import { parseSlide, mdInlineTokens, mdInlineToRuns } from "../src/parse.mjs";
import { mergeConfig } from "../src/config.mjs";

const config = mergeConfig({});
const parse = (md) => parseSlide(md, [], config);

test("mdInlineTokens splits bold, italic and code", () => {
  assert.deepEqual(mdInlineTokens("a **b** *c* `d` e"), [
    { text: "a " }, { text: "b", bold: true }, { text: " " }, { text: "c", italic: true },
    { text: " " }, { text: "d", code: true }, { text: " e" },
  ]);
  assert.deepEqual(mdInlineTokens(""), [{ text: "" }]);
});

test("mdInlineToRuns still produces the old pptxgenjs run shape", () => {
  const runs = mdInlineToRuns("x **y**", { fontSize: 13 }, config);
  assert.deepEqual(runs.map((r) => [r.text, r.options.bold === true]), [["x ", false], ["y", true]]);
  assert.equal(runs[1].options.color, config.colors.accentDark);
});

test("the last trailing HTML comment becomes speaker notes", () => {
  const s = parse("# Title\n\nBody text\n\n<!-- An earlier comment -->\n\nMore\n\n<!--\nLine one\n\nLine two\n-->\n");
  assert.equal(s.notes, "Line one\n\nLine two");
  assert.equal(s.title, "Title");
  assert.deepEqual(s.blocks.map((b) => b.text), ["Body text", "More"]);
});

test("a comment that is followed by content is not notes", () => {
  assert.equal(parse("# T\n\n<!-- not notes -->\n\nBody").notes, null);
});

test("a comment inside a code fence is not notes", () => {
  assert.equal(parse("# T\n\n```html\n<!-- in code -->\n```\n").notes, null);
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test`
Expected: FAIL — `mdInlineTokens` is not exported / `notes` is `undefined`.

- [ ] **Step 3: Implement** — in `src/parse.mjs`, replace the whole `mdInlineToRuns` function (lines 136-164) with:

```js
// Inline markdown -> tokens. Pure markdown only: what a token *looks* like (colors, fonts) is
// decided by layout.mjs.
export function mdInlineTokens(text) {
  const tokens = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)/g;
  let last = 0;
  let m;
  while ((m = re.exec(text))) {
    if (m.index > last) tokens.push({ text: text.slice(last, m.index) });
    const token = m[0];
    if (token.startsWith("**")) tokens.push({ text: token.slice(2, -2), bold: true });
    else if (token.startsWith("`")) tokens.push({ text: token.slice(1, -1), code: true });
    else tokens.push({ text: token.slice(1, -1), italic: true });
    last = re.lastIndex;
  }
  if (last < text.length) tokens.push({ text: text.slice(last) });
  if (tokens.length === 0) tokens.push({ text });
  return tokens;
}

// pptxgenjs-shaped runs for the legacy renderer (render.mjs); removed together with it.
function mdInlineToRuns(text, baseOpts, config) {
  const base = { fontFace: config.fonts.body, ...baseOpts };
  return mdInlineTokens(text).map((t) => {
    if (t.bold) return { text: t.text, options: { ...base, bold: true, color: config.colors.accentDark } };
    if (t.code) return { text: t.text, options: { ...base, fontFace: config.fonts.mono, color: config.colors.accentDark } };
    if (t.italic) return { text: t.text, options: { ...base, italic: true } };
    return { text: t.text, options: { ...base } };
  });
}
```

Then in `parseSlide`, insert after the fenced-code-block `text = text.replace(...)` statement and before `const blocks = flattenContent(text, config);`:

```js
  // Speaker notes: Slidev treats the slide's last HTML comment as its notes when nothing but
  // whitespace follows it. The capture can't contain "<!--", so an earlier comment never
  // swallows the real one. Runs after code fences are tokenized so comments in code are kept.
  let notes = null;
  const notesMatch = /<!--((?:(?!<!--)[\s\S])*?)-->\s*$/.exec(text);
  if (notesMatch) {
    notes = notesMatch[1].trim() || null;
    text = text.slice(0, notesMatch.index);
  }
```

and change the return to:

```js
  return { eyebrow, title, blocks, notes };
```

(The existing `export { mdInlineToRuns };` line at the bottom of the file stays until Task 14.)

- [ ] **Step 4: Run tests to verify they pass**

Run: `npm test`
Expected: PASS

- [ ] **Step 5: Check the old pipeline still works**

Run: `node bin/cli.mjs ../innotek-SlideDevBase/slides.md "${TMPDIR:-/tmp}/old-pipeline.pptx"` (skip if that deck isn't present)
Expected: `Wrote 7 slides to .../old-pipeline.pptx`

- [ ] **Step 6: Commit**

```bash
git add src/parse.mjs test/parse.test.mjs
git commit -m "Parse Slidev speaker notes and split inline markdown into style-free tokens

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 13: Layout engine (`layout.mjs`)

**Files:**
- Create: `src/layout.mjs`
- Modify: `package.json` (add `image-size` dependency)
- Test: `test/layout.test.mjs`

**Interfaces:**
- Consumes: `mdInlineTokens` (Task 12), `themeColors`, `toColor`, `hexOf`, `blendWithWhite` (Task 4), `slideGeometry` (Task 7). Produces `SlideSpec` (Task 8) with `TextShape` (Task 6), `TableShape` (Task 9), `ChartShape` (Task 10), image shapes (Task 8).
- Produces:
  - `deckStyle(config): { config, theme, geometry, color: { text, accentDark, panel, accent, accent2, muted, border }, hex(h): Color }`
  - `layoutDeck(parsedSlides, codeStore, config): SlideSpec[]` — `parsedSlides[i] = { eyebrow, title, blocks, notes, frontmatter }`.
- Rules (from the spec): first slide / `class` contains `cover` / `layout: cover|intro` → `title`; no title → `blank` (all shapes); `layout: center|end` → `centered` with `titleBox`/`bodyBox` overrides centering the stack; `layout: section` with text-only or empty body → `section`; otherwise text-only body → `content` (+ `bodyFontScale` when the estimate overflows), anything else → `titleOnly` with positioned shapes. Text-only = every block is `paragraph`, `bullets` or `heading`.

- [ ] **Step 1: Add the dependency**

Run: `npm install image-size@^2.0.2`

- [ ] **Step 2: Write the failing test** — `test/layout.test.mjs`

```js
import { test } from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";
import { join } from "node:path";
import { splitSlides, parseSlide } from "../src/parse.mjs";
import { mergeConfig } from "../src/config.mjs";
import { layoutDeck } from "../src/layout.mjs";

const FIXTURE_DIR = fileURLToPath(new URL("./fixtures/kitchen-sink/", import.meta.url));

function specsFor(md, userConfig = {}) {
  const config = mergeConfig(userConfig);
  config.inputDir = FIXTURE_DIR;
  config.publicDir = join(FIXTURE_DIR, "public");
  config.resolvedDataDir = null;
  const codeStore = [];
  const { slides } = splitSlides(md);
  const parsed = slides.map((s) => ({ ...parseSlide(s.content, codeStore, config), frontmatter: s.frontmatter }));
  return layoutDeck(parsed, codeStore, config);
}
const texts = (paras) => paras.map((p) => p.runs.map((r) => r.text).join(""));
const COVER = "# Cover\n\n";

test("first slide is a Title Slide with title and subtitle placeholders", () => {
  const [cover] = specsFor("# Deck\n\nThe **subtitle**\n");
  assert.equal(cover.layout, "title");
  assert.deepEqual(texts(cover.placeholders.title), ["Deck"]);
  assert.deepEqual(cover.placeholders.subtitle[0].runs, [
    { font: "body", text: "The " },
    { font: "body", text: "subtitle", b: true, color: "dk2" },
  ]);
});

test("text-only body goes into the content placeholder with native bullets", () => {
  const [, s] = specsFor(`${COVER}---\n\n<div class="eyebrow">Intro</div>\n\n# Plain\n\nPara with **bold**\n\n- one\n- two\n\n1. first\n2. second\n\n<!-- say this -->\n`);
  assert.equal(s.layout, "content");
  assert.deepEqual(texts(s.placeholders.title), ["Plain"]);
  assert.deepEqual(s.placeholders.body.map((p) => p.bullet), ["none", "char", "char", "number", "number"]);
  assert.equal(s.placeholders.body[0].runs[0].size, undefined); // inherits the master's 13pt
  assert.equal(s.placeholders.body[0].runs[0].color, undefined);
  assert.equal(s.placeholders.body[0].runs[1].color, "dk2");
  assert.equal(s.placeholders.body[1].spaceBefore, 12);
  assert.equal(s.placeholders.body[2].spaceBefore, undefined);
  assert.equal(s.placeholders.bodyFontScale, undefined);
  assert.equal(s.notes, "say this");
  const eyebrow = s.shapes[0];
  assert.deepEqual(eyebrow.paras[0].runs[0], { text: "INTRO", font: "heading", size: 9, b: true, color: "accent1", spacing: 1.5 });
});

test("an overflowing text-only body gets a shrink-on-overflow font scale", () => {
  const bullets = Array.from({ length: 40 }, (_, i) => `- bullet number ${i}`).join("\n");
  const [, s] = specsFor(`${COVER}---\n\n# Dense\n\n${bullets}\n`);
  assert.equal(s.layout, "content");
  assert.ok(s.placeholders.bodyFontScale >= 0.55 && s.placeholders.bodyFontScale < 1, String(s.placeholders.bodyFontScale));
});

test("a rich body uses Title Only plus positioned shapes", () => {
  const [, s] = specsFor(`${COVER}---\n\n# Table\n\n| a | b |\n| --- | --- |\n| 1 | 2 |\n`);
  assert.equal(s.layout, "titleOnly");
  const table = s.shapes.find((x) => x.kind === "table");
  assert.equal(table.header, true);
  assert.equal(table.rows.length, 2);
  assert.equal(table.y, 1.6);
});

test("layout: center with text-center centers the stack and the text", () => {
  const [, s] = specsFor(`${COVER}---\nlayout: center\nclass: text-center\n---\n\n# Thanks\n\nQuestions?\n`);
  assert.equal(s.layout, "centered");
  assert.equal(s.placeholders.title[0].align, "ctr");
  assert.equal(s.placeholders.body[0].align, "ctr");
  assert.ok(s.placeholders.titleBox.y > 2.5, String(s.placeholders.titleBox.y));
  assert.ok(s.placeholders.bodyBox.y > s.placeholders.titleBox.y);
});

test("layout: section maps to Section Header", () => {
  const [, s] = specsFor(`${COVER}---\nlayout: section\n---\n\n# Part two\n\nWhat comes next\n`);
  assert.equal(s.layout, "section");
  assert.deepEqual(texts(s.placeholders.body), ["What comes next"]);
});

test("a slide without a title is Blank with everything as shapes", () => {
  const [, s] = specsFor(`${COVER}---\n\nJust text\n`);
  assert.equal(s.layout, "blank");
  assert.equal(s.shapes[0].kind, "text");
});

test("charts: resolver colors map to theme slots, unresolved tags become placeholders", () => {
  const C = () => ({ kind: "barchart", labels: ["a", "b"], datasets: [{ name: "x", color: "2563EB", data: [1, 2] }, { name: "y", data: [3, 4] }], min: 0, max: 10 });
  const [, s] = specsFor(`${COVER}---\n\n# Charts\n\n<C />\n\n<Missing />\n`, { chartResolvers: { C } });
  const chart = s.shapes.find((x) => x.kind === "chart").chart;
  assert.equal(chart.type, "column");
  assert.deepEqual(chart.series.map((x) => x.color), ["accent1", "accent2"]);
  const missing = s.shapes.find((x) => x.kind === "text" && x.line && x.line.dash === "dash");
  assert.match(missing.paras[0].runs[0].text, /Missing — interactive chart, no data directory configured/);
});

test("images are contain-fitted with alt text; missing ones become placeholders", () => {
  const [, s] = specsFor(`${COVER}---\n\n# Pics\n\n![Alt text](/pic.png)\n\n![nope](/missing.png)\n`);
  const img = s.shapes.find((x) => x.kind === "image");
  assert.equal(img.alt, "Alt text");
  assert.ok(Math.abs(img.w / img.h - 2) < 1e-9);
  assert.ok(s.shapes.some((x) => x.kind === "text" && /image not found: \/missing.png/.test(x.paras?.[0]?.runs?.[0]?.text || "")));
});

test("opacity fades become blended hex colors", () => {
  const [, s] = specsFor(`${COVER}---\n\n# Fade\n\n<div class="opacity-50">Faded</div>\n`);
  assert.equal(s.placeholders.body[0].runs[0].color, "#8F949B");
});

test("card rows become rounded panels with their own content inside", () => {
  const [, s] = specsFor(`${COVER}---\n\n# Cards\n\n<div class="grid grid-cols-2 gap-4">\n<div>\n\n### Left\n\nText\n\n</div>\n<div>\n\n### Right\n\nMore\n\n</div>\n</div>\n`);
  assert.equal(s.layout, "titleOnly");
  assert.equal(s.shapes.filter((x) => x.geometry === "roundRect").length, 2);
});
```

- [ ] **Step 3: Run test to verify it fails**

Run: `npm test`
Expected: FAIL — `Cannot find module '.../src/layout.mjs'`

- [ ] **Step 4: Implement** — `src/layout.mjs`

```js
// Turns parsed slides (see parse.mjs) into SlideSpecs for the OOXML writer in ./ooxml/.
// Titles always go into the layout's title placeholder; text-only bodies (paragraphs, bullets,
// sub-headings) go into its body placeholder, so they behave like hand-made PowerPoint slides.
// Everything else is drawn as positioned shapes, using the geometry the pptxgenjs renderer
// used. Every color/font decision reads from `config` - nothing here is specific to one deck.

import { existsSync, readFileSync } from "node:fs";
import { join, extname } from "node:path";
import { imageSize } from "image-size";
import { mdInlineTokens } from "./parse.mjs";
import { themeColors, toColor, hexOf, blendWithWhite } from "./colors.mjs";
import { slideGeometry } from "./geometry.mjs";

// The master's level-1 body size (ooxml/master.mjs) - placeholder text inherits it.
const BODY_PT = 13;
const TEXT_ONLY = new Set(["paragraph", "bullets", "heading"]);
const IMAGE_EXTS = new Set([".png", ".jpg", ".jpeg", ".gif"]);
// Series colors for resolver datasets/panels that don't name one.
const PALETTE = ["accent1", "accent2", "accent5", "accent6", "accent3"];

export function deckStyle(config) {
  const theme = themeColors(config.colors);
  const slot = (key) => toColor(config.colors[key], theme);
  return {
    config,
    theme,
    geometry: slideGeometry(config.slide),
    color: {
      text: slot("text"), accentDark: slot("accentDark"), panel: slot("panel"), accent: slot("accent"),
      accent2: slot("accent2"), muted: slot("muted"), border: slot("border"),
    },
    hex: (h) => toColor(h, theme),
  };
}

// ---------------------------------------------------------------------------
// Height estimation (unchanged from the pptxgenjs renderer)
// ---------------------------------------------------------------------------

function estimateBlockHeight(block, width, textSize, codeStore) {
  const charsPerLine = Math.max(20, Math.floor((width * 72) / (textSize * 0.52)));
  const lineHeight = (textSize + 5) / 72;
  switch (block.type) {
    case "heading":
      return lineHeight * 1.5 + 0.08;
    case "paragraph":
      return Math.max(1, Math.ceil(block.text.length / charsPerLine)) * lineHeight + 0.1;
    case "quote":
      return Math.max(1, Math.ceil(block.text.length / charsPerLine)) * lineHeight + 0.16;
    case "bullets": {
      let lines = 0;
      for (const item of block.items) lines += Math.max(1, Math.ceil(item.length / (charsPerLine * 0.9)));
      return lines * lineHeight + 0.1;
    }
    case "code":
      return (codeStore[block.idx] || "").split("\n").length * lineHeight * 0.95 + 0.2;
    case "image":
      return 2.4;
    case "table":
      return block.rows.length * 0.32 + 0.15;
    case "deflist": {
      const perLine = Math.max(20, Math.floor((width * 0.58 * 72) / (textSize * 0.52)));
      let total = 0;
      for (const r of block.rows) total += Math.max(1, Math.ceil(r.desc.length / perLine)) * lineHeight + 0.12;
      return total + 0.05;
    }
    case "styledtext": {
      const size = block.style.size || textSize;
      const perLine = Math.max(20, Math.floor((width * 72) / (size * 0.55)));
      return Math.max(1, Math.ceil(block.text.length / perLine)) * ((size + 6) / 72) + 0.06;
    }
    case "chart": {
      const kind = block.resolved && block.resolved.kind;
      if (kind === "barchart-panels" || kind === "barchart") return 3.1;
      return 2.1; // statcards, or the unresolved placeholder
    }
    case "cardrow": {
      const n = block.cards.length;
      const cardW = (width - 0.18 * (n - 1)) / n;
      let maxH = 0.6; // floor so a one-line card isn't razor thin
      for (const card of block.cards) {
        maxH = Math.max(maxH, blocksHeight(card.blocks, cardW - 0.28, Math.max(8, textSize - 2), codeStore, 0.1) + 0.28);
      }
      return maxH;
    }
    default:
      return 0.3;
  }
}

function blocksHeight(blocks, width, textSize, codeStore, gap = 0.14) {
  let total = 0;
  for (const b of blocks) total += estimateBlockHeight(b, width, textSize, codeStore) + gap;
  return total;
}

// PowerPoint-style shrink: 2.5% steps, never below 55% (same floor as the free-shape scaling).
function shrinkScale(available, needed) {
  return Math.max(0.55, Math.floor((available / needed) * 40) / 40);
}

// ---------------------------------------------------------------------------
// Text
// ---------------------------------------------------------------------------

function inlineRuns(text, base, S) {
  return mdInlineTokens(text).map((t) => {
    if (t.bold) return { ...base, text: t.text, b: true, color: S.color.accentDark };
    if (t.code) return { ...base, text: t.text, font: "mono", color: S.color.accentDark };
    if (t.italic) return { ...base, text: t.text, i: true };
    return { ...base, text: t.text };
  });
}

// Tailwind/UnoCSS opacity-NN fades the whole element, including its bold/code runs.
function fadeRuns(runs, alpha, S) {
  return runs.map((r) => ({ ...r, color: `#${blendWithWhite(hexOf(r.color || S.color.text, S.theme), alpha)}` }));
}

// Paragraphs for text-only blocks. In a placeholder (`size` null) plain runs inherit the
// master's size and color; in a free text box they carry them explicitly.
function blockParas(blocks, size, S, align) {
  const base = size ? { font: "body", size, color: S.color.text } : { font: "body" };
  const paras = [];
  blocks.forEach((block, bi) => {
    const first = paras.length;
    const runsFor = (text) => {
      const runs = inlineRuns(text, base, S);
      return block.opacity != null ? fadeRuns(runs, block.opacity, S) : runs;
    };
    if (block.type === "paragraph") {
      paras.push({ bullet: "none", align, runs: runsFor(block.text) });
    } else if (block.type === "bullets") {
      for (const item of block.items) paras.push({ bullet: block.ordered ? "number" : "char", align, runs: runsFor(item) });
    } else if (block.type === "heading") {
      // Only the slide title uses the heading font - a ##/### inside the body is bold body text.
      const hs = (size || BODY_PT) + (block.level === 2 ? 5 : 2);
      paras.push({ bullet: "none", align, runs: [{ text: block.text, font: "body", b: true, size: hs, color: S.color.accentDark }] });
    }
    // Block-to-block gap (the old renderer's 0.16") as space before the block's first paragraph.
    if (bi > 0 && paras[first]) paras[first].spaceBefore = 12;
  });
  return paras;
}

function eyebrowShape(text, box, align, S) {
  return {
    kind: "text", ...box, geometry: "rect", anchor: "ctr",
    paras: [{ align, runs: [{ text: text.toUpperCase(), font: "heading", size: 9, b: true, color: S.color.accent, spacing: 1.5 }] }],
  };
}

function placeholderBox(box, label, textSize, S) {
  return {
    kind: "text", ...box, geometry: "rect", anchor: "ctr", fill: S.color.panel,
    line: { color: S.color.border, width: 1, dash: "dash" },
    paras: [{ align: "ctr", runs: [{ text: label, font: "body", size: textSize, i: true, color: S.color.muted }] }],
  };
}

// ---------------------------------------------------------------------------
// Images
// ---------------------------------------------------------------------------

// Slidev decks reference images either markdown-relative ("./foo.png") or public-dir-rooted
// ("/assets/foo.png", resolved against the deck's public/ folder like a Vite asset).
function resolveImageSrc(src, config) {
  if (/^https?:\/\//.test(src)) return null; // no network fetches during export
  const candidates = [];
  if (src.startsWith("/")) {
    if (config.publicDir) candidates.push(join(config.publicDir, src.slice(1)));
  } else {
    if (config.inputDir) candidates.push(join(config.inputDir, src));
    if (config.publicDir) candidates.push(join(config.publicDir, src));
  }
  return candidates.find((p) => existsSync(p)) || null;
}

// "contain" inside the box, anchored top-left like an inline Slidev image.
function fitContain(box, width, height) {
  if (!(width > 0 && height > 0)) return { ...box };
  const s = Math.min(box.w / width, box.h / height);
  return { x: box.x, y: box.y, w: width * s, h: height * s };
}

function imageShape(block, box, textSize, S) {
  const src = resolveImageSrc(block.src, S.config);
  if (!src) return placeholderBox(box, `[ image not found: ${block.src} ]`, textSize, S);
  if (!IMAGE_EXTS.has(extname(src).toLowerCase())) return placeholderBox(box, `[ unsupported image format: ${block.src} ]`, textSize, S);
  let dims;
  try {
    dims = imageSize(readFileSync(src));
  } catch {
    return placeholderBox(box, `[ unreadable image: ${block.src} ]`, textSize, S);
  }
  return { kind: "image", ...fitContain(box, dims.width, dims.height), path: src, alt: block.alt };
}

// ---------------------------------------------------------------------------
// Charts (resolver output -> native shapes/charts)
// ---------------------------------------------------------------------------

function label(text, box, size, { align, anchor = "ctr", ...run } = {}) {
  return { kind: "text", ...box, geometry: "rect", anchor, paras: [{ align, runs: [{ text: String(text), font: "body", size, ...run }] }] };
}

// statcards: a row of panel cards, each either a small grid of "big number + label" stat
// pairs, or one large centered percentage (+ subtitle).
function statCardShapes(cards, { x, y, w, h }, baseTextSize, S) {
  const { color } = S;
  const shapes = [];
  const gap = 0.2;
  const cardW = (w - gap * (cards.length - 1)) / cards.length;
  cards.forEach((card, i) => {
    const cx = x + i * (cardW + gap);
    shapes.push({ kind: "text", x: cx, y, w: cardW, h, geometry: "roundRect", radius: 0.1, fill: color.panel, shadow: true, paras: [] });
    let cy = y + 0.18;
    if (card.title) {
      const dot = 0.1;
      shapes.push({ kind: "text", x: cx + 0.18, y: cy + 0.03, w: dot, h: dot, geometry: "ellipse", fill: card.accent ? S.hex(card.accent) : color.accent, paras: [] });
      shapes.push(label(card.title, { x: cx + 0.18 + dot + 0.09, y: cy - 0.04, w: cardW - 0.36 - dot - 0.09, h: 0.3 }, Math.max(9, baseTextSize - 1), { b: true, color: color.text }));
      cy += 0.42;
    }
    if (card.big) {
      shapes.push(label(card.big.value, { x: cx, y: cy, w: cardW, h: 0.85 }, 38, { align: "ctr", anchor: "t", b: true, color: card.big.color ? S.hex(card.big.color) : color.text }));
      cy += 0.85;
      if (card.big.label) {
        shapes.push(label(card.big.label, { x: cx + 0.18, y: cy, w: cardW - 0.36, h: 0.4 }, Math.max(9, baseTextSize - 1), { align: "ctr", color: color.muted }));
        cy += 0.42;
      }
    }
    if (card.subtitle) {
      shapes.push(label(card.subtitle, { x: cx + 0.18, y: cy, w: cardW - 0.36, h: 0.32 }, Math.max(9, baseTextSize), { align: "ctr", b: true, color: color.text }));
    }
    if (card.stats && card.stats.length) {
      const cols = card.stats.length >= 4 ? 2 : card.stats.length;
      const statW = (cardW - 0.36) / cols;
      card.stats.forEach((s, si) => {
        const sx = cx + 0.18 + (si % cols) * statW;
        const sy = cy + Math.floor(si / cols) * 0.66;
        shapes.push(label(s.value, { x: sx, y: sy, w: statW, h: 0.34 }, 19, { b: true, color: s.color ? S.hex(s.color) : color.text }));
        shapes.push(label(s.label, { x: sx, y: sy + 0.32, w: statW, h: 0.3 }, Math.max(8, baseTextSize - 3), { color: color.muted }));
      });
    }
    if (card.footnote) {
      shapes.push(label(card.footnote, { x: cx + 0.18, y: y + h - 0.34, w: cardW - 0.36, h: 0.3 }, Math.max(8, baseTextSize - 4), { color: color.muted }));
    }
  });
  return shapes;
}

// barchart-panels: one native horizontal ranked bar chart per panel, side by side.
function barPanelShapes(resolved, { x, y, w, h }, S) {
  const { color } = S;
  const shapes = [];
  const gap = 0.2;
  const panelW = (w - gap * (resolved.panels.length - 1)) / resolved.panels.length;
  resolved.panels.forEach((panel, i) => {
    const px = x + i * (panelW + gap);
    const pc = panel.color ? S.hex(panel.color) : PALETTE[i % PALETTE.length];
    const dot = 0.1;
    shapes.push({ kind: "text", x: px, y, w: panelW, h, geometry: "roundRect", radius: 0.1, fill: color.panel, shadow: true, paras: [] });
    shapes.push({ kind: "text", x: px + 0.18, y: y + 0.16, w: dot, h: dot, geometry: "ellipse", fill: pc, paras: [] });
    shapes.push(label(panel.title, { x: px + 0.18 + dot + 0.09, y: y + 0.09, w: panelW - 0.36 - dot - 0.09, h: 0.3 }, 11, { b: true, color: color.text }));
    if (!panel.rows.length) return;
    // PowerPoint bar charts plot category 0 at the bottom; a Chart.js indexAxis:'y' chart (the
    // usual live-preview equivalent) plots it at the top - so reverse the already-sorted rows
    // to keep "highest first" reading top-down.
    const rows = [...panel.rows].reverse();
    shapes.push({
      kind: "chart", x: px + 0.12, y: y + 0.52, w: panelW - 0.24, h: h - 0.68,
      chart: {
        type: "bar", categories: rows.map((r) => String(r.model)),
        series: [{ name: "Value", color: pc, values: rows.map((r) => Number(r.success)) }],
        // The data labels already say the number, so the value axis is dropped entirely.
        valueAxis: false, legend: false, gapWidth: 35,
        dataLabels: { color: pc, size: 11 }, background: color.panel, catLabelColor: color.text, catLabelSize: 10,
      },
    });
  });
  return shapes;
}

// barchart: one native clustered column chart.
function groupedBarShape(resolved, box, S) {
  const { color } = S;
  return {
    kind: "chart", ...box,
    chart: {
      type: "column", categories: resolved.labels.map(String),
      series: resolved.datasets.map((d, j) => ({ name: String(d.name), color: d.color ? S.hex(d.color) : PALETTE[j % PALETTE.length], values: d.data.map(Number) })),
      min: resolved.min, max: resolved.max, valueAxis: true, legend: true, gapWidth: 45, overlap: -6,
      dataLabels: { color: color.text, size: 10 }, background: "lt1",
      catLabelColor: color.text, catLabelSize: 11, valLabelColor: color.muted, valLabelSize: 9, legendColor: color.text, legendSize: 11,
    },
  };
}

function chartShapes(block, box, textSize, S) {
  const r = block.resolved;
  if (!r) {
    const hint = S.config.resolvedDataDir ? "no matching data found" : "no data directory configured";
    return [placeholderBox(box, `[ ${block.name} — interactive chart, ${hint} ]`, textSize, S)];
  }
  if (r.kind === "statcards") return statCardShapes(r.cards, box, textSize, S);
  if (r.kind === "barchart-panels") return barPanelShapes(r, box, S);
  if (r.kind === "barchart") return [groupedBarShape(r, box, S)];
  return [placeholderBox(box, `[ ${block.name} — unknown chart kind "${r.kind}" ]`, textSize, S)];
}

// ---------------------------------------------------------------------------
// Positioned (free) shapes for rich bodies
// ---------------------------------------------------------------------------

function blockShapes(block, x, y, w, textSize, codeStore, scale, S) {
  const { color } = S;
  const h = Math.max(0.25, estimateBlockHeight(block, w, textSize, codeStore) * scale);
  const box = { x, y, w, h };
  const text = (paras, extra = {}) => ({ kind: "text", ...box, geometry: "rect", anchor: "t", paras, ...extra });
  let shapes;
  switch (block.type) {
    case "heading":
    case "paragraph":
    case "bullets":
      shapes = [text(blockParas([block], textSize, S))];
      break;
    case "quote": {
      const indent = 0.18;
      shapes = [
        { kind: "text", x, y, w: 0.04, h, geometry: "rect", fill: color.border, paras: [] },
        { ...text([{ runs: inlineRuns(block.text, { font: "body", size: textSize, i: true, color: color.muted }, S) }]), x: x + indent, w: w - indent },
      ];
      break;
    }
    case "table":
      // Header bold/color comes from the deck table style, so cells only carry size + inline markdown.
      shapes = [{ kind: "table", ...box, header: true, variant: "default", rows: block.rows.map((row) => row.map((c) => [{ runs: inlineRuns(c, { font: "body", size: textSize }, S) }])) }];
      break;
    case "deflist":
      // A compact 2-column reference table (term | description) instead of stacked paragraphs.
      shapes = [{
        kind: "table", ...box, colW: [w * 0.4, w * 0.6], header: false, variant: "deflist",
        rows: block.rows.map((r) => [
          [{ runs: [{ text: r.term.replace(/`/g, ""), font: "mono", size: Math.max(8, textSize - 1), color: color.accentDark }] }],
          [{ runs: inlineRuns(r.desc, { font: "body", size: Math.max(8, textSize - 2), color: color.muted }, S) }],
        ]),
      }];
      break;
    case "code": {
      const size = Math.max(8, textSize - 2);
      const lines = (codeStore[block.idx] || "").split("\n");
      shapes = [text(lines.map((line) => ({ runs: [{ text: line, font: "mono", size, color: color.accentDark }] })), { fill: color.panel, inset: 6 / 72 })];
      break;
    }
    case "image":
      shapes = [imageShape(block, box, textSize, S)];
      break;
    case "styledtext": {
      const st = block.style;
      const size = Math.max(8, Math.round((st.size || textSize) * scale));
      shapes = [text([{ align: st.align === "center" ? "ctr" : undefined, runs: [{ text: block.text, font: "body", size, b: st.bold, color: st.color ? S.hex(st.color) : color.text }] }])];
      break;
    }
    case "chart":
      shapes = chartShapes(block, box, textSize, S);
      break;
    case "cardrow": {
      const gap = 0.18;
      const pad = 0.14;
      const cardW = (w - gap * (block.cards.length - 1)) / block.cards.length;
      shapes = [];
      block.cards.forEach((card, i) => {
        const cx = x + i * (cardW + gap);
        shapes.push({ kind: "text", x: cx, y, w: cardW, h, geometry: "roundRect", radius: 0.06, fill: color.panel, line: { color: color.border, width: 1 }, paras: [] });
        // Each card's own blocks get the real per-type rendering - a table inside a card is a
        // real table, not joined "a | b" text.
        shapes.push(...sequenceShapes(card.blocks, cx + pad, y + pad, cardW - pad * 2, Math.max(8, textSize - 2), codeStore, scale, S, 0.1).shapes);
      });
      break;
    }
    default:
      shapes = [];
  }
  return { h, shapes };
}

function sequenceShapes(blocks, x, y, w, textSize, codeStore, scale, S, gap) {
  let cy = y;
  const shapes = [];
  for (const block of blocks) {
    const r = blockShapes(block, x, cy, w, textSize, codeStore, scale, S);
    shapes.push(...r.shapes);
    cy += r.h + gap;
  }
  return { shapes, h: cy - y };
}

// The old renderer's whole-body flow: pick a text size, scale everything down (to 55% at most)
// if the estimate doesn't fit, then stack the blocks.
function freeBody(blocks, top, available, codeStore, S) {
  const g = S.geometry;
  const textSize = blocks.length > 5 ? 11 : 13;
  const total = blocksHeight(blocks, g.contentW, textSize, codeStore);
  const scale = total > available && total > 0 ? Math.max(0.55, available / total) : 1;
  const drawTextSize = Math.max(9, Math.round(textSize * Math.min(1, scale + 0.3)));
  return sequenceShapes(blocks, g.marginX, top, g.contentW, drawTextSize, codeStore, scale, S, 0.16 * scale);
}

// ---------------------------------------------------------------------------
// Slides
// ---------------------------------------------------------------------------

function layoutCentered(slide, title, textOnly, align, spec, codeStore, S) {
  const g = S.geometry;
  const blocks = slide.blocks;
  const gap = 0.2;
  const eyebrowH = slide.eyebrow ? 0.35 : 0;
  spec.layout = "centered";
  // Room for the body if the stack were pinned to the top margin.
  const maxBodyH = g.height - g.bottomMargin - (g.eyebrow.y + eyebrowH + g.title.h + gap);
  let bodyH = 0;
  let free = null;
  if (textOnly) bodyH = blocksHeight(blocks, g.contentW, BODY_PT, codeStore);
  else if (blocks.length) {
    free = freeBody(blocks, 0, maxBodyH, codeStore, S);
    bodyH = free.h;
  }
  const stackH = eyebrowH + g.title.h + (blocks.length ? gap + Math.min(bodyH, maxBodyH) : 0);
  const top = Math.max(g.eyebrow.y, (g.height - stackH) / 2);

  if (slide.eyebrow) spec.shapes.push(eyebrowShape(slide.eyebrow, { ...g.eyebrow, y: top }, align, S));
  spec.placeholders.title = title;
  spec.placeholders.titleBox = { ...g.title, y: top + eyebrowH };
  const bodyTop = spec.placeholders.titleBox.y + g.title.h + gap;
  if (textOnly) {
    const room = g.height - g.bottomMargin - bodyTop;
    spec.placeholders.body = blockParas(blocks, null, S, align);
    spec.placeholders.bodyBox = { ...g.body, y: bodyTop, h: Math.max(0.5, Math.min(bodyH, room)) };
    if (bodyH > room) spec.placeholders.bodyFontScale = shrinkScale(room, bodyH);
  } else if (free) {
    spec.shapes.push(...free.shapes.map((s) => ({ ...s, y: s.y + bodyTop })));
  }
  return spec;
}

function layoutSlide(slide, index, codeStore, S) {
  const g = S.geometry;
  const fm = slide.frontmatter || {};
  const cls = String(fm.class || "");
  const align = /\btext-center\b/.test(cls) ? "ctr" : undefined;
  const blocks = slide.blocks;
  const spec = { layout: "blank", placeholders: {}, shapes: [] };
  if (slide.notes) spec.notes = slide.notes;
  const title = slide.title ? [{ align, runs: [{ text: slide.title, font: "heading" }] }] : null;

  if (index === 0 || /\bcover\b/.test(cls) || fm.layout === "cover" || fm.layout === "intro") {
    spec.layout = "title";
    if (slide.eyebrow) spec.shapes.push(eyebrowShape(slide.eyebrow, g.eyebrow, undefined, S));
    if (title) spec.placeholders.title = title;
    const first = blocks.find((b) => b.type === "paragraph");
    if (first) spec.placeholders.subtitle = [{ runs: inlineRuns(first.text, { font: "body" }, S) }];
    return spec;
  }

  if (!title) {
    // Nothing for a title placeholder: everything is drawn as shapes on the Blank layout.
    if (slide.eyebrow) spec.shapes.push(eyebrowShape(slide.eyebrow, g.eyebrow, align, S));
    const top = g.title.y;
    spec.shapes.push(...freeBody(blocks, top, g.height - g.bottomMargin - top, codeStore, S).shapes);
    return spec;
  }

  const textOnly = blocks.length > 0 && blocks.every((b) => TEXT_ONLY.has(b.type));
  if (fm.layout === "center" || fm.layout === "end") return layoutCentered(slide, title, textOnly, align, spec, codeStore, S);

  if (fm.layout === "section" && (textOnly || blocks.length === 0)) {
    spec.layout = "section";
    spec.placeholders.title = title;
    if (slide.eyebrow) spec.shapes.push(eyebrowShape(slide.eyebrow, { ...g.eyebrow, y: g.sectionTitle.y - 0.35 }, align, S));
    if (textOnly) spec.placeholders.body = blockParas(blocks, null, S, align);
    return spec;
  }

  if (slide.eyebrow) spec.shapes.push(eyebrowShape(slide.eyebrow, g.eyebrow, align, S));
  spec.placeholders.title = title;
  if (textOnly) {
    spec.layout = "content";
    spec.placeholders.body = blockParas(blocks, null, S, align);
    const needed = blocksHeight(blocks, g.body.w, BODY_PT, codeStore);
    if (needed > g.body.h) spec.placeholders.bodyFontScale = shrinkScale(g.body.h, needed);
    return spec;
  }
  spec.layout = "titleOnly";
  if (blocks.length) spec.shapes.push(...freeBody(blocks, g.body.y, g.body.h, codeStore, S).shapes);
  return spec;
}

export function layoutDeck(parsedSlides, codeStore, config) {
  const S = deckStyle(config);
  return parsedSlides.map((slide, i) => layoutSlide(slide, i, codeStore, S));
}
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `npm test`
Expected: PASS. If the centered-title test fails on `titleBox.y > 2.5`, print the spec (`console.log(JSON.stringify(s, null, 2))`) and check `layoutCentered`'s stack math before touching the assertion.

- [ ] **Step 6: Commit**

```bash
git add package.json package-lock.json src/layout.mjs test/layout.test.mjs
git commit -m "Add layout engine that maps parsed Slidev slides onto SlideSpecs

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 14: Switch `convert()` to the new pipeline; fixture deck; remove pptxgenjs

**Files:**
- Modify: `src/index.mjs` (whole file)
- Modify: `src/default-config.mjs` (add `lang`, `timestamps`)
- Modify: `src/parse.mjs` (delete `mdInlineToRuns` and its export)
- Modify: `test/parse.test.mjs` (delete the `mdInlineToRuns` test and import)
- Modify: `package.json` (remove `pptxgenjs`)
- Delete: `src/render.mjs`, `src/fix-pptx.mjs`
- Create: `test/fixtures/kitchen-sink/slides.md`, `test/fixtures/kitchen-sink/pptx.config.mjs`, `test/fixtures/kitchen-sink/public/data/chart.json`
- Test: `test/convert.test.mjs`

**Interfaces:**
- Consumes: `layoutDeck`, `deckStyle` (Task 13), `writePptx` (Tasks 8–11), `splitSlides`, `parseSlide` (Task 12).
- Produces: `convert(options)` — same signature and return value `{ outputPath, slideCount, translatedCharts, unresolvedCharts }` as before. Deck title = global frontmatter `title`, else the input file's basename.

- [ ] **Step 1: Create the fixture deck** — `test/fixtures/kitchen-sink/slides.md`

````md
---
title: Kitchen Sink
---

# Kitchen Sink Deck

A subtitle with **bold** text

<!--
Cover notes line one.

Cover notes after a blank line.
-->

---

<div class="eyebrow">Text only</div>

# Plain content slide

An intro paragraph with **bold**, *italic* and `code`.

- First bullet
- Second bullet with **bold** inside
- Third bullet

1. Numbered one
2. Numbered two

<!-- Notes for the text slide -->

---

# Rich content slide

| Option | Pros | Cons |
| --- | --- | --- |
| A | Simple | Slower |
| B | Fast | More setup |

```ts
const answer = await agent.ask("Which dataset covers this area?");
```

![A tiny green picture](/pic.png)

---
layout: center
class: text-center
---

# Centered thanks

Questions?

---
layout: section
---

# Section break

What comes next

---

<div class="grid grid-cols-2 gap-4">
<div>

### Left card

Card text with **bold**

</div>
<div>

### Right card

More card text

</div>
</div>

---

# Chart slide

<FixtureChart />

<UnknownWidget />

---

# Stats and panels

<FixtureStats />

<FixturePanels />

---

# Definitions and quotes

<div>
<div class="grid grid-cols-[1fr_2fr]">
<code>search(question)</code>
<div class="opacity-70">Searches the catalog</div>
</div>
<div class="grid grid-cols-[1fr_2fr]">
<code>fetch(id)</code>
<div class="opacity-70">Fetches one dataset</div>
</div>
</div>

> A quoted line

<div class="text-4xl text-center">42%</div>
````

- [ ] **Step 2: Create the fixture config and chart data**

`test/fixtures/kitchen-sink/pptx.config.mjs`:

```js
// Config for the kitchen-sink fixture deck: one resolver per chart kind the exporter supports.
import { readFileSync } from "node:fs";
import { join } from "node:path";

export default {
  author: "Fixture",
  colors: { accent: "1A833B", accentDark: "14532D", accent2: "1A589F", panel: "EEF4F0", border: "D9E3DB", text: "1F2937", muted: "6B7280" },
  fonts: { heading: "Plus Jakarta Sans", body: "Inter", mono: "JetBrains Mono" },
  logo: { path: "./public/pic.png", aspect: 2, height: 0.27 },
  dataDir: "public/data",
  chartResolvers: {
    FixtureChart(dataDir) {
      const d = JSON.parse(readFileSync(join(dataDir, "chart.json"), "utf8"));
      return { kind: "barchart", labels: d.labels, datasets: d.datasets, min: d.min, max: d.max };
    },
    FixtureStats: () => ({
      kind: "statcards",
      cards: [
        { title: "Speed", stats: [{ value: "3.4s", label: "p50" }, { value: "9.1s", label: "p95" }] },
        { title: "Quality", big: { value: "72%", label: "pass rate" }, footnote: "n = 120" },
      ],
    }),
    FixturePanels: () => ({
      kind: "barchart-panels",
      panels: [
        { title: "Model A", color: "1A833B", rows: [{ model: "alpha", success: 80 }, { model: "beta", success: 60 }] },
        { title: "Model B", color: "1A589F", rows: [{ model: "alpha", success: 70 }] },
      ],
      min: 0,
      max: 100,
    }),
  },
};
```

`test/fixtures/kitchen-sink/public/data/chart.json`:

```json
{
  "labels": ["Q1", "Q2", "Q3", "Q4"],
  "datasets": [
    { "name": "Before", "color": "9CA3AF", "data": [22, 31, 28, 35] },
    { "name": "After", "color": "1A833B", "data": [41, 56, 63, 72] }
  ],
  "min": 0,
  "max": 100
}
```

- [ ] **Step 3: Write the failing test** — `test/convert.test.mjs`

```js
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync, mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { convert } from "../src/index.mjs";
import fixtureConfig from "./fixtures/kitchen-sink/pptx.config.mjs";
import { lintPackage } from "./lint.mjs";
import { unzip, byLocal } from "./helpers.mjs";

const DIR = fileURLToPath(new URL("./fixtures/kitchen-sink/", import.meta.url));
const OUT = mkdtempSync(join(tmpdir(), "slidev-pptx-"));

async function exportFixture(name = "deck.pptx") {
  const result = await convert({ inputPath: join(DIR, "slides.md"), outputPath: join(OUT, name), config: fixtureConfig, configDir: DIR });
  const buffer = readFileSync(result.outputPath);
  return { result, buffer, z: await unzip(buffer) };
}

function layoutName(z, n) {
  const rels = z.text(`ppt/slides/_rels/slide${n}.xml.rels`);
  const target = /\/slideLayout" Target="..\/slideLayouts\/(slideLayout\d+\.xml)"/.exec(rels)[1];
  return byLocal(z.xml(`ppt/slideLayouts/${target}`), "cSld")[0].getAttribute("name");
}

test("fixture deck exports, lints clean and reports charts", async () => {
  const { result, buffer } = await exportFixture();
  assert.deepEqual(await lintPackage(buffer), []);
  assert.equal(result.slideCount, 9);
  assert.equal(result.translatedCharts, 3);
  assert.deepEqual(result.unresolvedCharts, ["UnknownWidget"]);
});

test("slides map to the expected layouts", async () => {
  const { z } = await exportFixture();
  assert.deepEqual([1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => layoutName(z, n)), [
    "Title Slide", "Title and Content", "Title Only", "Centered", "Section Header", "Blank", "Title Only", "Title Only", "Title Only",
  ]);
});

test("titles and text-only bodies live in placeholders", async () => {
  const { z } = await exportFixture();
  assert.match(z.text("ppt/slides/slide1.xml"), /<p:ph type="ctrTitle"\/>.*<a:t>Kitchen Sink Deck<\/a:t>/);
  const s2 = z.text("ppt/slides/slide2.xml");
  assert.match(s2, /<p:ph type="title"\/>.*<a:t>Plain content slide<\/a:t>/);
  assert.match(s2, /<p:ph idx="1"\/>.*<a:t>First bullet<\/a:t>/);
  assert.match(s2, /<a:buAutoNum type="arabicPeriod"\/>/);
});

test("speaker notes are exported", async () => {
  const { z } = await exportFixture();
  assert.match(z.text("ppt/notesSlides/notesSlide1.xml"), /<a:t>Cover notes line one.<\/a:t>.*<a:t>Cover notes after a blank line.<\/a:t>/);
  assert.match(z.text("ppt/notesSlides/notesSlide2.xml"), /<a:t>Notes for the text slide<\/a:t>/);
});

test("charts carry an embedded workbook with the data", async () => {
  const { z } = await exportFixture();
  assert.deepEqual(z.names.filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n)).length, 3);
  const book = await unzip(z.buffer("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"));
  assert.match(book.text("xl/sharedStrings.xml"), /Before/);
});

test("theme, fonts, logo and document title come from the deck", async () => {
  const { z } = await exportFixture();
  const theme = z.text("ppt/theme/theme1.xml");
  assert.match(theme, /<a:accent1><a:srgbClr val="1A833B"\/><\/a:accent1>/);
  assert.match(theme, /<a:majorFont><a:latin typeface="Plus Jakarta Sans"\/>/);
  assert.match(z.text("ppt/slideMasters/slideMaster1.xml"), /name="Logo"/);
  assert.match(z.text("docProps/core.xml"), /<dc:title>Kitchen Sink<\/dc:title><dc:creator>Fixture<\/dc:creator>/);
});

test("two exports of the same deck are byte-identical", async () => {
  const a = await exportFixture("a.pptx");
  const b = await exportFixture("b.pptx");
  assert.ok(a.buffer.equals(b.buffer));
});
```

- [ ] **Step 4: Run test to verify it fails**

Run: `npm test`
Expected: FAIL — the current `convert()` still uses pptxgenjs (lint fails on duplicate `a:pPr`/three `c:axId`, layouts are all `DEFAULT`).

- [ ] **Step 5: Replace** `src/index.mjs`

```js
// Public API: convert(options) turns a Slidev markdown deck into a fully editable native
// .pptx - real PowerPoint placeholders, text, tables, shapes and charts on a themed slide
// master, not the screenshot-per-slide export `slidev export` produces. See README.md for the
// config shape and bin/cli.mjs for the command-line wrapper.

import { readFileSync, existsSync } from "node:fs";
import { writeFile } from "node:fs/promises";
import { resolve, dirname, basename, extname } from "node:path";
import { mergeConfig } from "./config.mjs";
import { splitSlides, parseSlide } from "./parse.mjs";
import { layoutDeck, deckStyle } from "./layout.mjs";
import { writePptx } from "./ooxml/index.mjs";

const LOGO_EXTS = new Set([".png", ".jpg", ".jpeg", ".gif"]);

// Footer logo drawn bottom-left on every slide (it lives on the slide master), or null when
// none is configured or the file isn't a format PowerPoint embeds natively.
function logoFor(config, geometry) {
  const logo = config.logo;
  if (!logo || !logo.path || !existsSync(logo.path)) return null;
  if (!LOGO_EXTS.has(extname(logo.path).toLowerCase())) {
    console.warn(`Warning: logo ${logo.path} is not PNG/JPEG/GIF - skipped.`);
    return null;
  }
  const h = logo.height || 0.27;
  return { path: logo.path, box: { x: 0.5, y: geometry.height - 0.29 - h, w: h * (logo.aspect || 1), h } };
}

/**
 * @param {object} options
 * @param {string} options.inputPath - path to the deck's .md file
 * @param {string} [options.outputPath] - defaults to "<input-basename>-editable.pptx" next to inputPath
 * @param {object} [options.config] - see default-config.mjs; typically loaded from a deck's pptx.config.mjs
 * @param {string} [options.configDir] - directory config-relative paths (logo.path) resolve against; defaults to inputPath's directory
 * @returns {Promise<{ outputPath: string, slideCount: number, translatedCharts: number, unresolvedCharts: string[] }>}
 */
export async function convert(options) {
  const inputPath = resolve(options.inputPath);
  const raw = readFileSync(inputPath, "utf8");
  const inputDir = dirname(inputPath);
  const configDir = options.configDir ? resolve(options.configDir) : inputDir;

  const config = mergeConfig(options.config || {});
  config.inputDir = inputDir;
  config.publicDir = resolve(inputDir, "public");
  if (!existsSync(config.publicDir)) config.publicDir = null;

  const dataDir = resolve(inputDir, config.dataDir);
  config.resolvedDataDir = existsSync(dataDir) ? dataDir : null;
  if (!config.resolvedDataDir && Object.keys(config.chartResolvers).length) {
    console.warn(`Warning: data directory "${config.dataDir}" not found next to ${basename(inputPath)} - charts will fall back to placeholders.`);
  }

  if (config.logo && config.logo.path) {
    config.logo = { ...config.logo, path: resolve(configDir, config.logo.path) };
  }

  const { globalFrontmatter, slides } = splitSlides(raw);
  const codeStore = [];
  const parsed = slides.map((s) => ({ ...parseSlide(s.content, codeStore, config), frontmatter: s.frontmatter }));
  const specs = layoutDeck(parsed, codeStore, config);
  const style = deckStyle(config);

  const buffer = await writePptx(specs, {
    geometry: style.geometry,
    theme: style.theme,
    fonts: config.fonts,
    lang: config.lang,
    title: globalFrontmatter.title || basename(inputPath, extname(inputPath)),
    author: config.author,
    timestamps: config.timestamps === true,
    logo: logoFor(config, style.geometry),
  });

  const outputPath = resolve(options.outputPath || `${basename(inputPath, extname(inputPath))}-editable.pptx`);
  await writeFile(outputPath, buffer);

  const chartBlocks = parsed.flatMap((s) => s.blocks).filter((b) => b.type === "chart");
  const translatedCharts = chartBlocks.filter((b) => b.resolved).length;
  const unresolvedCharts = [...new Set(chartBlocks.filter((b) => !b.resolved).map((b) => b.name))];

  return { outputPath, slideCount: parsed.length, translatedCharts, unresolvedCharts };
}

export { mergeConfig } from "./config.mjs";
export { defaultConfig } from "./default-config.mjs";
```

Note: chart blocks nested inside card rows aren't counted - same as before (`parsed.flatMap((s) => s.blocks)` only looks at top-level blocks).

- [ ] **Step 6: Add the new config keys** — in `src/default-config.mjs`, after the `author: "",` line add:

```js

  // Language tag written on every text run - PowerPoint uses it for spell-checking.
  lang: "en-US",

  // true writes the export time into docProps/core.xml. Off by default so exporting the same
  // deck twice gives a byte-identical .pptx (clean diffs, reproducible builds).
  timestamps: false,
```

- [ ] **Step 7: Remove the legacy renderer and pptxgenjs**

```bash
git rm src/render.mjs src/fix-pptx.mjs
npm uninstall pptxgenjs
```

In `src/parse.mjs`, delete the `mdInlineToRuns` function (and its comment) and the final `export { mdInlineToRuns };` line. In `test/parse.test.mjs`, delete the `"mdInlineToRuns still produces the old pptxgenjs run shape"` test and drop `mdInlineToRuns` from the import.

Run: `grep -rn "pptxgenjs\|mdInlineToRuns\|render.mjs\|fix-pptx" src bin test package.json`
Expected: no matches.

- [ ] **Step 8: Run tests to verify they pass**

Run: `npm test`
Expected: PASS (all files)

- [ ] **Step 9: Check the CLI end to end**

Run: `node bin/cli.mjs test/fixtures/kitchen-sink/slides.md "${TMPDIR:-/tmp}/kitchen-sink.pptx"`
Expected (three lines):
```
Wrote 9 slides to .../kitchen-sink.pptx
Translated 3 chart(s) into native editable PowerPoint content.
No resolver for: UnknownWidget - exported as placeholder box(es). Add a resolver to chartResolvers in your pptx.config.mjs to translate them too.
```

- [ ] **Step 10: Commit**

```bash
git add -A src test package.json package-lock.json
git commit -m "Switch convert() to the OOXML writer and drop pptxgenjs

Titles and text-only bodies now go into real layout placeholders on a
themed slide master, speaker notes are exported, Slidev center/section/
cover layouts map to matching PowerPoint layouts, and output is
byte-reproducible. A kitchen-sink fixture deck covers every block type.

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 15: LibreOffice render check, README, version

**Files:**
- Create: `scripts/render.mjs`
- Modify: `.gitignore` (add `out/`)
- Modify: `README.md`
- Modify: `package.json` (`"version": "0.2.0"`; `scripts/` stays out of the published `files`)

**Interfaces:**
- Consumes: `convert` (Task 14), `splitSlides`, `parseSlide` (Task 12), `mergeConfig`.
- Produces: `npm run render [-- deck.md ...]` — exports each deck to `out/render/<deck-dir>/deck.pptx`, converts to PDF with LibreOffice, writes `slide-N.png` pages, checks page count == slide count and every slide title appears in `pdftotext` output. Exit code 1 on a failed check, 0 (with a message) when `soffice` is missing.

- [ ] **Step 1: Implement** — `scripts/render.mjs`

```js
#!/usr/bin/env node
// LibreOffice render check. Exports decks, converts them to PDF and per-slide PNGs for visual
// review, and checks page count and slide titles with poppler. Skips (exit 0) when soffice
// isn't installed. Usage: node scripts/render.mjs [path/to/slides.md ...]
// Default decks: the kitchen-sink fixture, plus ../innotek-SlideDevBase if it's checked out.

import { execFileSync, spawnSync } from "node:child_process";
import { existsSync, mkdirSync, rmSync, readFileSync } from "node:fs";
import { resolve, dirname, basename, join } from "node:path";
import { pathToFileURL, fileURLToPath } from "node:url";
import { convert, mergeConfig } from "../src/index.mjs";
import { splitSlides, parseSlide } from "../src/parse.mjs";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const DEFAULT_DECKS = [join(ROOT, "test/fixtures/kitchen-sink/slides.md"), resolve(ROOT, "../innotek-SlideDevBase/slides.md")];

const has = (cmd) => spawnSync("which", [cmd]).status === 0;
const squash = (s) => s.replace(/\s+/g, " ");

async function loadConfig(deckPath) {
  const path = join(dirname(deckPath), "pptx.config.mjs");
  if (!existsSync(path)) return { config: {}, configDir: dirname(deckPath) };
  return { config: (await import(pathToFileURL(path).href)).default || {}, configDir: dirname(path) };
}

function slideTitles(deckPath, config) {
  const merged = mergeConfig(config);
  const { slides } = splitSlides(readFileSync(deckPath, "utf8"));
  return slides.map((s) => parseSlide(s.content, [], merged).title).filter(Boolean);
}

async function main() {
  if (!has("soffice")) {
    console.log("soffice (LibreOffice) not found - skipping render check.");
    return;
  }
  const args = process.argv.slice(2).map((p) => resolve(p));
  const decks = (args.length ? args : DEFAULT_DECKS).filter((p) => existsSync(p));
  let failed = false;
  for (const deck of decks) {
    const name = basename(dirname(deck));
    const outDir = join(ROOT, "out/render", name);
    rmSync(outDir, { recursive: true, force: true });
    mkdirSync(outDir, { recursive: true });

    const { config, configDir } = await loadConfig(deck);
    const result = await convert({ inputPath: deck, outputPath: join(outDir, "deck.pptx"), config, configDir });
    execFileSync("soffice", ["--headless", "--convert-to", "pdf", "--outdir", outDir, result.outputPath], { stdio: "ignore", timeout: 180000 });
    const pdf = join(outDir, "deck.pdf");
    const pages = Number(/Pages:\s+(\d+)/.exec(execFileSync("pdfinfo", [pdf], { encoding: "utf8" }))[1]);
    const text = squash(execFileSync("pdftotext", [pdf, "-"], { encoding: "utf8" }));
    const missing = slideTitles(deck, config).filter((t) => !text.includes(squash(t)));
    execFileSync("pdftoppm", ["-r", "60", "-png", pdf, join(outDir, "slide")]);

    const ok = pages === result.slideCount && missing.length === 0;
    failed ||= !ok;
    const detail = missing.length ? `, titles missing from the render: ${missing.join(" | ")}` : "";
    console.log(`${ok ? "ok  " : "FAIL"} ${name}: ${pages}/${result.slideCount} pages${detail} -> ${outDir}`);
  }
  if (failed) process.exitCode = 1;
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
```

- [ ] **Step 2: Ignore render output** — append to `.gitignore`:

```
out/
```

- [ ] **Step 3: Run the render check**

Run: `npm run render`
Expected: `ok   kitchen-sink: 9/9 pages -> .../out/render/kitchen-sink` and (if present) `ok   innotek-SlideDevBase: 7/7 pages -> ...`.

Then look at the PNGs (`out/render/*/slide-*.png`) and compare against the pre-change baseline render: cover centered title + subtitle, bullets with green markers on the text slide (including the `**bold**` bullet), table with a panel-colored header, centered "Thanks"/"Centered" slides vertically centered, charts with data labels, logo bottom-left on every slide. Fix any regression in `layout.mjs`/`ooxml/` (with a test) before continuing.

- [ ] **Step 4: Update `README.md`**

Insert this new section directly after the intro paragraphs (before `## Usage`):

````md
## What you get in PowerPoint

- **Real layouts and placeholders.** Slide titles sit in the layout's title placeholder, so
  outline view, the Accessibility Checker, Reset Slide and layout switching work. Text-only
  slides (paragraphs, bullets, sub-headings) put their body in the content placeholder, with
  native bullets and numbering and "shrink text on overflow". Slides with cards, tables, code,
  charts or images use the *Title Only* layout plus positioned shapes.
- **A theme built from your config.** `colors` and `fonts` become the deck's theme colors and
  heading/body fonts, so *Design ▸ Variants* recolors or re-fonts the whole deck.
- **Slidev layouts mapped:** first slide / `cover` / `intro` → *Title Slide*, `center` / `end`
  → *Centered*, `section` → *Section Header*, everything else → *Title and Content* or *Title
  Only*.
- **Speaker notes** from the slide's last `<!-- ... -->` comment (Slidev's convention).
- **Native tables** styled by a deck table style, and **native charts** with an embedded
  workbook (*Edit Data* works).
- **Clean files:** content types and relationships are generated from the parts actually
  written (no repair prompts in PowerPoint Online), and the same input always produces a
  byte-identical `.pptx`.
````

In the `## Config` code block, add these two keys after `slide: {...}`:

```js
  lang: "en-US",     // language tag on every text run (spell-check)
  timestamps: false, // true = write the export time into the file (breaks byte-identical output)
```

Add a section before `## What it doesn't do`:

````md
## Development

```bash
npm test         # unit tests, package lint, kitchen-sink fixture export
npm run render   # LibreOffice render of the fixture (and ../innotek-SlideDevBase if present) to out/render/
```
````

- [ ] **Step 5: Bump the version** — in `package.json` set `"version": "0.2.0"`.

- [ ] **Step 6: Full verification**

Run: `npm test && npm run render`
Expected: all tests pass; render check prints `ok` for every deck.

- [ ] **Step 7: Commit**

```bash
git add scripts/render.mjs .gitignore README.md package.json
git commit -m "Add LibreOffice render check, document the native output, bump to 0.2.0

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

## Manual spot check (owner)

After Task 15, open `out/render/kitchen-sink/deck.pptx` and `out/render/innotek-SlideDevBase/deck.pptx` in each target app and check:

| Check | PowerPoint desktop | PowerPoint Online | Google Slides | Keynote |
|---|---|---|---|---|
| Opens with no repair prompt | | | | |
| Outline view / slide titles present | | | n/a | n/a |
| Design ▸ Variants recolors the deck | | | n/a | n/a |
| Table header is panel-colored with bold dark text | | | | |
| Chart shows bars + labels; Edit Data opens the workbook | | | (bars only) | (bars only) |
| Speaker notes present | | | | |
| Logo bottom-left on every slide | | | | |

If tables look unstyled in Keynote or Google Slides, follow the fallback in the spec's Risks section (direct cell borders/fills in `src/ooxml/table.mjs`).
