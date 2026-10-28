import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { sceneToSpec } from "../src/render/scene.mjs";
import { slideGeometry } from "../src/geometry.mjs";
import { themeColors } from "../src/colors.mjs";
import { writePptx } from "../src/ooxml/index.mjs";
import { DECK, PIC } from "./deck.mjs";
import { lintPackage } from "./lint.mjs";
import { unzip } from "./helpers.mjs";

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
  // left-aligned stated outright: a placeholder would inherit its layout's (Title Slide centres)
  assert.deepEqual(spec.placeholders.title, [{ align: "l", lineSpacing: pt(40), runs: [{ text: "Card row", size: pt(36), font: "heading", b: true, color: "dk1" }] }]);
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

test("a shape drawn before the title in DOM order sits beneath it: titleZ counts those shapes", () => {
  const hero = { kind: "box", rect: { x: 0, y: 100, w: 980, h: 300 }, fill: C("EEF4F0"), border: null, sides: null, radius: 0, shadow: null, opacity: 1 };
  const spec = sceneToSpec(scene([hero, text([run("Over the band", { size: 36 })], { role: "title", rect: { x: 56, y: 180, w: 500, h: 40 }, lineHeight: 40 })]), STYLE);
  assert.equal(spec.placeholders.titleZ, 1);
  assert.equal(spec.shapes.length, 1);
});

test("a title drawn first has no titleZ", () => {
  const card = { kind: "box", rect: { x: 0, y: 0, w: 100, h: 100 }, fill: C("EEF4F0"), border: null, sides: null, radius: 0, shadow: null, opacity: 1 };
  const spec = sceneToSpec(scene([text([run("Title", { size: 36 })], { role: "title" }), card]), STYLE);
  assert.equal("titleZ" in spec.placeholders, false);
});

test("a regular-weight title says so: the title placeholder would inherit the master's bold title style", async () => {
  const spec = sceneToSpec(scene([text([run("Release notes", { size: 36 }), run(" 0.4", { size: 36, bold: true })], { role: "title" })]), STYLE);
  assert.deepEqual(spec.placeholders.title[0].runs.map((r) => r.b), [false, true]);
  const files = await unzip(await writePptx([spec], { ...DECK, logo: null }));
  const title = files.text("ppt/slides/slide1.xml").match(/<p:sp>(?:(?!<\/p:sp>).)*type="title".*?<\/p:sp>/)[0];
  assert.match(title, /<a:rPr [^>]*b="0"[^>]*>(?:(?!<\/a:rPr>).)*<\/a:rPr><a:t>Release notes<\/a:t>/);
  assert.match(title, /<a:rPr [^>]*b="1"[^>]*>(?:(?!<\/a:rPr>).)*<\/a:rPr><a:t> 0.4<\/a:t>/);
});

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

test("table cells carry row spans, bottom anchors and the node's opacity on fills and borders", () => {
  const cell = (t, extra = {}) => ({ runs: [run(t, { size: 14 })], align: "left", valign: "bottom", lineHeight: 0, fill: null, borders: {}, padding: { t: 0, r: 0, b: 0, l: 0 }, colspan: 1, rowspan: 1, ...extra });
  const node = {
    kind: "table", rect: { x: 40, y: 40, w: 200, h: 60 }, colW: [100, 100], rowH: [30, 30], opacity: 0.5,
    rows: [
      [cell("a", { rowspan: 2, fill: C("EEF4F0", 0.8), borders: { l: { width: 1, color: C("D9E3DB", 0.4) } } }), cell("b", { valign: "middle" })],
      [{ covered: "v" }, cell("c")],
    ],
  };
  const [t] = sceneToSpec(scene([node]), STYLE).shapes;
  const [a, b] = t.rows[0];
  assert.deepEqual({ rowSpan: a.rowSpan, anchor: a.anchor, fill: a.fill, fillAlpha: a.fillAlpha, l: a.borders.l, para: a.paras[0].lineSpacing }, {
    rowSpan: 2, anchor: "b", fill: "lt2", fillAlpha: 0.4, l: { width: pt(1), color: "accent4", alpha: 0.2 }, para: undefined,
  });
  assert.equal(b.anchor, "ctr");
  assert.equal(b.paras[0].runs[0].color, "dk1");
  assert.equal(b.paras[0].runs[0].alpha, 0.5);
  assert.deepEqual(t.rows[1][0], { paras: [], hMerge: false, vMerge: true });
});

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

test("a horizontal chart lists its categories top-down; a translucent series keeps its alpha; a hidden value axis stays hidden", () => {
  const node = {
    kind: "chart", id: 4, rect: { x: 40, y: 200, w: 900, h: 240 }, components: [], opacity: 1,
    chart: {
      type: "bar", categories: ["Q1", "Q2"], series: [{ name: "A", values: [22, 31], color: C("1A833B", 0.4) }, { name: "B", values: [1, 2], color: C("9CA3AF") }],
      min: 0, max: null, stacked: false, valueAxis: false, legend: false, legendPos: "t", numFmt: null, gridColor: null, labelColor: null, legendColor: null,
    },
  };
  const { chart } = sceneToSpec(scene([node]), STYLE).shapes[0];
  assert.equal(chart.reverseCategories, true);
  assert.equal(chart.valueAxis, false);
  assert.deepEqual(chart.series.map((s) => [s.color, s.alpha]), [["accent1", 0.4], ["accent5", undefined]]);
  assert.ok(!("gridColor" in chart) && !("max" in chart));
  const column = sceneToSpec(scene([{ ...node, chart: { ...node.chart, type: "column", valueAxis: true } }]), STYLE).shapes[0].chart;
  assert.ok(!("reverseCategories" in column));
  assert.equal(column.valueAxis, true);
});

test("gradients: radial backgrounds composite over their base and rescale to the corner radius; linear angles turn into DrawingML's", () => {
  const bg = {
    gradient: { type: "radial", cx: 0.5, cy: 0.5, rx: 100, ry: 100, stops: [{ pos: 0, color: C("1A833B", 0.5) }, { pos: 0.5, color: { hex: "000000", alpha: 0 } }] },
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

test("gradient stops follow the box's opacity and run the full 0..1 range", () => {
  const node = {
    kind: "box", rect: { x: 0, y: 0, w: 100, h: 100 }, border: null, sides: null, radius: 0, shadow: null, opacity: 0.5,
    fill: { gradient: { type: "linear", angle: 180, stops: [{ pos: 0.2, color: C("FF0000") }, { pos: 0.6, color: C("0000FF", 0) }] }, base: null },
  };
  const [shape] = sceneToSpec(scene([node]), STYLE).shapes;
  assert.deepEqual(shape.fill, {
    type: "linear", angle: 90,
    stops: [{ pos: 0, color: "#FF0000", alpha: 0.5 }, { pos: 0.2, color: "#FF0000", alpha: 0.5 }, { pos: 0.6, color: "#FF0000", alpha: 0 }, { pos: 1, color: "#FF0000", alpha: 0 }],
  });
  const opaque = sceneToSpec(scene([{ ...node, fill: { ...node.fill, base: C("FFFFFF") } }]), STYLE).shapes[0].fill;
  assert.deepEqual(opaque.stops.map((x) => [x.color, x.alpha]), [["#FF0000", 0.5], ["#FF0000", 0.5], ["lt1", 0.5], ["lt1", 0.5]]);
});

const gradBox = (fill, extra = {}) => ({
  kind: "box", rect: { x: 0, y: 0, w: 200, h: 100 }, border: null, sides: null, radius: 0, shadow: null, opacity: 1, fill, ...extra,
});

test("a translucent base composites under every stop (source-over), then the box's opacity scales the result", () => {
  const gradient = { type: "linear", angle: 90, stops: [{ pos: 0, color: C("FF0000") }, { pos: 0.5, color: C("0000FF", 0.5) }, { pos: 1, color: C("0000FF", 0) }] };
  // a = as + ab(1 - as); c = (cs*as + cb*ab*(1 - as)) / a. Stop 2: a = 0.5 + 0.25 = 0.75, c = (0*0.5 + 255*0.25) / 0.75 = 85.
  const [shape] = sceneToSpec(scene([gradBox({ gradient, base: C("FFFFFF", 0.5) })]), STYLE).shapes;
  assert.deepEqual(shape.fill.stops, [
    { pos: 0, color: "#FF0000" },
    { pos: 0.5, color: "#5555FF", alpha: 0.75 },
    { pos: 1, color: "lt1", alpha: 0.5 },
  ]);
  const [faded] = sceneToSpec(scene([gradBox({ gradient, base: C("FFFFFF", 0.5) }, { opacity: 0.5 })]), STYLE).shapes;
  assert.deepEqual(faded.fill.stops.map((x) => x.alpha), [0.5, 0.375, 0.25]);
});

test("stops outside 0..1 are cut there, with the color at the cut interpolated", () => {
  const fillOf = (stops) => sceneToSpec(scene([gradBox({ gradient: { type: "linear", angle: 90, stops }, base: null })]), STYLE).shapes[0].fill.stops;
  // red -> blue over 0..150%: at 100% the gradient is two thirds of the way there.
  assert.deepEqual(fillOf([{ pos: 0, color: C("FF0000") }, { pos: 1.5, color: C("0000FF") }]), [{ pos: 0, color: "#FF0000" }, { pos: 1, color: "#5500AA" }]);
  // red at -25%: at 0% it is a fifth of the way to blue.
  assert.deepEqual(fillOf([{ pos: -0.25, color: C("FF0000") }, { pos: 1, color: C("0000FF") }]), [{ pos: 0, color: "#CC0033" }, { pos: 1, color: "#0000FF" }]);
  // A middle stop stays; only the stop past 1 is replaced by the color at 1.
  assert.deepEqual(fillOf([{ pos: 0, color: C("FF0000") }, { pos: 0.5, color: C("00FF00") }, { pos: 2, color: C("0000FF") }]), [{ pos: 0, color: "#FF0000" }, { pos: 0.5, color: "#00FF00" }, { pos: 1, color: "#00AA55" }]);
  // Every stop beyond the box: the end color fills it.
  assert.deepEqual(fillOf([{ pos: 1.2, color: C("FF0000") }, { pos: 2, color: C("0000FF") }]), [{ pos: 0, color: "#FF0000" }, { pos: 1, color: "#FF0000" }]);
});

test("radial stops scale by the ellipse's equal-area radius over the focus-to-corner distance", () => {
  const gradient = { type: "radial", cx: 0.5, cy: 0.5, rx: 141.42, ry: 70.71, stops: [{ pos: 0, color: C("000000") }, { pos: 1, color: C("FFFFFF") }] };
  const [shape] = sceneToSpec(scene([gradBox({ gradient, base: null })]), STYLE).shapes;
  const k = Math.round((Math.sqrt(141.42 * 70.71) / Math.hypot(100, 50)) * 10000) / 10000;
  assert.deepEqual(shape.fill.stops.map((x) => x.pos), [0, k, 1]);
  // A circle's equal-area radius is its radius.
  const circle = { ...gradient, rx: 100, ry: 100 };
  const [round] = sceneToSpec(scene([gradBox({ gradient: circle, base: null })]), STYLE).shapes;
  assert.deepEqual(round.fill.stops.map((x) => x.pos), [0, Math.round((100 / Math.hypot(100, 50)) * 10000) / 10000, 1]);
});

test("an SVG image keeps its SVG bytes next to the PNG", () => {
  const svg = Buffer.from("<svg/>");
  const png = Buffer.from("png");
  const node = { kind: "image", id: 1, rect: { x: 0, y: 0, w: 80, h: 40 }, src: "/logo.svg", format: "svg", fit: "fill", natural: null, alt: "Logo", border: null, shadow: null, opacity: 1, data: png, ext: "png", svg };
  const [shape] = sceneToSpec(scene([node]), STYLE).shapes;
  assert.equal(shape.svg, svg);
  assert.equal(shape.data, png);
});
