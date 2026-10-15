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
