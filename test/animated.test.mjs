import { test } from "node:test";
import assert from "node:assert/strict";
import { slideSpecs, slidevTransition } from "../src/render/animated.mjs";
import { slideGeometry } from "../src/geometry.mjs";
import { themeColors } from "../src/colors.mjs";
import { writePptx } from "../src/ooxml/index.mjs";
import { DECK } from "./deck.mjs";
import { lintPackage } from "./lint.mjs";
import { unzip } from "./helpers.mjs";

const STYLE = {
  geometry: slideGeometry({ width: 13.333, height: 7.5, marginX: 0.6 }),
  theme: themeColors({ accent: "1A833B", accentDark: "14532D", accent2: "1A589F", panel: "EEF4F0", border: "D9E3DB", text: "1F2937", muted: "6B7280" }),
  fonts: { heading: "Plus Jakarta Sans", body: "Inter", mono: "JetBrains Mono" },
};
const C = (hex) => ({ hex, alpha: 1 });
const CANVAS = { w: 980, h: 552 };
const run = (text) => ({ text, font: "Inter", size: 20, bold: false, italic: false, color: C("1F2937"), spacing: 0, underline: false, strike: false, highlight: null });
const text = (t, role = "body") => ({ kind: "text", role, rect: { x: 50, y: 50, w: 300, h: 40 }, align: "left", lineHeight: 24, singleLine: true, opacity: 1, runs: [run(t)] });
const title = (t) => text(t, "title");
const box = (hex) => ({ kind: "box", rect: { x: 10, y: 100, w: 100, h: 100 }, fill: C(hex) });
const raster = (hash) => ({ kind: "raster", rect: { x: 0, y: 0, ...CANVAS }, reason: "canvas", label: "canvas", png: Buffer.from(hash) });
const cell = (t) => ({ runs: [run(t)], align: "left", valign: "top", lineHeight: 20, fill: null, borders: {}, padding: { t: 4, r: 8, b: 4, l: 8 }, colspan: 1, rowspan: 1 });
const table = (t) => ({ kind: "table", rect: { x: 50, y: 200, w: 300, h: 30 }, colW: [300], rowH: [30], opacity: 1, rows: [[cell(t)]] });
const CLIP = { data: Buffer.from("mp4"), poster: Buffer.from("jpeg"), duration: 2, natural: { w: 1920, h: 1080 } };
const fail = (line) => assert.fail(`unexpected notice: ${line}`);

const scene = (nodes) => ({ index: 3, canvas: CANVAS, background: { color: C("FFFFFF") }, nodes });
const slideOf = (states, clicks = [], meta = {}) => ({ index: 3, sourceIndex: 2, filepath: null, notes: null, transition: null, states: states.map(scene), clicks, ...meta });
const REVEAL = { kind: "reveal" };
const MOTION = { kind: "motion", clip: CLIP };

test("transitions: the forward half maps; anything else is null", () => {
  assert.equal(slidevTransition("fade"), "fade");
  assert.equal(slidevTransition("fade-out"), "fade");
  assert.deepEqual(slidevTransition("slide-left | fade"), { push: "l" });
  assert.deepEqual(slidevTransition("slide-right"), { push: "r" });
  assert.deepEqual(slidevTransition("slide-up"), { push: "u" });
  assert.deepEqual(slidevTransition("slide-down"), { push: "d" });
  assert.equal(slidevTransition("view-transition"), null);
  assert.equal(slidevTransition(undefined), null);
  assert.equal(slidevTransition(null), null);
});

test("a transition written as an object maps by its name", () => {
  assert.equal(slidevTransition({ name: "fade", mode: "out-in" }), "fade");
  assert.deepEqual(slidevTransition({ name: "slide-left | fade" }), { push: "l" });
  assert.equal(slidevTransition({ mode: "out-in" }), null);
});

test("an unknown transition written as an object is named in its notice", () => {
  const lines = [];
  slideSpecs(slideOf([[box("FF0000")]], [], { transition: { name: "zoom" } }), STYLE, { notice: (l) => lines.push(l) });
  assert.match(lines[0], /^Notice: slide 3: transition "zoom" has no form/);
});

test("a click that changes the slide's background opens a new slide in that state, with a fade", () => {
  const slide = slideOf([[box("FF0000")], [box("FF0000"), box("00FF00")], [box("FF0000"), box("00FF00"), box("0000FF")]], [null, REVEAL, REVEAL]);
  slide.states[1].background = { color: C("102030") };
  slide.states[2].background = { color: C("102030") };
  const specs = slideSpecs(slide, STYLE, { notice: fail });
  assert.equal(specs.length, 2);
  assert.equal(specs[0].background, undefined);
  assert.deepEqual(specs[1].background, { color: "#102030" });
  assert.equal(specs[1].transition, "fade");
  assert.deepEqual(specs[1].shapes.filter((s) => s.effects?.length).map((s) => s.effects), [[{ click: 1, dir: "in" }]]);
});

test("a slide with no clicks is one spec with no effects", () => {
  const specs = slideSpecs(slideOf([[title("T"), box("FF0000")]]), STYLE, { notice: fail });
  assert.equal(specs.length, 1);
  assert.ok(specs[0].shapes.every((s) => !s.effects?.length));
  assert.ok(specs[0].placeholders.title);
});

test("reveals stay on one spec: a table per state, swapped on the click", () => {
  const states = [[title("T"), table("a")], [title("T"), table("b")]];
  const [spec, ...rest] = slideSpecs(slideOf(states, [null, REVEAL]), STYLE, { notice: fail });
  assert.equal(rest.length, 0);
  assert.ok(spec.animated);
  const effectsOfTable = (t) => spec.shapes.find((s) => s.kind === "table" && s.rows[0][0].paras[0].runs[0].text === t).effects;
  assert.deepEqual(effectsOfTable("a"), [{ click: 1, dir: "out" }]);
  assert.deepEqual(effectsOfTable("b"), [{ click: 1, dir: "in" }]);
  assert.equal(spec.shapes.filter((s) => s.kind === "table").length, 2);
  assert.ok(spec.placeholders.title, "a title that stays stays a placeholder");
});

test("a title that changes on a reveal becomes plain text shapes", () => {
  const states = [[title("One")], [title("Two")]];
  const [spec] = slideSpecs(slideOf(states, [null, REVEAL]), STYLE, { notice: fail });
  assert.ok(!spec.placeholders.title);
  assert.equal(spec.shapes.filter((s) => s.effects?.length).length, 2);
});

test("a motion click: a slide holding only the clip, which fills it and starts by itself, with no transition", () => {
  const states = [[raster("old"), title("T")], [raster("map"), title("T")]];
  const specs = slideSpecs(slideOf(states, [null, MOTION], { transition: "fade", notes: "N" }), STYLE, { notice: () => {} });
  assert.equal(specs.length, 2);
  assert.equal(specs[0].transition, "fade");
  assert.equal(specs[1].transition, undefined);
  assert.equal(specs[1].layout, "blank");
  assert.equal(specs[1].shapes.length, 1);
  assert.equal(specs[1].placeholders.title, undefined);
  assert.ok(specs[1].animated);
  const [video] = specs[1].shapes;
  assert.equal(video.kind, "video");
  assert.equal(video.autoplay, true);
  assert.equal(video.muted, true);
  assert.equal(video.loop, false);
  assert.equal(video.ext, "mp4");
  assert.equal(video.duration, 2);
  assert.deepEqual(video.poster, { data: CLIP.poster, ext: "jpg" });
  assert.equal(video.crop, undefined);
  assert.deepEqual([video.x, video.y], [0, 0]);
  assert.ok(Math.abs(video.w - 13.333) < 0.001 && Math.abs(video.h - 7.51) < 0.01, `${video.w} x ${video.h}`);
  assert.match(video.alt, /Slide 3, click 1/);
  assert.deepEqual(specs.map((s) => s.notes), ["N", "N"]);
});

test("a clip on a deck's first slide is still a blank slide", () => {
  const slide = slideOf([[raster("a")], [raster("b")]], [null, MOTION], { index: 1 });
  slide.states.forEach((s) => (s.index = 1));
  const specs = slideSpecs(slide, STYLE, { notice: () => {} });
  assert.equal(specs[0].layout, "title");
  assert.equal(specs[1].layout, "blank");
});

test("a reveal after a clip: a new slide opens in that state with a fade, and only later reveals are effects", () => {
  const states = [[raster("a")], [raster("b")], [raster("b"), box("00FF00")], [raster("b"), box("00FF00"), box("0000FF")]];
  const specs = slideSpecs(slideOf(states, [null, MOTION, REVEAL, REVEAL], { transition: "slide-left" }), STYLE, { notice: () => {} });
  assert.equal(specs.length, 3);
  assert.deepEqual(specs[0].transition, { push: "l" });
  assert.equal(specs[2].transition, "fade");
  assert.deepEqual(specs[2].shapes.filter((s) => s.effects?.length).map((s) => s.effects), [[{ click: 1, dir: "in" }]]);
});

test("two clips in a row are two slides with no transition between", () => {
  const specs = slideSpecs(slideOf([[raster("a")], [raster("b")], [raster("c")]], [null, MOTION, MOTION]), STYLE, { notice: () => {} });
  assert.deepEqual(specs.map((s) => s.shapes[0].kind), ["image", "video", "video"]);
  assert.deepEqual(specs.map((s) => s.transition), [undefined, undefined, undefined]);
});

test("a deck's own video on a slide with clicks stays a video with its own autoplay and loop", () => {
  const own = { kind: "video", rect: { x: 0, y: 0, w: 100, h: 100 }, natural: { w: 16, h: 9 }, data: Buffer.from("own"), ext: "mp4", posterImage: { data: Buffer.from("p"), ext: "png" }, autoplay: true, loop: true, muted: true };
  const [spec] = slideSpecs(slideOf([[own], [own, box("FF0000")]], [null, REVEAL]), STYLE, { notice: fail });
  const video = spec.shapes.find((s) => s.kind === "video");
  assert.equal(video.loop, true);
  assert.equal(video.autoplay, true);
  assert.ok(!video.effects?.length);
  assert.ok(spec.animated);
});

test("every spec is animated, so a deck's own video uses the form the web plays", () => {
  const specs = slideSpecs(slideOf([[box("FF0000")]]), STYLE, { notice: fail });
  assert.ok(specs[0].animated);
});

test("an unknown transition gets a notice and no transition; none says nothing", () => {
  const lines = [];
  const [spec] = slideSpecs(slideOf([[box("FF0000")]], [], { transition: "view-transition" }), STYLE, { notice: (l) => lines.push(l) });
  assert.equal(spec.transition, undefined);
  assert.equal(lines.length, 1);
  assert.match(lines[0], /^Notice: slide 3: transition "view-transition"/);
  slideSpecs(slideOf([[box("FF0000")]], [], { transition: "none" }), STYLE, { notice: fail });
});

test("what the default export reports is reported once per slide, however many states show it", () => {
  const lines = [];
  const missing = { kind: "missing", label: "img", src: "http://x/a.png" };
  const slide = slideOf([[raster("m"), missing], [raster("m"), missing, box("FF0000")]], [null, REVEAL]);
  slide.states[1].unsettled = true;
  slideSpecs(slide, STYLE, { notice: (l) => lines.push(l) });
  assert.deepEqual(lines.sort(), [
    "Notice: slide 3 was still changing when render.timeout ran out; exported as it stood then (raise render.timeout in pptx.config.mjs if it needs longer).",
    "Notice: slide 3: canvas exported as a picture (canvas).",
    "Notice: slide 3: image http://x/a.png didn't load - skipped.",
  ]);
});

test("the specs write a clean package: effects, the clip's timing and transitions in place", async () => {
  const states = [[table("a")], [table("b")], [raster("b")], [raster("b"), box("00FF00")]];
  const specs = slideSpecs(slideOf(states, [null, REVEAL, MOTION, REVEAL], { transition: "fade" }), STYLE, { notice: () => {} });
  assert.equal(specs.length, 3);
  const buffer = await writePptx(specs, DECK);
  assert.deepEqual(await lintPackage(buffer), []);
  const z = await unzip(buffer);
  const [one, clip, three] = [1, 2, 3].map((n) => z.text(`ppt/slides/slide${n}.xml`));
  assert.match(one, /<p:transition><p:fade\/><\/p:transition><p:timing>.*nodeType="clickEffect"/);
  assert.match(clip, /nodeType="afterEffect"/);
  assert.doesNotMatch(clip, /clickEffect|<p:transition/);
  assert.match(three, /<p:transition><p:fade\/><\/p:transition>/);
  assert.doesNotMatch(three, /clickEffect/);
});
