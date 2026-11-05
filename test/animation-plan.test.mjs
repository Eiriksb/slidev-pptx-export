import { test } from "node:test";
import assert from "node:assert/strict";
import { clickKind, effectsOf, mergeRuns, pictureChanged, planSlide, stateSignature } from "../src/render/animation-plan.mjs";

const box = (id, x = 0) => ({ kind: "box", id, rect: { x, y: 0, w: 10, h: 10 }, fill: { hex: id } });
const raster = (hash, w = 1000) => ({ kind: "raster", rect: { x: 0, y: 0, w, h: 500 }, png: Buffer.from(hash) });

test("a signature covers every kind, ignores ids and labels, and hashes bytes", () => {
  const table = { kind: "table", id: 1, rect: { x: 0, y: 0, w: 1, h: 1 }, rows: [["a"]], label: "x" };
  assert.equal(stateSignature(table), stateSignature({ ...table, id: 2, label: "y" }));
  assert.notEqual(stateSignature(table), stateSignature({ ...table, rows: [["b"]] }));
  assert.equal(stateSignature(raster("a")), stateSignature(raster("a")));
  assert.notEqual(stateSignature(raster("a")), stateSignature(raster("b")));
  assert.ok(stateSignature(raster("a")).length < 200, "bytes are hashed, not spelled out");
});

test("a signature sees bytes nested in a field, and ignores what the slide draws above", () => {
  const video = (data) => ({ kind: "video", rect: { x: 0, y: 0, w: 1, h: 1 }, posterImage: { data: Buffer.from(data), ext: "jpg" } });
  assert.notEqual(stateSignature(video("a")), stateSignature(video("b")));
  assert.equal(stateSignature({ ...raster("a"), above: ["x"] }), stateSignature({ ...raster("a"), above: ["y"] }));
});

test("a signature rounds position to half a pixel, and doesn't depend on the order of fields", () => {
  const at = (x) => ({ kind: "box", rect: { x, y: 0, w: 10, h: 10 } });
  assert.equal(stateSignature(at(10.1)), stateSignature(at(10.2)));
  assert.notEqual(stateSignature(at(10)), stateSignature(at(11)));
  assert.equal(stateSignature({ kind: "box", rect: at(1).rect, fill: 1 }), stateSignature({ fill: 1, rect: at(1).rect, kind: "box" }));
});

test("runs: what stays spans the states; what enters is inserted after its predecessor", () => {
  const runs = mergeRuns([[box("a"), box("c")], [box("a"), box("b"), box("c")]]);
  assert.deepEqual(runs.map((r) => [r.node.fill.hex, r.from, r.to]), [["a", 0, 1], ["b", 1, 1], ["c", 0, 1]]);
});

test("an element that hides and returns is two runs", () => {
  const runs = mergeRuns([[box("a")], [], [box("a")]]);
  assert.deepEqual(runs.map((r) => [r.from, r.to]).sort((x, y) => x[0] - y[0]), [[0, 0], [2, 2]]);
});

test("two identical elements are two runs, not one", () => {
  assert.equal(mergeRuns([[box("a"), box("a")]]).length, 2);
});

test("one of two identical elements leaving ends one run and keeps the other", () => {
  const runs = mergeRuns([[box("a"), box("a")], [box("a")]]);
  assert.deepEqual(runs.map((r) => [r.from, r.to]), [[0, 1], [0, 0]]);
});

test("when two states disagree on order, the later order wins by repeating the node", () => {
  const runs = mergeRuns([[box("a"), box("b")], [box("b"), box("a")]]);
  assert.equal(runs.length, 3);
  const inState1 = runs.filter((r) => r.from <= 1 && r.to >= 1).map((r) => r.node.fill.hex);
  assert.deepEqual(inState1, ["b", "a"]);
});

test("every state still paints in its own order", () => {
  const states = [[box("a"), box("b"), box("c")], [box("c"), box("d"), box("a")], [box("d"), box("b")]];
  const runs = mergeRuns(states);
  states.forEach((state, i) => {
    assert.deepEqual(runs.filter((r) => r.from <= i && r.to >= i).map((r) => r.node.fill.hex), state.map((n) => n.fill.hex));
  });
});

test("a changed picture is one whose signature was not there before", () => {
  assert.equal(pictureChanged([raster("old")], [raster("logo", 50), raster("new")]), true);
  assert.equal(pictureChanged([raster("same")], [raster("same")]), false);
  assert.equal(pictureChanged([raster("same")], [raster("same"), box("a")]), false);
  assert.equal(pictureChanged([raster("gone")], []), false);
});

test("click kinds", () => {
  const motion = { motionFrames: 40, settled: true, changedVisibly: true };
  assert.deepEqual(clickKind({ changed: false, record: null, encoder: {} }), { kind: "reveal", reason: null });
  assert.deepEqual(clickKind({ changed: true, record: { ...motion, motionFrames: 3 }, encoder: {} }), { kind: "reveal", reason: null });
  assert.deepEqual(clickKind({ changed: true, record: motion, encoder: {} }), { kind: "motion", reason: null });
  assert.equal(clickKind({ changed: true, record: motion, encoder: null }).kind, "reveal");
  assert.match(clickKind({ changed: true, record: motion, encoder: null }).reason, /no encoder/);
  assert.match(clickKind({ changed: true, record: { ...motion, settled: false }, encoder: {} }).reason, /never settled/);
  assert.match(clickKind({ changed: true, record: { ...motion, changedVisibly: false }, encoder: {} }).reason, /no visible change/);
});

const shape = (segs) => segs.map((s) => (s.kind === "clip" ? ["clip", s.click] : [s.first, s.last]));

test("a slide with no clicks is one still segment", () => {
  assert.deepEqual(shape(planSlide([[box("a")]], [null])), [[0, 0]]);
});

test("reveals stay on one still segment", () => {
  assert.deepEqual(shape(planSlide([[box("a")], [box("b")], [box("c")]], [null, "reveal", "reveal"])), [[0, 2]]);
});

test("a motion click is a clip of its own, and the next reveal opens a still slide in its own state", () => {
  const states = [[box("a")], [box("a"), box("b")], [raster("x")], [raster("x"), box("c")], [raster("x"), box("c"), box("d")]];
  assert.deepEqual(shape(planSlide(states, [null, "reveal", "motion", "reveal", "reveal"])), [[0, 1], ["clip", 2], [3, 4]]);
});

test("two motion clicks in a row are two clips, with no still slide between", () => {
  const states = [[raster("a")], [raster("b")], [raster("c")]];
  assert.deepEqual(shape(planSlide(states, [null, "motion", "motion"])), [[0, 0], ["clip", 1], ["clip", 2]]);
});

test("a slide that ends on a clip has nothing after it", () => {
  assert.deepEqual(shape(planSlide([[raster("a")], [raster("b")]], [null, "motion"])), [[0, 0], ["clip", 1]]);
});

test("a still slide's runs are numbered by the slide's states, and the state it opens in has no effect", () => {
  const states = [[raster("a")], [raster("b")], [raster("b"), box("c")], [raster("b"), box("c"), box("d")]];
  const still = planSlide(states, [null, "motion", "reveal", "reveal"])[2];
  assert.deepEqual([still.first, still.last], [2, 3]);
  assert.deepEqual(still.runs.map((r) => [r.from, r.to]), [[2, 3], [2, 3], [3, 3]]);
  assert.deepEqual(still.runs.map((r) => effectsOf(r, still)), [[], [], [{ click: 1, dir: "in" }]]);
});

test("effects are numbered within the segment", () => {
  const seg = { first: 2, last: 4 };
  assert.deepEqual(effectsOf({ from: 3, to: 3 }, seg), [{ click: 1, dir: "in" }, { click: 2, dir: "out" }]);
  assert.deepEqual(effectsOf({ from: 2, to: 4 }, seg), []);
});
