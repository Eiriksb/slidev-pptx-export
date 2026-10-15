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

function setup(md, { userConfig = {}, slides, canvas = { w: 980, h: 552 }, captureError = null, serverLog = null } = {}) {
  const config = mergeConfig(userConfig);
  const codeStore = [];
  const parsed = splitSlides(md).slides.map((s) => ({ ...parseSlide(s.content, codeStore, config), frontmatter: s.frontmatter }));
  const calls = { stopped: 0 };
  const deps = {
    chromiumFor: async () => ({ fake: true }),
    start: async () => ({
      url: "http://127.0.0.1:1",
      version: "52",
      log: () => serverLog ?? "",
      stop: async () => { calls.stopped++; },
    }),
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

test("a capture failure carries Slidev's own output, and the server is still stopped", async () => {
  const { args, calls } = setup("# A\n", { captureError: new Error("unsupported Slidev version x"), serverLog: "boom line" });
  await assert.rejects(renderSpecs(args), (err) => {
    assert.match(err.message, /unsupported Slidev version/);
    assert.match(err.message, /Slidev output:\nboom line/);
    return true;
  });
  assert.equal(calls.stopped, 1);
});

test("a capture failure with no Slidev output keeps its message as is", async () => {
  const { args } = setup("# A\n", { captureError: new Error("browser crashed"), serverLog: "" });
  await assert.rejects(renderSpecs(args), (err) => {
    assert.equal(err.message, "browser crashed");
    return true;
  });
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

test("screenshots line up with the specs: a slide's own capture, null for markdown fallbacks", async () => {
  const shot = Buffer.from("png");
  const { args } = setup("# A\n\n---\n\n# B\n", { slides: [{ ...scene(1, 0, "A"), screenshot: shot }, { index: 2, sourceIndex: 1, error: "boom" }] });
  const { result } = await withWarnings(() => renderSpecs({ ...args, screenshots: true }));
  assert.equal(result.specs.length, 2);
  assert.deepEqual(result.screenshots, [shot, null]);
});

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
