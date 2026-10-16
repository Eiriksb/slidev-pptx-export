import { test } from "node:test";
import assert from "node:assert/strict";
import { renderSpecs, nodeSignature, repeatedSignatures, uncoveredSignatures } from "../src/render/index.mjs";
import { mergeConfig } from "../src/config.mjs";
import { splitSlides, parseSlide } from "../src/parse.mjs";

const C = (hex) => ({ hex, alpha: 1 });
const run = (text) => ({ text, font: "Inter", size: 36, bold: true, italic: false, color: C("1F2937"), spacing: 0, underline: false, strike: false, highlight: null });
const scene = (index, sourceIndex, title, nodes = []) => ({
  index, sourceIndex, layout: "default", canvas: { w: 980, h: 552 }, notes: "n", background: { color: C("FFFFFF") }, screenshot: null,
  nodes: [{ kind: "text", role: "title", rect: { x: 50, y: 50, w: 300, h: 40 }, align: "left", lineHeight: 40, singleLine: true, opacity: 1, runs: [run(title)] }, ...nodes],
});
const titleOf = (spec) => spec.placeholders.title.map((p) => p.runs.map((r) => r.text).join("")).join("");

function setup(md, { userConfig = {}, slides, canvas = { w: 980, h: 552 }, captureError = null, serverLog = null, launchError = null } = {}) {
  const config = mergeConfig(userConfig);
  const codeStore = [];
  const parsed = splitSlides(md).slides.map((s) => ({ ...parseSlide(s.content, codeStore, config), frontmatter: s.frontmatter }));
  const calls = { started: 0, stopped: 0, closed: 0 };
  const browser = { close: async () => { calls.closed++; } };
  const deps = {
    chromiumFor: async () => ({
      launch: async () => {
        if (launchError) throw launchError;
        return browser;
      },
    }),
    start: async () => {
      calls.started++;
      return { url: "http://127.0.0.1:1", version: "52", log: () => serverLog ?? "", stop: async () => { calls.stopped++; } };
    },
    capture: async ({ baseUrl, browser: given }) => {
      calls.baseUrl = baseUrl;
      calls.browserGiven = given === browser;
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

test("only a failed slide is laid out from the markdown, numbered by its place in the .pptx", async () => {
  // Both "Full" slides are too full for the static layout; only the one that failed is laid out, as slide 2.
  const full = (t) => `# ${t}\n\n` + Array.from({ length: 60 }, (_, i) => `Line ${i} of a paragraph long enough to wrap across the body of the slide at least once or twice.`).join("\n\n");
  const md = ["# Cover", full("Rendered"), "---\nhide: true\n---\n\n# Hidden", full("Failed")].join("\n\n---\n\n");
  const { args } = setup(md, { slides: [scene(1, 0, "Cover"), { index: 2, sourceIndex: 3, filepath: "/deck/slides.md", error: "boom" }, scene(3, 1, "Rendered")] });
  const { result, warnings } = await withWarnings(() => renderSpecs(args));
  assert.deepEqual(result.specs.map(titleOf), ["Cover", "Failed", "Rendered"]);
  assert.deepEqual(warnings, [
    "Warning: slide 2 could not be rendered (boom) - exported from markdown instead.",
    'Warning: slide 2 "Failed" is still too full after shrinking to 55% - content may run off the slide.',
  ]);
});

test("a failed slide with no markdown slide to fall back on is left blank, with a warning saying why", async () => {
  const md = "# A\n\n---\n\n# B\n";
  const failed = (extra) => ({ index: 2, sourceIndex: 1, error: "boom", ...extra });
  const cases = [
    [{ sourceIndex: null }, "left blank (its markdown slide wasn't found)"],
    [{ sourceIndex: 7 }, "left blank (its markdown slide wasn't found)"],
    [{ filepath: "/deck/pages/imported.md" }, "left blank (no markdown fallback for imported slides)"],
  ];
  for (const [extra, outcome] of cases) {
    const { args } = setup(md, { slides: [scene(1, 0, "A"), failed(extra), scene(3, 1, "B")] });
    const { result, warnings } = await withWarnings(() => renderSpecs(args));
    assert.equal(result.specs.length, 3, JSON.stringify(extra));
    assert.deepEqual(result.specs[1], { layout: "blank", placeholders: {}, shapes: [] });
    assert.deepEqual(result.screenshots, [null, null, null]);
    assert.deepEqual(warnings, [`Warning: slide 2 could not be rendered (boom) - ${outcome}.`]);
  }
  // The entry file itself (however Slidev spells the path) is a markdown fallback.
  const { args } = setup(md, { slides: [scene(1, 0, "A"), failed({ filepath: "/deck/./slides.md" })] });
  assert.deepEqual((await withWarnings(() => renderSpecs(args))).result.specs.map(titleOf), ["A", "B"]);
});

test("the server is stopped even when capture fails, and the error propagates", async () => {
  const { args, calls } = setup("# A\n", { captureError: new Error("browser crashed") });
  await assert.rejects(renderSpecs(args), /browser crashed/);
  assert.deepEqual([calls.stopped, calls.closed], [1, 1]);
});

test("Chromium is launched before Slidev starts: a browser that won't launch stops the export with no Slidev output", async () => {
  const { args, calls } = setup("# A\n", { launchError: new Error("browserType.launch: Executable doesn't exist at /x/chrome\nLooks like Playwright was just installed"), serverLog: "boom line" });
  await assert.rejects(renderSpecs(args), (err) => {
    assert.equal(err.message, "render mode could not launch Playwright's Chromium (browserType.launch: Executable doesn't exist at /x/chrome) - run `npx playwright install chromium`.");
    return true;
  });
  assert.equal(calls.started, 0);
  const ok = setup("# A\n", { slides: [scene(1, 0, "A")] });
  await withWarnings(() => renderSpecs(ok.args));
  assert.deepEqual([ok.calls.browserGiven, ok.calls.closed, ok.calls.stopped], [true, 1, 1], "capture gets the launched browser; it is closed after");
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

test("a picture that could not be taken is skipped with its own notice, not announced as a picture", async () => {
  const lost = { kind: "raster", id: 0, rect: { x: 0, y: 0, w: 10, h: 10 }, reason: "inline SVG", label: "svg.icon", alt: "", png: null };
  const { args } = setup("# A\n", { slides: [scene(1, 0, "A", [lost])] });
  const { result, warnings } = await withWarnings(() => renderSpecs(args));
  assert.deepEqual(warnings, ["Notice: slide 1: svg.icon skipped (its picture could not be taken)."]);
  assert.deepEqual(result.specs[0].shapes, []);
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

test("a resolver's malformed result warns and keeps the chart drawn on the page", async () => {
  const chartNode = {
    kind: "chart", id: 1, rect: { x: 40, y: 200, w: 900, h: 240 }, components: ["ExampleChart"], opacity: 1,
    chart: { type: "column", categories: ["Q1"], series: [{ name: "Page", values: [1], color: null }], min: null, max: null, stacked: false, legend: true, legendPos: "b", numFmt: null, gridColor: null, labelColor: null, legendColor: null },
  };
  const userConfig = { chartResolvers: { ExampleChart: () => ({ kind: "barchart", labels: ["From"] }) } };
  const { args } = setup("# A\n", { userConfig, slides: [scene(1, 0, "A", [chartNode])] });
  const { result, warnings } = await withWarnings(() => renderSpecs(args));
  assert.deepEqual(result.specs[0].shapes.find((s) => s.kind === "chart").chart.series.map((s) => s.name), ["Page"]);
  assert.equal(warnings.filter((w) => /^Warning: chart resolver ExampleChart failed \(/.test(w) && w.endsWith("using the chart drawn on the page.")).length, 1, warnings.join("\n"));
});

test("a chart exported as a picture is replaced by its component's chartResolvers chart, in the canvas's content box, with no picture notice", async () => {
  // The same pictured chart on both slides: resolved before the master is built, so it stays on the slides.
  const pictured = () => ({
    kind: "raster", id: 2, rect: { x: 40, y: 200, w: 900, h: 240 }, chartRect: { x: 50, y: 210, w: 880, h: 220 }, reason: "chart type line", label: "canvas",
    alt: "", components: ["ExampleChart", "Slide"], png: Buffer.from("png"),
  });
  const userConfig = { chartResolvers: { ExampleChart: () => ({ kind: "barchart", labels: ["From", "File"], datasets: [{ name: "Resolved", color: "1A833B", data: [5, 6] }], min: 0, max: 10 }) } };
  const { args } = setup("# A\n\n---\n\n# B\n", { userConfig, slides: [scene(1, 0, "A", [pictured()]), scene(2, 1, "B", [pictured()])] });
  const { result, warnings } = await withWarnings(() => renderSpecs(args));
  assert.deepEqual(warnings, []);
  assert.deepEqual(result.masterShapes, []);
  for (const spec of result.specs) {
    assert.deepEqual(spec.shapes.map((x) => x.kind), ["chart"]);
    const [shape] = spec.shapes;
    assert.deepEqual(shape.chart.categories, ["From", "File"]);
    assert.deepEqual([shape.x, shape.y, shape.w, shape.h], [0.6803, 2.8571, 11.9725, 2.9931]);
  }
  assert.equal(result.translatedCharts, 2);
});

test("a pictured chart with no resolver for its components stays a picture, with its notice", async () => {
  const pictured = { kind: "raster", id: 2, rect: { x: 40, y: 200, w: 900, h: 240 }, chartRect: { x: 40, y: 200, w: 900, h: 240 }, reason: "chart type line", label: "canvas", alt: "", components: ["OtherChart"], png: Buffer.from("png") };
  const userConfig = { chartResolvers: { ExampleChart: () => ({ kind: "barchart", labels: ["A"], datasets: [{ name: "S", data: [1] }] }) } };
  const { args } = setup("# A\n", { userConfig, slides: [scene(1, 0, "A", [pictured])] });
  const { result, warnings } = await withWarnings(() => renderSpecs(args));
  assert.deepEqual(warnings, ["Notice: slide 1: canvas exported as a picture (chart type line)."]);
  assert.deepEqual(result.specs[0].shapes.map((x) => x.kind), ["image"]);
  assert.equal(result.translatedCharts, 0);
});

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

test("a node's signature is its kind, rect (to 0.5 px) and content; titles and bytes-less images have none", () => {
  const img = (rect, data, svg) => ({ kind: "image", rect, data: Buffer.from(data), ...(svg ? { svg: Buffer.from(svg) } : {}) });
  const at = { x: 10, y: 20, w: 30, h: 40 };
  assert.equal(nodeSignature(img(at, "a")), nodeSignature(img({ x: 10.2, y: 19.8, w: 30, h: 40.1 }, "a")));
  assert.notEqual(nodeSignature(img(at, "a")), nodeSignature(img({ ...at, x: 11 }, "a")));
  assert.notEqual(nodeSignature(img(at, "a")), nodeSignature(img(at, "b")));
  assert.notEqual(nodeSignature(img(at, "a")), nodeSignature(img(at, "a", "<svg/>")));
  assert.equal(nodeSignature({ kind: "image", rect: at }), null);
  assert.equal(nodeSignature({ kind: "text", role: "title", rect: at, runs: [run("T")] }), null);
  assert.notEqual(nodeSignature({ kind: "text", rect: at, runs: [run("A")] }), nodeSignature({ kind: "text", rect: at, runs: [run("B")] }));
  assert.equal(nodeSignature({ kind: "missing", src: "x" }), null);
});

test("a signature is repeated only when every scene has it, and there are at least two scenes", () => {
  const note = (text) => ({ kind: "text", rect: { x: 800, y: 520, w: 100, h: 12 }, runs: [run(text)] });
  const sig = nodeSignature(note("Confidential"));
  assert.deepEqual([...repeatedSignatures([{ nodes: [note("Confidential")] }, { nodes: [note("Confidential"), note("Other")] }])], [`${sig}#0`]);
  assert.equal(repeatedSignatures([{ nodes: [note("Confidential")] }, { nodes: [note("Other")] }]).size, 0);
  assert.equal(repeatedSignatures([{ nodes: [note("Confidential")] }]).size, 0);
});

test("a repeated picture is announced once, for the slide master; a failed slide neither counts nor blocks it", async () => {
  const raster = () => ({ kind: "raster", id: 1, rect: { x: 10, y: 500, w: 40, h: 20 }, reason: "canvas", label: "canvas", alt: "", png: Buffer.from("png") });
  const { args } = setup("# A\n\n---\n\n# B\n\n---\n\n# C\n", { slides: [scene(1, 0, "A", [raster()]), { index: 2, sourceIndex: 1, error: "Timeout" }, scene(3, 2, "C", [raster()])] });
  const { result, warnings } = await withWarnings(() => renderSpecs(args));
  assert.deepEqual(warnings.filter((w) => w.startsWith("Notice")), ["Notice: slide master: canvas exported as a picture (canvas)."]);
  assert.equal(result.masterShapes.length, 1);
  assert.equal(result.masterShapes[0].kind, "image");
  assert.deepEqual(result.specs.map((s) => s.shapes.length), [0, 0, 0]);
});

const footerLogo = (extra = {}) => ({ kind: "image", id: 9, rect: { x: 48, y: 498, w: 95, h: 26 }, src: "/logo.png", format: "png", fit: "fill", natural: null, alt: "Logo", border: null, shadow: null, opacity: 1, data: Buffer.from("same"), ext: "png", ...extra });
const boxNode = (rect, fill, extra = {}) => ({ kind: "box", rect, fill, border: null, sides: null, radius: 0, shadow: null, opacity: 1, rotate: null, ...extra });
const BAND = { x: 0, y: 450, w: 980, h: 102 };

test("a logo the page draws above a slide's opaque band stays on the slides, not the master", async () => {
  const band = () => boxNode(BAND, C("1F2937"));
  const { args } = setup("# A\n\n---\n\n# B\n", { slides: [scene(1, 0, "A", [band(), footerLogo()]), scene(2, 1, "B", [footerLogo()])] });
  const { result } = await withWarnings(() => renderSpecs(args));
  assert.deepEqual(result.masterShapes, []);
  assert.deepEqual(result.specs.map((sp) => sp.shapes.map((x) => x.kind)), [["text", "image"], ["image"]]);
});

test("a band the page draws after the logo doesn't stop the logo going to the master", async () => {
  const { args } = setup("# A\n\n---\n\n# B\n", { slides: [scene(1, 0, "A", [footerLogo(), boxNode(BAND, C("1F2937"))]), scene(2, 1, "B", [footerLogo()])] });
  const { result } = await withWarnings(() => renderSpecs(args));
  assert.deepEqual(result.masterShapes.map((x) => x.alt), ["Logo"]);
  assert.deepEqual(result.specs.map((sp) => sp.shapes.length), [1, 0]);
});

test("anything visible drawn before the logo keeps it off the master: a translucent veil, a faded box, a side bar, text", async () => {
  const veil = boxNode(BAND, { hex: "1F2937", alpha: 0.3 });
  const faded = boxNode(BAND, C("1F2937"), { opacity: 0.5 });
  const bar = boxNode(BAND, null, { sides: [{ side: "t", width: 2, color: C("1A833B") }] });
  const outline = boxNode(BAND, null, { border: { width: 1, color: C("1A833B") } });
  const shadow = boxNode(BAND, null, { shadow: { x: 0, y: 2, blur: 4, color: C("000000") } });
  const caption = { kind: "text", role: "body", rect: { x: 40, y: 500, w: 200, h: 20 }, align: "left", lineHeight: 20, singleLine: true, opacity: 1, runs: [run("Caption")] };
  for (const before of [veil, faded, bar, outline, shadow, caption]) {
    const { args } = setup("# A\n\n---\n\n# B\n", { slides: [scene(1, 0, "A", [before, footerLogo()]), scene(2, 1, "B", [footerLogo()])] });
    const { result } = await withWarnings(() => renderSpecs(args));
    assert.deepEqual(result.masterShapes, [], JSON.stringify(before));
  }
  // A box that draws nothing doesn't count.
  const empty = boxNode(BAND, null);
  const { args } = setup("# A\n\n---\n\n# B\n", { slides: [scene(1, 0, "A", [empty, footerLogo()]), scene(2, 1, "B", [footerLogo()])] });
  assert.deepEqual((await withWarnings(() => renderSpecs(args))).result.masterShapes.map((x) => x.alt), ["Logo"]);
});

test("a repeated element that a dropped repeated one would cover is dropped too", async () => {
  const hero = boxNode({ x: 0, y: 0, w: 980, h: 552 }, C("0B3D91"));
  const picture = { kind: "raster", id: 3, rect: { x: 0, y: 0, w: 980, h: 552 }, reason: "canvas", label: "canvas", alt: "", png: Buffer.from("own") };
  const { args } = setup("# A\n\n---\n\n# B\n", { slides: [scene(1, 0, "A", [picture, hero, footerLogo()]), scene(2, 1, "B", [hero, footerLogo()])] });
  const { result } = await withWarnings(() => renderSpecs(args));
  assert.deepEqual(result.masterShapes, []);
  assert.deepEqual(result.specs.map((sp) => sp.shapes.length), [3, 2]);
});

test("uncoveredSignatures keeps a repeated signature only when nothing painted and unshared precedes it on any scene", () => {
  const scenes = [{ nodes: [boxNode(BAND, C("111111")), footerLogo()] }, { nodes: [footerLogo()] }];
  const repeated = repeatedSignatures(scenes);
  assert.equal(repeated.size, 1);
  assert.equal(uncoveredSignatures(scenes, repeated).size, 0);
  const apart = [{ nodes: [boxNode({ x: 0, y: 0, w: 980, h: 100 }, C("111111")), footerLogo()] }, { nodes: [footerLogo()] }];
  assert.equal(uncoveredSignatures(apart, repeatedSignatures(apart)).size, 1);
  const code = [{ nodes: [{ kind: "code", rect: BAND, lines: [] }, footerLogo()] }, { nodes: [footerLogo()] }];
  assert.equal(uncoveredSignatures(code, repeatedSignatures(code)).size, 0);
});

test("two identical elements on one slide and one on another: only one goes to the master", async () => {
  const card = () => boxNode({ x: 100, y: 200, w: 200, h: 100 }, C("E5F1E8"));
  const { args } = setup("# A\n\n---\n\n# B\n", { slides: [scene(1, 0, "A", [card(), card()]), scene(2, 1, "B", [card()])] });
  const { result } = await withWarnings(() => renderSpecs(args));
  assert.equal(result.masterShapes.length, 1);
  assert.deepEqual(result.specs.map((sp) => sp.shapes.length), [1, 0]);
});

test("the same element at another opacity, rotation, fit or alignment is not the same element", async () => {
  const at = { x: 10, y: 20, w: 30, h: 40 };
  const base = footerLogo({ rect: at });
  for (const change of [{ opacity: 0.5 }, { rotate: 5 }, { fit: "cover" }, { border: { width: 1, color: C("000000") } }, { shadow: { x: 1, y: 1, blur: 2, color: C("000000") } }]) {
    assert.notEqual(nodeSignature(base), nodeSignature({ ...base, ...change }), JSON.stringify(change));
  }
  const text = { kind: "text", rect: at, runs: [run("A")], align: "left", lineHeight: 20, singleLine: true, opacity: 1 };
  for (const change of [{ opacity: 0.5 }, { rotate: 5 }, { align: "center" }, { lineHeight: 24 }, { singleLine: false }]) {
    assert.notEqual(nodeSignature(text), nodeSignature({ ...text, ...change }), JSON.stringify(change));
  }
  const box = boxNode(at, C("112233"));
  for (const change of [{ opacity: 0.5 }, { rotate: 5 }]) assert.notEqual(nodeSignature(box), nodeSignature({ ...box, ...change }), JSON.stringify(change));
  const { args } = setup("# A\n\n---\n\n# B\n", { slides: [scene(1, 0, "A", [footerLogo()]), scene(2, 1, "B", [footerLogo({ opacity: 0.5 })])] });
  assert.deepEqual((await withWarnings(() => renderSpecs(args))).result.masterShapes, []);
});
