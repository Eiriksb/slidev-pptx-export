// Render mode orchestration: start the deck's own Slidev, capture every slide in Chromium, and
// turn the scenes into SlideSpecs. A slide that couldn't be captured comes from the markdown
// (the static export) instead.

import { dirname, resolve } from "node:path";
import { createHash } from "node:crypto";
import { startDeckServer } from "./server.mjs";
import { loadChromium } from "./browser.mjs";
import { captureDeck } from "./capture.mjs";
import { sceneToSpec, shapesOf, unitsFor } from "./scene.mjs";
import { deckStyle, layoutOne, groupedBarChart } from "../layout.mjs";

// Config keys that only restate what the page shows; render mode reads the page instead.
const UNUSED_KEYS = ["logo", "eyebrowClass", "cssVarColors"];

// A chartResolvers entry for the chart's Vue component (or any component around it) overrides
// the data captured from the page - but only a "barchart"; other resolver kinds describe HTML
// the page already renders. A canvas exported as a picture (a chart that couldn't be drawn
// natively, or no readable chart at all) becomes the resolver's chart, in the canvas's content box.
function resolveChart(node, config, S) {
  const name = (node.components || []).find((n) => typeof config.chartResolvers?.[n] === "function");
  if (!name) return;
  try {
    const resolved = config.chartResolvers[name](config.resolvedDataDir);
    if (!resolved || resolved.kind !== "barchart") return;
    node.resolvedChart = groupedBarChart(resolved, S);
    if (node.kind === "raster") Object.assign(node, { kind: "chart", rect: node.chartRect || node.rect });
  } catch (err) {
    console.warn(`Warning: chart resolver ${name} failed (${err.message}) - using the chart drawn on the page.`);
  }
}

// What became of a raster node: a picture, or nothing when its screenshot failed.
const pictureNotice = (where, n) => (n.png ? `Notice: ${where}: ${n.label} exported as a picture (${n.reason}).` : `Notice: ${where}: ${n.label} skipped (its picture could not be taken).`);

const sha1 = (b) => createHash("sha1").update(b).digest("hex");

// What makes two nodes "the same element": kind, position (to 0.5 px), content and everything
// that changes how it is drawn. Titles never count - they differ per slide even when they sit
// in the same place.
export function nodeSignature(n) {
  if (!n.rect) return null;
  const r = [n.rect.x, n.rect.y, n.rect.w, n.rect.h].map((v) => Math.round(v * 2) / 2).join(",");
  const look = JSON.stringify([n.opacity ?? 1, n.rotate ?? 0]);
  if (n.kind === "image" || n.kind === "raster") {
    const bytes = n.png || n.data;
    return bytes ? `${n.kind}|${r}|${sha1(bytes)}|${n.svg ? sha1(n.svg) : ""}|${look}|${JSON.stringify([n.fit, n.crop, n.border, n.shadow])}` : null;
  }
  if (n.kind === "text" && n.role !== "title") return `text|${r}|${JSON.stringify(n.runs)}|${look}|${JSON.stringify([n.align, n.lineHeight, n.singleLine])}`;
  if (n.kind === "box") return `box|${r}|${JSON.stringify([n.fill, n.border, n.sides, n.radius, n.shadow])}|${look}`;
  return null;
}

// A signature per node (null for a node that can't be shared), numbered within the scene: two
// identical elements on one slide are two signatures, so one of them can be repeated while the
// other is not.
export function signaturesOf(nodes) {
  const seen = new Map();
  return nodes.map((n) => {
    const sig = nodeSignature(n);
    if (!sig) return null;
    const k = seen.get(sig) || 0;
    seen.set(sig, k + 1);
    return `${sig}#${k}`;
  });
}

// Signatures present on every scene (when there are at least two).
export function repeatedSignatures(scenes) {
  if (scenes.length < 2) return new Set();
  const counts = new Map();
  for (const s of scenes) {
    for (const sig of signaturesOf(s.nodes)) if (sig) counts.set(sig, (counts.get(sig) || 0) + 1);
  }
  return new Set([...counts].filter(([, n]) => n === scenes.length).map(([sig]) => sig));
}

const intersects = (a, b) => a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h;

// Whether a node draws anything: every kind does except a box with no fill, border, side bars or
// shadow. Even a translucent veil or a line of text changes how what lies beneath it looks.
function paints(n) {
  if (!n.rect) return false;
  if (n.kind !== "box") return true;
  return !!(n.fill || n.border || n.sides?.length || n.shadow);
}

// The master is drawn beneath every slide's own content, so an element the page draws above
// something slide-specific can't go there: a repeated signature is dropped when, on any scene, a
// node that draws something and is not repeated comes before it in DOM order and overlaps it. A
// dropped signature stays on the slides, so it can cover later ones - repeat until nothing changes.
export function uncoveredSignatures(scenes, repeated) {
  const kept = new Set(repeated);
  const sigs = scenes.map((s) => signaturesOf(s.nodes));
  for (let changed = true; changed; ) {
    changed = false;
    scenes.forEach((s, k) => {
      s.nodes.forEach((n, i) => {
        if (!kept.has(sigs[k][i])) return;
        const covered = s.nodes.slice(0, i).some((m, j) => !kept.has(sigs[k][j]) && paints(m) && intersects(m.rect, n.rect));
        if (covered) {
          kept.delete(sigs[k][i]);
          changed = true;
        }
      });
    });
  }
  return kept;
}

// A slide that couldn't be captured, from the markdown: its own slide laid out alone, numbered as
// the export's slide `no`. Slidev numbers a slide imported with `src:` within its own file, which
// the static parser doesn't read, so that one (or one with no markdown slide) is left blank.
function markdownFallback(scene, no, { inputPath, parsed, codeStore, config }) {
  const imported = scene.filepath && resolve(scene.filepath) !== resolve(inputPath);
  const slide = !imported && Number.isInteger(scene.sourceIndex) ? parsed[scene.sourceIndex] : null;
  const why = `Warning: slide ${no} could not be rendered (${scene.error})`;
  if (!slide) {
    console.warn(`${why} - left blank (${imported ? "no markdown fallback for imported slides" : "its markdown slide wasn't found"}).`);
    return [{ layout: "blank", placeholders: {}, shapes: [] }];
  }
  console.warn(`${why} - exported from markdown instead.`);
  return layoutOne(slide, no - 1, codeStore, config);
}

export async function renderSpecs({ inputPath, userConfig = {}, config, parsed, codeStore, screenshots = false, deps = {} }) {
  const { start = startDeckServer, chromiumFor = loadChromium, capture = captureDeck } = deps;
  const unused = UNUSED_KEYS.filter((k) => userConfig[k] != null);
  if (unused.length) console.warn(`Notice: render mode ignores ${unused.join(", ")} from pptx.config.mjs - the rendered page already shows them.`);

  // Playwright and its Chromium first: a missing package or browser fails before Slidev starts.
  const chromium = await chromiumFor(dirname(inputPath));
  let browser;
  try {
    browser = await chromium.launch();
  } catch (err) {
    throw new Error(`render mode could not launch Playwright's Chromium (${String(err.message || err).split("\n")[0]}) - run \`npx playwright install chromium\`.`, { cause: err });
  }
  let captured;
  try {
    const server = await start({ inputPath });
    try {
      captured = await capture({ baseUrl: server.url, browser, timeout: config.render?.timeout ?? 15000, screenshots, slidevVersion: server.version });
    } catch (err) {
      // A deck that fails to build or boot shows up here; Slidev's own output says why.
      const output = server.log();
      if (!output) throw err;
      throw new Error(`${err.message}\n\nSlidev output:\n${output}`, { cause: err });
    } finally {
      await server.stop();
    }
  } finally {
    await browser.close();
  }

  // A non-16:9 deck keeps its shape: the slide height follows Slidev's canvas.
  const { w, h } = captured.canvas;
  if (Math.abs(w / h - config.slide.width / config.slide.height) > 0.01) {
    config.slide = { ...config.slide, height: Math.round(((config.slide.width * h) / w) * 1000) / 1000 };
  }
  const style = deckStyle(config);
  const sceneStyle = { geometry: style.geometry, theme: style.theme, fonts: config.fonts };

  const rendered = captured.slides.filter((s) => !s.error);
  // Resolver charts first: a resolved chart is a chart, never a shared picture on the master.
  for (const s of rendered) {
    for (const node of s.nodes) if (node.kind === "chart" || node.kind === "raster") resolveChart(node, config, style);
  }

  // Elements every slide shares - and no slide content covers - go on the slide master once.
  const onMaster = uncoveredSignatures(rendered, repeatedSignatures(rendered));
  let masterShapes = [];
  if (onMaster.size) {
    const sigs = rendered.map((s) => signaturesOf(s.nodes));
    const first = rendered[0];
    const shared = first.nodes.filter((n, i) => onMaster.has(sigs[0][i]));
    for (const n of shared) if (n.kind === "raster") console.warn(pictureNotice("slide master", n));
    masterShapes = shared.flatMap((n) => shapesOf(n, unitsFor(first, style.geometry), sceneStyle));
    rendered.forEach((s, k) => {
      s.nodes = s.nodes.filter((n, i) => !onMaster.has(sigs[k][i]));
    });
  }

  const specs = [];
  const shots = [];
  for (const scene of captured.slides) {
    const no = specs.length + 1;
    if (scene.error) {
      const fallback = markdownFallback(scene, no, { inputPath, parsed, codeStore, config });
      specs.push(...fallback);
      shots.push(...fallback.map(() => null));
      continue;
    }
    for (const node of scene.nodes) {
      if (node.kind === "raster") console.warn(pictureNotice(`slide ${no}`, node));
      if (node.kind === "missing") console.warn(`Notice: slide ${no}: image ${node.src} didn't load - skipped.`);
    }
    specs.push(sceneToSpec(scene, sceneStyle));
    shots.push(scene.screenshot || null);
  }
  const translatedCharts = specs.reduce((n, s) => n + (s.shapes || []).filter((x) => x.kind === "chart").length, 0);
  return { specs, style, screenshots: shots, translatedCharts, masterShapes };
}
