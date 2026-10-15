// Render mode orchestration: start the deck's own Slidev, capture every slide in Chromium, and
// turn the scenes into SlideSpecs. A slide that couldn't be captured comes from the markdown
// (the static export) instead.

import { dirname } from "node:path";
import { createHash } from "node:crypto";
import { startDeckServer } from "./server.mjs";
import { loadChromium } from "./browser.mjs";
import { captureDeck } from "./capture.mjs";
import { sceneToSpec, shapesOf, unitsFor } from "./scene.mjs";
import { deckStyle, layoutEach, groupedBarChart } from "../layout.mjs";

// Config keys that only restate what the page shows; render mode reads the page instead.
const UNUSED_KEYS = ["logo", "eyebrowClass", "cssVarColors"];

// A chartResolvers entry for the chart's Vue component (or any component around it) overrides
// the data captured from the page - but only a "barchart"; other resolver kinds describe HTML
// the page already renders.
function resolveChart(node, config, S) {
  const name = node.components.find((n) => typeof config.chartResolvers?.[n] === "function");
  if (!name) return;
  try {
    const resolved = config.chartResolvers[name](config.resolvedDataDir);
    if (resolved && resolved.kind === "barchart") node.resolvedChart = groupedBarChart(resolved, S);
  } catch (err) {
    console.warn(`Warning: chart resolver ${name} failed (${err.message}) - using the chart drawn on the page.`);
  }
}

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
  } catch (err) {
    // A deck that never boots only shows up as a capture timeout; Slidev's own output says why.
    const output = server.log();
    if (!output) throw err;
    throw new Error(`${err.message}\n\nSlidev output:\n${output}`, { cause: err });
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
  // Laid out from the markdown only when a slide needs it, so a clean render prints none of the
  // static layout's warnings.
  let fromMarkdown = null;

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

  const specs = [];
  const shots = [];
  for (const scene of captured.slides) {
    const no = specs.length + 1;
    if (scene.error) {
      console.warn(`Warning: slide ${no} could not be rendered (${scene.error}) - exported from markdown instead.`);
      fromMarkdown ??= layoutEach(parsed, codeStore, config);
      const fallback = fromMarkdown[scene.sourceIndex] || [];
      specs.push(...fallback);
      shots.push(...fallback.map(() => null));
      continue;
    }
    for (const node of scene.nodes) {
      if (node.kind === "chart") resolveChart(node, config, style);
      if (node.kind === "raster") console.warn(`Notice: slide ${no}: ${node.label} exported as a picture (${node.reason}).`);
      if (node.kind === "missing") console.warn(`Notice: slide ${no}: image ${node.src} didn't load - skipped.`);
    }
    specs.push(sceneToSpec(scene, sceneStyle));
    shots.push(scene.screenshot || null);
  }
  const translatedCharts = specs.reduce((n, s) => n + (s.shapes || []).filter((x) => x.kind === "chart").length, 0);
  return { specs, style, screenshots: shots, translatedCharts, masterShapes };
}
