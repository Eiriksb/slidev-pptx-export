// Render mode orchestration: start the deck's own Slidev, capture every slide in Chromium, and
// turn the scenes into SlideSpecs. A slide that couldn't be captured comes from the markdown
// (the static export) instead.

import { dirname } from "node:path";
import { startDeckServer } from "./server.mjs";
import { loadChromium } from "./browser.mjs";
import { captureDeck } from "./capture.mjs";
import { sceneToSpec } from "./scene.mjs";
import { deckStyle, layoutEach, groupedBarChart } from "../layout.mjs";

// Config keys that only restate what the page shows; render mode reads the page instead.
const UNUSED_KEYS = ["logo", "eyebrowClass", "cssVarColors"];

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
  return { specs, style, screenshots: shots, translatedCharts };
}
