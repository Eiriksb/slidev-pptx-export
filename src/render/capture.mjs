// Render mode capture: drives Chromium through every slide of the deck's running Slidev dev
// server and returns one Scene per slide (see extract.js), plus the bytes the writer needs -
// image files, and PNG screenshots of elements drawn as pictures.

import { readFileSync } from "node:fs";
import { imageSize } from "image-size";

const EXTRACT = readFileSync(new URL("./extract.js", import.meta.url), "utf8");

export async function captureDeck({ baseUrl, chromium, timeout = 15000, screenshots = false, slidevVersion = "unknown" }) {
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage({ viewport: { width: 980, height: 552 }, deviceScaleFactor: 2 });
    page.setDefaultTimeout(timeout);
    await open(page, `${baseUrl}/1?print`);
    try {
      await page.waitForFunction(() => window.__slidev__?.nav?.total > 0 && !!document.querySelector(".slidev-page"));
    } catch {
      throw new Error(`unsupported Slidev version ${slidevVersion}: the page has no slide navigation data (window.__slidev__) or .slidev-page elements`);
    }
    const info = await page.evaluate(() => {
      const s = window.__slidev__;
      return {
        canvasWidth: s.configs.canvasWidth || 980,
        aspectRatio: s.configs.aspectRatio || 16 / 9,
        slides: s.nav.slides.map((r) => ({ no: r.no, sourceIndex: r.meta?.slide?.sourceIndex ?? r.no - 1, notes: r.meta?.slide?.note || null })),
      };
    });
    // Slidev's own slide height: Math.ceil(canvasWidth / aspectRatio) (980 x 16/9 -> 552).
    const canvas = { w: info.canvasWidth, h: Math.ceil(info.canvasWidth / info.aspectRatio) };
    await page.setViewportSize({ width: canvas.w, height: canvas.h });
    const slides = [];
    for (const s of info.slides) {
      try {
        slides.push(await captureSlide(page, baseUrl, s, canvas, screenshots));
      } catch (err) {
        slides.push({ index: s.no, sourceIndex: s.sourceIndex, error: String(err.message || err).split("\n")[0] });
      }
    }
    return { canvas, slides };
  } finally {
    await browser.close();
  }
}

async function captureSlide(page, baseUrl, s, canvas, screenshots) {
  await open(page, `${baseUrl}/${s.no}?print`);
  const sel = `.slidev-page-${s.no}`;
  await page.waitForFunction((pageSel) => {
    const p = document.querySelector(pageSel);
    return !!p && p.getBoundingClientRect().width > 0;
  }, sel);
  await page.evaluate(() => document.fonts.ready.then(() => true));
  await assertSlideLoaded(page, sel);
  await waitForStableCanvases(page);
  // Again, in case Vite reloaded the page while it settled.
  await hideViteOverlay(page);
  const screenshot = screenshots ? await page.locator(sel).screenshot({ animations: "disabled" }) : null;
  await page.addScriptTag({ content: EXTRACT });
  const scene = await page.evaluate(({ pageSel, size }) => window.__pptxExtract(document.querySelector(pageSel), size), { pageSel: sel, size: canvas });
  for (const node of scene.nodes) {
    if (node.kind === "image") await loadImage(page, node);
    else if (node.kind === "raster") node.png = await elementShot(page, node.id);
  }
  // Last: it hides the slide's contents.
  if (scene.background.raster) scene.background.png = await backgroundShot(page, sel);
  return { ...scene, index: s.no, sourceIndex: s.sourceIndex, notes: s.notes, screenshot };
}

// /<n>?print has every slide's resources in the DOM, so one image that never answers would keep
// "networkidle" from ever settling. Wait for the DOM, then give the network a few seconds to quiet
// down; the callers' own readiness checks decide whether the page is usable.
async function open(page, url) {
  await page.goto(url, { waitUntil: "domcontentloaded" });
  await page.waitForLoadState("networkidle", { timeout: 3000 }).catch(() => {});
  await hideViteOverlay(page);
}

// Every page load renders every slide, so one slide's compile error makes Vite draw its full-page
// error overlay (with the deck's file paths) over all of them. That slide fails on its own (see
// assertSlideLoaded); the overlay must not end up in the other slides' pictures.
async function hideViteOverlay(page) {
  const present = await page.evaluate(() => !!document.getElementById("pptx-hide-vite-overlay"));
  if (!present) await page.addStyleTag({ content: "vite-error-overlay{display:none!important}" }).then((h) => h.evaluate((el) => (el.id = "pptx-hide-vite-overlay")));
}

// When a slide's module fails (a missing asset, a template error, a broken layout) Slidev draws its
// SlideError layout instead, which has no stable class - only this text.
async function assertSlideLoaded(page, sel) {
  const failed = await page.evaluate((pageSel) => /An error occurred on this slide|Failed to fetch this slide/.test(document.querySelector(pageSel).textContent), sel);
  if (failed) throw new Error("Slidev could not load this slide (see its terminal output)");
}

// Charts animate in; wait until every canvas stops changing (up to ~5 s).
async function waitForStableCanvases(page) {
  let previous = null;
  for (let i = 0; i < 20; i++) {
    const sig = await page.evaluate(() =>
      [...document.querySelectorAll("canvas")]
        .map((c) => {
          try {
            const d = c.toDataURL();
            return `${c.width}x${c.height}:${d.length}:${d.slice(-48)}`;
          } catch {
            return "tainted";
          }
        })
        .join("|"),
    );
    if (sig === "" || sig === previous) return;
    previous = sig;
    await page.waitForTimeout(250);
  }
}

async function fetchBytes(page, src) {
  if (src.startsWith("data:")) {
    const comma = src.indexOf(",");
    const body = src.slice(comma + 1);
    return src.slice(0, comma).endsWith(";base64") ? Buffer.from(body, "base64") : Buffer.from(decodeURIComponent(body), "utf8");
  }
  const res = await page.request.get(src);
  if (!res.ok()) throw new Error(`HTTP ${res.status()}`);
  return Buffer.from(await res.body());
}

// The SVG file's own bytes drawn at the element's size (2x) into a transparent PNG, the fallback next
// to the SVG. Drawn from a data URL, not the page's request: no second fetch, no cross-origin taint,
// and the PNG always matches the SVG that is embedded.
async function svgPng(page, svg, rect) {
  const dataUrl = await page.evaluate(
    async ({ url, w, h }) => {
      const img = new Image();
      img.src = url;
      await img.decode();
      const c = document.createElement("canvas");
      c.width = Math.max(1, Math.round(w * 2));
      c.height = Math.max(1, Math.round(h * 2));
      c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
      return c.toDataURL("image/png");
    },
    { url: `data:image/svg+xml;base64,${svg.toString("base64")}`, w: rect.w, h: rect.h },
  );
  return Buffer.from(dataUrl.slice(dataUrl.indexOf(",") + 1), "base64");
}

// Image bytes for the writer. An <img> that can't be fetched is drawn as a picture; a CSS background
// (the extractor gives those no natural size) is dropped, since a picture of its container would
// bake the container's text in.
async function loadImage(page, node) {
  const isBackground = !node.natural;
  try {
    node.data = await fetchBytes(page, node.src);
    if (node.format === "svg") {
      const svg = node.data;
      node.data = await svgPng(page, svg, node.rect);
      node.svg = svg;
      node.ext = "png";
      // The extractor only lets an SVG through when its box has the SVG's aspect, so it fills the box.
      node.fit = "fill";
      return;
    }
    node.ext = node.format;
    if (!node.natural) {
      const d = imageSize(node.data);
      node.natural = { w: d.width, h: d.height };
    }
  } catch (err) {
    delete node.data;
    if (isBackground) {
      Object.assign(node, { kind: "missing", label: node.alt || node.src });
      return;
    }
    Object.assign(node, { kind: "raster", reason: `image not fetched: ${String(err.message || err).split("\n")[0]}`, label: node.alt || node.src, png: await elementShot(page, node.id) });
  }
}

async function elementShot(page, id) {
  try {
    return await page.locator(`[data-pptx-node="${id}"]`).screenshot({ animations: "disabled" });
  } catch {
    return null;
  }
}

// The slide's background alone: everything inside the layout hidden, then one screenshot.
async function backgroundShot(page, sel) {
  await page.evaluate((pageSel) => {
    const pageEl = document.querySelector(pageSel);
    const root = pageEl.querySelector(".slidev-layout") || pageEl.firstElementChild || pageEl;
    for (const c of root.children) c.style.setProperty("visibility", "hidden", "important");
  }, sel);
  return page.locator(sel).screenshot({ animations: "disabled" });
}
