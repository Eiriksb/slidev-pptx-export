// Render mode capture: drives Chromium through every slide of the deck's running Slidev dev
// server and returns one Scene per slide (see extract.js), plus the bytes the writer needs -
// image files, and PNG screenshots of elements drawn as pictures.

import { readFileSync } from "node:fs";
import { imageSize } from "image-size";

const EXTRACT = readFileSync(new URL("./extract.js", import.meta.url), "utf8");

// `browser`: a launched Playwright browser to use (left open), else one is launched from `chromium`
// (and closed).
export async function captureDeck({ baseUrl, browser: given = null, chromium, timeout = 15000, screenshots = false, slidevVersion = "unknown" }) {
  const browser = given ?? (await chromium.launch());
  let page = null;
  try {
    page = await browser.newPage({ viewport: { width: 980, height: 552 }, deviceScaleFactor: 2 });
    page.setDefaultTimeout(timeout);
    await open(page, `${baseUrl}/1?print`);
    try {
      await page.waitForFunction(() => window.__slidev__?.nav?.total > 0 && !!document.querySelector(".slidev-page"));
    } catch {
      throw new Error(await whyNoDeck(page, timeout, slidevVersion));
    }
    const info = await page.evaluate(() => {
      const s = window.__slidev__;
      return {
        canvasWidth: s.configs.canvasWidth || 980,
        aspectRatio: s.configs.aspectRatio || 16 / 9,
        // sourceIndex counts within the slide's own markdown file (filepath; an imported `src:` file
        // has its own).
        slides: s.nav.slides.map((r) => ({ no: r.no, sourceIndex: r.meta?.slide?.sourceIndex ?? null, filepath: r.meta?.slide?.filepath || null, notes: r.meta?.slide?.note || null })),
      };
    });
    // Slidev's own slide height: Math.ceil(canvasWidth / aspectRatio) (980 x 16/9 -> 552).
    const canvas = { w: info.canvasWidth, h: Math.ceil(info.canvasWidth / info.aspectRatio) };
    await page.setViewportSize({ width: canvas.w, height: canvas.h });
    const slides = [];
    for (const s of info.slides) {
      try {
        slides.push(await captureSlide(page, baseUrl, s, canvas, screenshots, timeout));
      } catch (err) {
        slides.push({ index: s.no, sourceIndex: s.sourceIndex, filepath: s.filepath, error: String(err.message || err).split("\n")[0] });
      }
    }
    return { canvas, slides };
  } finally {
    if (given) await page?.close();
    else await browser.close();
  }
}

// The first page never showed Slidev's slide data and slides. A page that drew nothing is a deck
// that didn't load in time (Slidev's first start builds its dependencies, which can take a while);
// one with Vite's error overlay didn't build; one that drew an app without them is a Slidev whose
// page this tool doesn't know.
async function whyNoDeck(page, timeout, slidevVersion) {
  const state = await page
    .evaluate(() => {
      const root = document.getElementById("app") || document.body;
      const drawn = !!root && [...root.children].some((e) => !/^(SCRIPT|STYLE|LINK|NOSCRIPT|TEMPLATE)$/.test(e.tagName));
      return document.querySelector("vite-error-overlay") ? "error" : drawn ? "drawn" : "blank";
    })
    .catch(() => "blank");
  if (state === "error") return "the deck failed to build (Vite reported an error)";
  if (state === "blank") return `the deck didn't load within ${timeout} ms (Slidev may still be starting up - raise render.timeout in pptx.config.mjs)`;
  return `unsupported Slidev version ${slidevVersion}: the page has no slide navigation data (window.__slidev__) or .slidev-page elements`;
}

async function captureSlide(page, baseUrl, s, canvas, screenshots, timeout) {
  await open(page, `${baseUrl}/${s.no}?print`);
  const sel = `.slidev-page-${s.no}`;
  await page.waitForFunction((pageSel) => {
    const p = document.querySelector(pageSel);
    return !!p && p.getBoundingClientRect().width > 0;
  }, sel);
  await page.evaluate(() => document.fonts.ready.then(() => true));
  await assertSlideLoaded(page, sel);
  const settled = await waitForQuietSlide(page, sel, timeout);
  await waitForStableCanvases(page);
  // Again, in case Vite reloaded the page while it settled.
  await hideViteOverlay(page);
  const screenshot = screenshots ? await page.locator(sel).screenshot({ animations: "disabled" }) : null;
  await videosLoaded(page, sel);
  await page.addScriptTag({ content: EXTRACT });
  const scene = await page.evaluate(({ pageSel, size }) => window.__pptxExtract(document.querySelector(pageSel), size), { pageSel: sel, size: canvas });
  const slide = { sel, canvas };
  // A picture shows what is behind its element, never what is drawn above it: the slide's
  // pictures are taken without the global top layer over them, the global bottom layer's without
  // the slide either. Those are exported on their own, above the picture.
  for (const [layer, above] of [["bottom", [sel, TOP_LAYER]], [undefined, [TOP_LAYER]], ["top", []]]) {
    const own = scene.nodes.filter((n) => n.layer === layer && (n.kind === "image" || n.kind === "raster" || n.kind === "video"));
    if (!own.length) continue;
    const shown = await hide(page, above);
    try {
      for (const node of own) {
        if (node.kind === "image") await loadImage(page, node, slide);
        else {
          // ... nor what the slide draws over it: that is exported on its own, above the picture.
          const uncovered = await hide(page, node.above || []);
          try {
            if (node.kind === "video") await loadVideo(page, node, slide);
            else node.png = await pictureOf(page, node, slide);
          } finally {
            await uncovered();
          }
        }
      }
    } finally {
      await shown();
    }
  }
  if (scene.background.raster) scene.background.png = await backgroundShot(page, sel);
  scene.nodes = withoutCoveredPictures(scene.nodes);
  return { ...scene, index: s.no, sourceIndex: s.sourceIndex, notes: s.notes, screenshot, ...(settled ? {} : { unsettled: true }) };
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

// How long a slide's DOM must stay unchanged to count as finished drawing, and how long after that
// it may wait for images still loading (one that never answers mustn't hold up the export).
const QUIET_MS = 500;
const IMAGE_GRACE_MS = 3000;

// A slide's own scripts may still be drawing when it has loaded: a map flying to its place and
// loading tiles, an animation run by requestAnimationFrame. Waits until the slide's DOM stops
// changing and its images have loaded. False when it was still changing after `limit` ms.
function waitForQuietSlide(page, sel, limit) {
  return page.evaluate(({ pageSel, quiet, grace, limit: cap }) => new Promise((done) => {
    const root = document.querySelector(pageSel);
    const pending = () => [...root.querySelectorAll("img")].filter((img) => !img.complete);
    // An image still loading on a page open this long has had open()'s wait already: it may never answer.
    const waitedFor = new Set(performance.now() > grace ? pending() : []);
    const loading = () => pending().some((img) => !waitedFor.has(img));
    let timer;
    let quietSince = null;
    const finish = (settled) => {
      observer.disconnect();
      clearTimeout(timer);
      clearTimeout(end);
      done(settled);
    };
    const check = () => {
      quietSince ??= performance.now();
      if (loading() && performance.now() - quietSince < grace) timer = setTimeout(check, 100);
      else finish(true);
    };
    const observer = new MutationObserver(() => {
      quietSince = null;
      clearTimeout(timer);
      timer = setTimeout(check, quiet);
    });
    observer.observe(root, { subtree: true, childList: true, attributes: true, characterData: true });
    timer = setTimeout(check, quiet);
    const end = setTimeout(() => finish(false), cap);
  }), { pageSel: sel, quiet: QUIET_MS, grace: IMAGE_GRACE_MS, limit });
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
async function loadImage(page, node, slide) {
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
    Object.assign(node, { kind: "raster", reason: `image not fetched: ${String(err.message || err).split("\n")[0]}`, label: node.alt || node.src, png: await elementShot(page, node, slide) });
  }
}

// Formats PowerPoint plays embedded (others stay pictures of their current frame).
const VIDEO_FORMATS = { mp4: "mp4", m4v: "mp4", mov: "mov" };

// Waits (up to 8 s) until the slide's videos know their size, so they can be exported as videos.
async function videosLoaded(page, sel) {
  await page.evaluate(async (pageSel) => {
    const vids = [...(document.querySelector(pageSel)?.querySelectorAll("video") || [])];
    await Promise.all(vids.map((v) => (v.readyState >= 1 ? null : new Promise((ok) => {
      v.addEventListener("loadedmetadata", ok, { once: true });
      v.addEventListener("error", ok, { once: true });
      setTimeout(ok, 8000);
    }))));
  }, sel);
}

// A video's file and its poster for the writer. An autoplaying video is seen playing, not its poster
// (which may even belong to another of its sources), so its poster is the file's own first frame that
// shows something; otherwise (or when there is none) the poster attribute's picture, else the frame the
// element shows. A video that can't be
// fetched, or that PowerPoint can't play, is a picture of its frame.
async function loadVideo(page, node, slide) {
  const ext = VIDEO_FORMATS[(node.src.split(/[?#]/)[0].split(".").pop() || "").toLowerCase()];
  try {
    if (!ext) throw new Error("not a format PowerPoint plays embedded");
    node.data = await fetchBytes(page, node.src);
    node.ext = ext;
    if (node.autoplay) node.posterImage = await firstFrame(page, node.src);
    if (node.poster && !node.posterImage) {
      try {
        const data = await fetchBytes(page, node.poster);
        const t = imageSize(data).type;
        if (t === "png" || t === "jpg" || t === "gif") node.posterImage = { data, ext: t };
      } catch {
        // the element's own frame below
      }
    }
    if (!node.posterImage) node.posterImage = { data: await elementShot(page, node, slide), ext: "png" };
  } catch (err) {
    delete node.data;
    Object.assign(node, { kind: "raster", reason: `video not embedded: ${String(err.message || err).split("\n")[0]}`, png: await elementShot(page, node, slide) });
  }
}

// The first frame of a video file that shows something, as a JPEG, drawn in the page (same origin,
// so the canvas can be read): tried at 0, 0.5, 1 and 2 s, each once the frame is decoded (a seek
// can report done before the frame is there, which draws an all-black picture). A frame of one flat
// colour (a fade-in, a black lead-in) is skipped; null when none shows anything within 15 s.
async function firstFrame(page, src) {
  const dataUrl = await page.evaluate(async (url) => {
    const v = document.createElement("video");
    v.muted = true;
    v.playsInline = true;
    v.preload = "auto";
    v.style.cssText = "position:fixed;left:0;top:0;width:64px;height:36px;opacity:0;pointer-events:none";
    document.body.appendChild(v);
    const wait = (ev, ms) => new Promise((ok, fail) => {
      v.addEventListener(ev, ok, { once: true });
      v.addEventListener("error", fail, { once: true });
      setTimeout(fail, ms);
    });
    // resolves once a frame at the current position has been decoded (or after `ms`)
    const decoded = (ms) => new Promise((ok) => {
      if (v.requestVideoFrameCallback) v.requestVideoFrameCallback(() => ok());
      setTimeout(ok, ms);
    });
    const draw = (w, h) => {
      const c = document.createElement("canvas");
      c.width = w;
      c.height = h;
      c.getContext("2d").drawImage(v, 0, 0, w, h);
      return c;
    };
    // whether a frame has more than one flat colour: the spread of its luminance on a small copy
    const shows = () => {
      const d = draw(32, 18).getContext("2d").getImageData(0, 0, 32, 18).data;
      let sum = 0, sq = 0, n = 0;
      for (let i = 0; i < d.length; i += 4, n++) {
        const y = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
        sum += y;
        sq += y * y;
      }
      return Math.sqrt(Math.max(0, sq / n - (sum / n) ** 2)) > 4;
    };
    try {
      v.src = url;
      await wait("loadeddata", 8000);
      for (const t of [0, 0.5, 1, 2]) {
        if (Number.isFinite(v.duration) && t > v.duration) break;
        const seeked = wait("seeked", 3000);
        const frame = decoded(1500);
        v.currentTime = t || 0.001;
        await seeked;
        await frame;
        if (shows()) {
          const c = draw(v.videoWidth, v.videoHeight);
          return c.toDataURL("image/jpeg", 0.9);
        }
      }
      return null;
    } catch {
      return null;
    } finally {
      v.removeAttribute("src");
      v.load();
      v.remove();
    }
  }, src);
  return dataUrl ? { data: Buffer.from(dataUrl.slice(dataUrl.indexOf(",") + 1), "base64"), ext: "jpg" } : null;
}

const TOP_LAYER = '[data-pptx-layer="top"]';

// Hides what `selectors` match (opacity, so nothing inside can show through) until the returned
// function is called.
async function hide(page, selectors) {
  if (!selectors.length) return async () => {};
  const style = await page.addStyleTag({ content: `${selectors.join(",")}{opacity:0!important}` });
  return () => style.evaluate((el) => el.remove());
}

// A raster node's picture: its element as drawn - or, for a background picture, without its contents.
function pictureOf(page, node, slide) {
  if (!node.contentsHidden) return elementShot(page, node, slide);
  return withoutContents(page, `[data-pptx-node="${node.id}"]`, () => elementShot(page, node, slide));
}

// A picture is an opaque screenshot of everything beneath it, so a picture wholly under a later one
// never shows and is already part of that one.
function withoutCoveredPictures(nodes) {
  const covers = (a, b) => a.x <= b.x && a.y <= b.y && a.x + a.w >= b.x + b.w && a.y + a.h >= b.y + b.h;
  return nodes.filter((n, i) => n.kind !== "raster" || !nodes.slice(i + 1).some((m) => m.kind === "raster" && m.png && m.layer === n.layer && covers(m.rect, n.rect)));
}

// The part of `rect` on the slide, when some of it runs past the slide's edges; else null.
export function partOnSlide(rect, canvas) {
  const x = Math.max(0, rect.x);
  const y = Math.max(0, rect.y);
  const right = Math.min(canvas.w, rect.x + rect.w);
  const bottom = Math.min(canvas.h, rect.y + rect.h);
  if (x === rect.x && y === rect.y && right === rect.x + rect.w && bottom === rect.y + rect.h) return null;
  return { x, y, w: right - x, h: bottom - y };
}

// An element running past the slide (Leaflet's overlay is larger than its map) is shot only as far as
// it is on the slide, the rest being invisible anyway: Chromium's capture of an element that doesn't
// fit the viewport draws the page at the wrong scale. The node then covers just that part.
async function elementShot(page, node, { sel, canvas }) {
  try {
    const part = partOnSlide(node.rect, canvas);
    if (!part) return await page.locator(`[data-pptx-node="${node.id}"]`).screenshot({ animations: "disabled" });
    const origin = await page.locator(sel).evaluate((el) => {
      el.scrollIntoView({ block: "start", inline: "start" });
      const r = el.getBoundingClientRect();
      return { x: r.left, y: r.top, width: r.width };
    });
    const scale = origin.width / canvas.w || 1;
    const clip = { x: origin.x + part.x * scale, y: origin.y + part.y * scale, width: part.w * scale, height: part.h * scale };
    const png = await page.screenshot({ clip, animations: "disabled" });
    node.rect = part;
    return png;
  } catch {
    return null;
  }
}

// Runs `shoot` with the contents of the element `sel` matches hidden - its children, and its own
// text made transparent (glyphs only: a border in currentColor stays) - and puts them back after.
async function withoutContents(page, sel, shoot) {
  const found = await page.evaluate((s) => {
    const el = document.querySelector(s);
    if (!el) return false;
    el.__pptxStyles = [el, ...el.children].map((e) => [e, e.getAttribute("style")]);
    for (const c of el.children) c.style.setProperty("visibility", "hidden", "important");
    el.style.setProperty("-webkit-text-fill-color", "transparent", "important");
    el.style.setProperty("-webkit-text-stroke-width", "0", "important");
    el.style.setProperty("text-shadow", "none", "important");
    el.style.setProperty("text-decoration-color", "transparent", "important");
    return true;
  }, sel);
  try {
    return await shoot();
  } finally {
    if (found) {
      await page.evaluate((s) => {
        const el = document.querySelector(s);
        for (const [e, style] of el.__pptxStyles) style == null ? e.removeAttribute("style") : e.setAttribute("style", style);
        delete el.__pptxStyles;
      }, sel);
    }
  }
}

// The slide's background alone: one screenshot of the page with the layout's contents and the
// global layers hidden.
async function backgroundShot(page, sel) {
  await page.evaluate((pageSel) => {
    const pageEl = document.querySelector(pageSel);
    (pageEl.querySelector(".slidev-layout") || pageEl.firstElementChild || pageEl).setAttribute("data-pptx-background", "");
  }, sel);
  const shown = await hide(page, ["[data-pptx-layer]"]);
  try {
    return await withoutContents(page, "[data-pptx-background]", () => page.locator(sel).screenshot({ animations: "disabled" }));
  } finally {
    await shown();
  }
}
