import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, readdirSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { imageSize } from "image-size";
import { MOTION_FRAMES, clipFrame, clipViewport, recordClick } from "../../src/render/record.mjs";
import { usableChromium } from "./helpers.mjs";

const chromium = await usableChromium();
const browser = chromium ? await chromium.launch() : null;
const opts = { skip: chromium ? false : "Playwright/Chromium not available" };
test.after(() => browser?.close());

const page = ({ css = "", body = "", next = "", script = "" }) =>
  `<!doctype html><html><head><style>html,body{margin:0;background:#fff}${css}</style></head><body>` +
  `<div class="slidev-page slidev-page-1" style="position:absolute;inset:0">${body}</div>` +
  `<script>window.__slidev__={nav:{next(){${next}}}};${script}</script></body></html>`;

const ROOT_EL = "document.querySelector('.slidev-page-1')";
const appendBox = (rgb) => `const d=document.createElement('div');d.style='position:absolute;left:0;top:0;width:300px;height:100px;background:rgb(${rgb})';${ROOT_EL}.append(d);`;

// A red box moved by script for 1 s, with each frame slow to compute.
const ANIMATE_RAF_1S = page({
  body: '<div id="m" style="position:absolute;left:0;top:100px;width:100px;height:100px;background:#c00"></div>',
  next: "const start=performance.now();requestAnimationFrame(function f(now){for(let i=0;i<2e7;i++);const p=(now-start)/1000;m.style.left=Math.min(p,1)*800+'px';if(p<1)requestAnimationFrame(f);});",
});
const CSS_TRANSITION_500MS = page({
  css: ".b{position:absolute;left:0;top:0;width:200px;height:200px;background:#00c;opacity:0;transition:opacity .5s linear}.b.on{opacity:1}",
  body: '<div class="b"></div>',
  next: "document.querySelector('.b').classList.add('on');",
});
const GREEN = [0, 160, 0];
const CAPTION_APPEARS = page({ next: appendBox(GREEN.join(",")) });
const RAF_DRAWS_THE_SAME = page({ next: "const start=performance.now();requestAnimationFrame(function f(now){if(now-start<1000)requestAnimationFrame(f);});" });
const RAF_FOREVER = page({ next: "requestAnimationFrame(function f(){requestAnimationFrame(f);});" });
const ENDLESS_CSS_ANIMATION = page({
  css: "@keyframes p{from{opacity:.5}to{opacity:1}}.dot{position:absolute;left:500px;top:0;width:50px;height:50px;background:#888;animation:p 1s infinite}",
  body: '<div class="dot"></div>',
  next: appendBox(GREEN.join(",")),
});
const RED_SVG = '<svg xmlns="http://www.w3.org/2000/svg" width="300" height="300"><rect width="300" height="300" fill="#d00000"/></svg>';
const TILE_APPEARS = page({ next: `const i=new Image();i.src='/tile.svg';i.style='position:absolute;left:0;top:0';${ROOT_EL}.append(i);` });
const IMAGE_OUTSIDE_THE_SLIDE = page({ script: "const i=new Image();i.src='/tile.svg';document.body.append(i);", next: appendBox(GREEN.join(",")) });

// A red box moved by setInterval for 1 s: ten steps of 100 ms.
const ANIMATE_INTERVAL_1S = page({
  body: '<div id="m" style="position:absolute;left:0;top:100px;width:100px;height:100px;background:#c00"></div>',
  next: "let i=0;const id=setInterval(()=>{m.style.left=(++i*80)+'px';if(i>=10)clearInterval(id);},100);",
});
// Date() without new returns a string in a browser; a page may call it as the click starts.
const DATE_CALLED = page({ next: "window.__when = Date(); if (typeof window.__when !== 'string') throw new Error('not a string');" + appendBox(GREEN.join(",")) });

const frameName = (i) => `f${String(i).padStart(5, "0")}.jpg`;

async function recordStub(t, html, { tile, ...over } = {}) {
  const r = await recordClick({
    browser,
    url: "http://stub.test/1?clicks=0",
    sel: ".slidev-page-1",
    canvas: { w: 980, h: 552 },
    setup: async (context) => {
      await context.route("**/*", (route) => route.fulfill({ contentType: "text/html", body: html }));
      if (tile) await context.route("**/tile.svg", tile);
    },
    ...over,
  });
  t.after(() => rmSync(r.dir, { recursive: true, force: true }));
  return r;
}

// Whether a frame has a pixel of about this colour.
async function frameHasColor(dir, i, [r, g, b]) {
  const p = await browser.newPage();
  try {
    return await p.evaluate(
      async ({ src, r, g, b }) => {
        const img = new Image();
        img.src = src;
        await img.decode();
        const c = document.createElement("canvas");
        c.width = img.width;
        c.height = img.height;
        const x = c.getContext("2d");
        x.drawImage(img, 0, 0);
        const d = x.getImageData(0, 0, c.width, c.height).data;
        for (let k = 0; k < d.length; k += 4) if (Math.abs(d[k] - r) + Math.abs(d[k + 1] - g) + Math.abs(d[k + 2] - b) < 30) return true;
        return false;
      },
      { src: `data:image/jpeg;base64,${readFileSync(join(dir, frameName(i))).toString("base64")}`, r, g, b },
    );
  } finally {
    await p.close();
  }
}

test("a requestAnimationFrame animation is recorded at an exact frame rate, however slowly it paints", opts, async (t) => {
  const r = await recordStub(t, ANIMATE_RAF_1S);
  assert.ok(r.settled);
  assert.equal(r.fps, 30);
  assert.ok(r.motionFrames >= 28 && r.motionFrames <= 32, `motionFrames ${r.motionFrames}`);
  assert.ok(r.changedVisibly);
  assert.deepEqual(readdirSync(r.dir).sort(), Array.from({ length: r.count }, (_, i) => frameName(i)));
});

test("frames are the clip's size", opts, async (t) => {
  const r = await recordStub(t, CAPTION_APPEARS);
  const { width, height } = imageSize(readFileSync(join(r.dir, frameName(0))));
  assert.deepEqual({ width, height }, { width: 1920, height: 1080 });
});

// Slidev scales the slide to the window's width and centres it on black: at 1920 wide a 980x552 slide
// is 1081.47 tall, which leaves a band too thin to see in the page but a whole black row in a frame.
test("the black band Slidev leaves round the slide stays out of the frames", opts, async (t) => {
  const letterboxed =
    '<!doctype html><html><head><style>html,body{margin:0;height:100%;background:#000}' +
    ".slidev-page{position:absolute;left:0;top:50%;width:100vw;height:calc(100vw * 552 / 980);transform:translateY(-50%);background:#fff}</style></head>" +
    `<body><div class="slidev-page slidev-page-1"></div><script>window.__slidev__={nav:{next(){${appendBox(GREEN.join(","))}}}};</script></body></html>`;
  const r = await recordStub(t, letterboxed);
  const p = await browser.newPage();
  t.after(() => p.close());
  const edges = await p.evaluate(
    async (src) => {
      const img = new Image();
      img.src = src;
      await img.decode();
      const c = document.createElement("canvas");
      c.width = img.width;
      c.height = img.height;
      const x = c.getContext("2d");
      x.drawImage(img, 0, 0);
      return [0, img.height - 1].map((y) => x.getImageData(img.width / 2, y, 1, 1).data[0]);
    },
    `data:image/jpeg;base64,${readFileSync(join(r.dir, frameName(0))).toString("base64")}`,
  );
  for (const level of edges) assert.ok(level > 200, `edge row at level ${level}`);
});

test("an animation run by setInterval follows the stepped clock and counts as script motion", opts, async (t) => {
  const r = await recordStub(t, ANIMATE_INTERVAL_1S);
  assert.ok(r.settled);
  assert.ok(r.motionFrames >= 28 && r.motionFrames <= 32, `motionFrames ${r.motionFrames}`);
  assert.ok(r.changedVisibly);
});

test("a page that calls Date() without new records as it would play", opts, async (t) => {
  const r = await recordStub(t, DATE_CALLED);
  assert.ok(r.settled);
  assert.equal(await frameHasColor(r.dir, r.count - 1, GREEN), true);
});

test("a CSS transition after the click is drawn but is not script motion", opts, async (t) => {
  const r = await recordStub(t, CSS_TRANSITION_500MS);
  assert.ok(r.settled);
  assert.ok(r.motionFrames < MOTION_FRAMES, `motionFrames ${r.motionFrames}`);
  assert.ok(r.changedVisibly);
  assert.ok(r.count > 15, `count ${r.count}`);
});

test("the whole slide is in the frames, including a text that appears after the click", opts, async (t) => {
  const r = await recordStub(t, CAPTION_APPEARS);
  assert.equal(await frameHasColor(r.dir, 0, GREEN), false);
  assert.equal(await frameHasColor(r.dir, r.count - 1, GREEN), true);
});

test("a click that changes nothing reports no visible change", opts, async (t) => {
  const r = await recordStub(t, RAF_DRAWS_THE_SAME);
  assert.ok(r.settled);
  assert.ok(r.motionFrames >= MOTION_FRAMES);
  assert.equal(r.changedVisibly, false);
});

test("a page whose script never stops gives up at the cap", opts, async (t) => {
  const r = await recordStub(t, RAF_FOREVER, { maxSeconds: 1 });
  assert.equal(r.settled, false);
  assert.ok(r.count <= 31, `count ${r.count}`);
});

test("a CSS animation that never ends doesn't keep the page from settling", opts, async (t) => {
  const r = await recordStub(t, ENDLESS_CSS_ANIMATION, { maxSeconds: 3 });
  assert.ok(r.settled);
  assert.equal(await frameHasColor(r.dir, r.count - 1, GREEN), true);
});

test("an image still loading is waited for, so it is in the frames", opts, async (t) => {
  const tile = async (route) => {
    await new Promise((ok) => setTimeout(ok, 300));
    await route.fulfill({ contentType: "image/svg+xml", body: RED_SVG });
  };
  const r = await recordStub(t, TILE_APPEARS, { tile });
  assert.ok(r.settled);
  assert.equal(await frameHasColor(r.dir, r.count - 1, [208, 0, 0]), true);
});

test("an image that never answers doesn't hold the recording up", opts, async (t) => {
  const r = await recordStub(t, TILE_APPEARS, { tile: () => {}, imageGrace: 200, maxSeconds: 3 });
  assert.ok(r.settled);
});

test("an image outside the slide is not waited for", opts, async (t) => {
  const r = await recordStub(t, IMAGE_OUTSIDE_THE_SLIDE, { tile: () => {}, imageGrace: 60000, maxSeconds: 2 });
  assert.ok(r.settled);
});

test("a recording that fails leaves no files", opts, async (t) => {
  const tmpRoot = mkdtempSync(join(tmpdir(), "pptx-record-test-"));
  t.after(() => rmSync(tmpRoot, { recursive: true, force: true }));
  const broken = page({ script: "delete window.__slidev__;" });
  await assert.rejects(recordStub(t, broken, { tmpRoot }));
  assert.deepEqual(readdirSync(tmpRoot), []);
});

test("the window keeps a 4:3 deck's shape, with even dimensions", () => {
  assert.deepEqual(clipViewport({ w: 980, h: 552 }), { width: 1920, height: 1082 });
  assert.deepEqual(clipViewport({ w: 1024, h: 768 }), { width: 1920, height: 1440 });
  assert.deepEqual(clipViewport({ w: 1440, h: 810 }), { width: 1920, height: 1080 });
});

test("a clip's frame is the part of the window the slide fills, with even dimensions", () => {
  assert.deepEqual(clipFrame({ w: 980, h: 552 }), { x: 0, y: 1, width: 1920, height: 1080 });
  assert.deepEqual(clipFrame({ w: 1024, h: 768 }), { x: 0, y: 0, width: 1920, height: 1440 });
  assert.deepEqual(clipFrame({ w: 1440, h: 810 }), { x: 0, y: 0, width: 1920, height: 1080 });
});
