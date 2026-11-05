import { test } from "node:test";
import assert from "node:assert/strict";
import { createServer } from "node:http";
import { existsSync, mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { usableChromium } from "./helpers.mjs";
import { captureDeck } from "../../src/render/capture.mjs";
import { captureClickStates } from "../../src/render/click-states.mjs";

const chromium = await usableChromium();
const browser = chromium ? await chromium.launch() : null;
const opts = { skip: chromium ? false : "Playwright/Chromium not available", timeout: 120000 };
test.after(() => browser?.close());

// A white JPEG, standing in for recorded frames.
const WHITE_JPEG = browser ? await whiteJpeg() : null;

async function whiteJpeg() {
  const page = await browser.newPage({ viewport: { width: 16, height: 16 } });
  const jpeg = await page.screenshot({ type: "jpeg", quality: 92 });
  await page.close();
  return jpeg;
}
const FAKE_ENCODER = { kind: "ffmpeg" };
const fakeEncode = () => Buffer.from("mp4");

const slidePage = (n, inner) =>
  `<div class="slidev-page slidev-page-${n}" style="position:relative;width:980px;height:552px;overflow:hidden">` +
  `<div class="slidev-layout default" style="width:100%;height:100%;box-sizing:border-box;padding:40px">${inner}</div></div>`;

// What Slidev serves at /<n>?clicks=k (and /<n>?print, with `k` null), boiled down: `slides` is an array
// of { clicksTotal, inner(k), meta }, and only the page asked for draws its state; the slide after it is
// mounted too, as Slidev mounts neighbours.
function deckHtml(slides, n, k) {
  const data = {
    nav: { total: slides.length, clicksTotal: slides[n - 1].clicksTotal ?? 0, slides: slides.map((s, i) => ({ no: i + 1, meta: { slide: { sourceIndex: i }, ...s.meta } })) },
    configs: { canvasWidth: 980, aspectRatio: 16 / 9, transition: "fade" },
  };
  const neighbour = slides[n] ? slidePage(n + 1, slides[n].inner(0)) : "";
  return (
    `<!doctype html><html><head><style>body{margin:0;font-family:Arial,sans-serif}</style><script>window.__slidev__ = ${JSON.stringify(data)};</script></head>` +
    `<body>${slidePage(n, slides[n - 1].inner(k))}${neighbour}</body></html>`
  );
}

async function serve(t, slides) {
  const server = createServer((req, res) => {
    const url = new URL(req.url, "http://localhost");
    const m = url.pathname.match(/^\/(\d+)$/);
    if (!m) return void res.writeHead(404).end("not found");
    const k = url.searchParams.has("clicks") ? Number(url.searchParams.get("clicks")) : null;
    res.writeHead(200, { "content-type": "text/html" }).end(deckHtml(slides, Number(m[1]), k));
  });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  t.after(() => new Promise((resolve) => server.close(resolve)));
  return `http://127.0.0.1:${server.address().port}`;
}

const captureStates = async (t, slides, over = {}) =>
  captureClickStates({ baseUrl: await serve(t, slides), browser, timeout: 15000, record: async () => assert.fail("nothing should be recorded"), encoder: FAKE_ENCODER, encode: fakeEncode, notice: () => {}, ...over });

const para = (text) => `<p style="margin:0 0 8px;font:20px Arial">${text}</p>`;
const canvas = (style, color) => `<canvas width="100" height="100" style="position:absolute;${style}" data-fill="${color}"></canvas><script>document.currentScript.previousElementSibling.getContext("2d").fillStyle="${color}";document.currentScript.previousElementSibling.getContext("2d").fillRect(0,0,100,100);</script>`;
const FULL = "left:0;top:0;width:980px;height:552px";

const TWO_CLICKS = [
  { clicksTotal: 2, inner: (k) => para("One") + (k >= 1 ? para("Two") : "") + (k >= 2 ? para("Three") : "") },
  { inner: () => para("Neighbour") + para("Slide") },
];
const TRANSITIONS = [
  { meta: { transition: "slide-left | fade" }, inner: () => para("A") },
  { inner: () => para("B") },
];
const BIG_AND_SMALL_PICTURE = [{ inner: () => canvas(FULL, "#3366aa") + canvas("left:20px;top:20px;width:60px;height:60px", "#aa3333") }];
const pictureChanges = (to) => [{ clicksTotal: 1, inner: (k) => canvas(FULL, k === 0 || k === null ? "#808080" : to) }];
const TO_WHITE = pictureChanges("#ffffff");
const TO_BLUE = pictureChanges("#0000ff");
const CLICK_STATE_FAILS = [{ clicksTotal: 1, inner: (k) => (k === 1 ? "<div>An error occurred on this slide. Check the terminal.</div>" : para("Fine")) }];
const ALWAYS_FAILS = [{ clicksTotal: 1, inner: () => "<div>An error occurred on this slide. Check the terminal.</div>" }];

function fakeFrames(count = 2) {
  const dir = mkdtempSync(join(tmpdir(), "pptx-fake-frames-"));
  for (let i = 0; i < count; i++) writeFileSync(join(dir, `f${String(i).padStart(5, "0")}.jpg`), WHITE_JPEG);
  return dir;
}
const recording = (dir, over = {}) => async () => ({ dir, count: 2, fps: 30, motionFrames: 40, settled: true, changedVisibly: true, ...over });

test("every click state is captured, and neighbouring slides stay out", opts, async (t) => {
  const { slides, canvas: size } = await captureStates(t, TWO_CLICKS);
  assert.deepEqual(size, { w: 980, h: 552 });
  assert.equal(slides[0].states.length, 3);
  assert.deepEqual(slides[0].states.map((s) => s.nodes.length), [1, 2, 3]);
  assert.deepEqual(slides[0].clicks.slice(1).map((c) => c.kind), ["reveal", "reveal"]);
  assert.equal(slides[1].states.length, 1);
  assert.deepEqual(slides[1].clicks, []);
});

test("Slidev's click total is read through a ref too", opts, async (t) => {
  const slides = [{ ...TWO_CLICKS[0] }];
  const baseUrl = await serve(t, slides);
  const { slides: out } = await captureClickStates({
    baseUrl,
    browser: {
      newPage: async (o) => {
        const p = await browser.newPage(o);
        await p.addInitScript(() => {
          let nav;
          Object.defineProperty(window, "__slidev__", {
            configurable: true,
            get: () => nav,
            set: (v) => {
              nav = { ...v, nav: { ...v.nav, clicksTotal: { value: v.nav.clicksTotal } } };
            },
          });
        });
        return p;
      },
    },
    timeout: 15000,
    encoder: FAKE_ENCODER,
    notice: () => {},
  });
  assert.equal(out[0].states.length, 3);
});

test("a state's slide fields are what the writer's scene needs, and no screenshots are left", opts, async (t) => {
  const { slides } = await captureStates(t, TWO_CLICKS);
  const [s] = slides;
  assert.equal(s.index, 1);
  assert.equal(s.sourceIndex, 0);
  assert.equal(s.states[0].index, 1);
  assert.ok(s.states[0].canvas);
  assert.ok(s.states.every((state) => !state.screenshot));
});

test("the transition comes from the slide, else the deck", opts, async (t) => {
  const { slides } = await captureStates(t, TRANSITIONS);
  assert.deepEqual(slides.map((s) => s.transition), ["slide-left | fade", "fade"]);
});

test("in this mode a large picture is JPEG; a small one stays PNG; the default capture keeps PNG", opts, async (t) => {
  const { slides } = await captureStates(t, BIG_AND_SMALL_PICTURE);
  const [big, small] = slides[0].states[0].nodes.filter((n) => n.kind === "raster");
  assert.equal(big.png.subarray(0, 2).toString("hex"), "ffd8");
  assert.equal(big.pictureExt, "jpg");
  assert.equal(small.png.subarray(0, 4).toString("hex"), "89504e47");
  assert.equal(small.pictureExt, undefined);
  const plain = await captureDeck({ baseUrl: await serve(t, BIG_AND_SMALL_PICTURE), browser, timeout: 15000 });
  for (const n of plain.slides[0].nodes.filter((node) => node.kind === "raster")) {
    assert.equal(n.png.subarray(0, 4).toString("hex"), "89504e47");
    assert.equal(n.pictureExt, undefined);
  }
});

test("no recording happens when no picture changed", opts, async (t) => {
  let recorded = 0;
  await captureStates(t, TWO_CLICKS, { record: async () => { recorded++; } });
  assert.equal(recorded, 0);
});

test("a recording starts from the state before the click, on the slide's own page", opts, async (t) => {
  const dir = fakeFrames();
  const calls = [];
  await captureStates(t, TO_WHITE, { record: async (o) => { calls.push(o); return recording(dir)(); } });
  assert.equal(calls.length, 1);
  assert.match(calls[0].url, /\/1\?clicks=0$/);
  assert.equal(calls[0].sel, ".slidev-page-1");
  assert.deepEqual(calls[0].canvas, { w: 980, h: 552 });
  assert.equal(calls[0].browser, browser);
});

test("a motion click keeps its clip, with its first frame as the poster, and its frames are removed", opts, async (t) => {
  const dir = fakeFrames(60);
  const lines = [];
  const { slides } = await captureStates(t, TO_WHITE, { record: recording(dir, { count: 60 }), notice: (l) => lines.push(l) });
  const click = slides[0].clicks[1];
  assert.equal(click.kind, "motion");
  assert.equal(click.clip.data.toString(), "mp4");
  assert.equal(click.clip.duration, 2);
  assert.ok(click.clip.poster.equals(WHITE_JPEG));
  assert.deepEqual(click.clip.natural, { w: 1920, h: 1080 });
  assert.equal(existsSync(dir), false);
  assert.match(lines.join("\n"), /^Notice: slide 1 click 1: clip, 2\.0 s, 0\.0 MB\.$/m);
  assert.doesNotMatch(lines.join("\n"), /differs/);
});

test("a click with too little script motion is a reveal with no notice, and its frames are removed", opts, async (t) => {
  const dir = fakeFrames();
  const lines = [];
  const { slides } = await captureStates(t, TO_WHITE, { record: recording(dir, { motionFrames: 2 }), notice: (l) => lines.push(l) });
  assert.equal(slides[0].clicks[1].kind, "reveal");
  assert.deepEqual(lines, []);
  assert.equal(existsSync(dir), false);
});

for (const [name, over, reason] of [
  ["with no encoder", { encoder: null }, /no encoder.*install ffmpeg/],
  ["that never settles", { settled: false }, /never settled/],
  ["with no visible change", { changedVisibly: false }, /no visible change/],
]) {
  test(`a click ${name} becomes a reveal with a notice, and its frames are removed`, opts, async (t) => {
    const lines = [];
    const dir = fakeFrames();
    const { encoder = FAKE_ENCODER, ...fields } = over;
    const { slides } = await captureStates(t, TO_WHITE, { encoder, record: recording(dir, fields), notice: (l) => lines.push(l) });
    assert.equal(slides[0].clicks[1].kind, "reveal");
    assert.match(lines[0], new RegExp(`^Notice: slide 1 click 1: .*${reason.source}.*a Fade`));
    assert.equal(existsSync(dir), false);
  });
}

test("a recording or an encode that fails is a reveal with a notice, and the export carries on", opts, async (t) => {
  const lines = [];
  const boom = async () => {
    throw new Error("page crashed\nwith details");
  };
  const first = await captureStates(t, TO_WHITE, { record: boom, notice: (l) => lines.push(l) });
  assert.equal(first.slides[0].clicks[1].kind, "reveal");
  assert.match(lines[0], /^Notice: slide 1 click 1: .*page crashed.*a Fade/);
  assert.doesNotMatch(lines[0], /with details/);

  const dir = fakeFrames();
  const second = await captureStates(t, TO_WHITE, { record: recording(dir), encode: () => { throw new Error("no codec"); }, notice: (l) => lines.push(l) });
  assert.equal(second.slides[0].clicks[1].kind, "reveal");
  assert.match(lines[1], /no codec/);
  assert.equal(existsSync(dir), false);
});

test("a clip's last frame unlike the direct load is kept, with a notice", opts, async (t) => {
  const lines = [];
  const { slides } = await captureStates(t, TO_BLUE, { record: recording(fakeFrames()), notice: (l) => lines.push(l) });
  assert.equal(slides[0].clicks[1].kind, "motion");
  assert.match(lines.join("\n"), /^Notice: slide 1 click 1: the clip's last frame differs/m);
});

test("a slide whose click states cannot be captured exports its final state, with a notice", opts, async (t) => {
  const lines = [];
  const { slides } = await captureStates(t, CLICK_STATE_FAILS, { notice: (l) => lines.push(l) });
  assert.equal(slides[0].states.length, 1);
  assert.deepEqual(slides[0].clicks, []);
  assert.match(lines[0], /^Notice: slide 1: /);
});

test("a slide that can't be captured at all is a failed slide, with no notice from here", opts, async (t) => {
  const lines = [];
  const { slides } = await captureStates(t, ALWAYS_FAILS, { notice: (l) => lines.push(l) });
  assert.match(slides[0].error, /Slidev could not load this slide/);
  assert.equal(slides[0].sourceIndex, 0);
  assert.deepEqual(lines, []);
});
