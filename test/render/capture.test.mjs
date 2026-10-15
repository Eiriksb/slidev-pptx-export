import { test } from "node:test";
import assert from "node:assert/strict";
import { createServer } from "node:http";
import { usableChromium } from "./helpers.mjs";
import { captureDeck } from "../../src/render/capture.mjs";

const chromium = await usableChromium();
const opts = { skip: chromium ? false : "Playwright/Chromium not available - run npm i -D playwright-chromium", timeout: 120000 };
const textOf = (node) => node.runs.map((r) => r.text).join("");

// What Slidev's /<n>?print serves, boiled down: every slide in the DOM at once, plus the
// window.__slidev__ data capture reads. Slide 1 has an image whose server never answers,
// slide 2 is Slidev's slide-error layout, slide 3 has a CSS background that 404s.
const page = (n, inner) =>
  `<div class="slidev-page slidev-page-${n}" style="position:relative;width:980px;height:552px;overflow:hidden">` +
  `<div class="slidev-layout default" style="width:100%;height:100%;box-sizing:border-box;padding:40px">${inner}</div></div>`;
const HTML =
  "<!doctype html><html><head><style>body{margin:0;font-family:Arial,sans-serif}</style><script>" +
  "window.__slidev__ = { nav: { total: 3, slides: [1, 2, 3].map((no) => ({ no, meta: { slide: { sourceIndex: no - 1 } } })) }, configs: { canvasWidth: 980, aspectRatio: 16 / 9 } };" +
  "</script></head><body>" +
  page(1, '<p style="margin:0;font:20px Arial">Hello there</p><img src="/hang.png" style="width:100px;height:50px">') +
  page(2, "<div>An error occurred on this slide. Check the terminal for more information.</div>") +
  page(3, '<div style="width:200px;height:100px;background-image:url(/404.png);background-size:cover;background-position:center">Inside</div>') +
  "</body></html>";

test("a hanging image doesn't stall the export, a slide error is a failed slide, an unfetchable background is skipped", opts, async (t) => {
  const sockets = new Set();
  const server = createServer((req, res) => {
    const path = new URL(req.url, "http://localhost").pathname;
    if (path === "/hang.png") return; // accepts the request and never answers
    if (/^\/\d+$/.test(path)) return void res.writeHead(200, { "content-type": "text/html" }).end(HTML);
    res.writeHead(404).end("not found");
  });
  server.on("connection", (socket) => {
    sockets.add(socket);
    socket.on("close", () => sockets.delete(socket));
  });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  t.after(async () => {
    for (const socket of sockets) socket.destroy();
    await new Promise((resolve) => server.close(resolve));
  });

  const started = Date.now();
  const { canvas, slides } = await captureDeck({ baseUrl: `http://127.0.0.1:${server.address().port}`, chromium, timeout: 15000 });
  // Every navigation waits out the 3 s network-idle cap while /hang.png stays pending: (3 slides + the
  // first load) x 3 s. The point is that it finishes, not that it raises the 15 s navigation timeout.
  assert.ok(Date.now() - started < 15000, `capture took ${Date.now() - started} ms`);
  assert.deepEqual(canvas, { w: 980, h: 552 });
  assert.equal(slides.length, 3);

  const [one, two, three] = slides;
  assert.equal(one.error, undefined);
  assert.ok(one.nodes.some((n) => n.kind === "text" && textOf(n) === "Hello there"));
  assert.ok(one.nodes.some((n) => n.kind === "missing" && n.src.endsWith("/hang.png")));

  assert.match(two.error, /could not load this slide/);
  assert.deepEqual([two.index, two.sourceIndex], [2, 1]);

  assert.equal(three.error, undefined);
  assert.deepEqual(three.nodes.filter((n) => n.kind === "raster"), []);
  assert.ok(three.nodes.some((n) => n.kind === "missing" && n.src.endsWith("/404.png")));
  assert.ok(three.nodes.some((n) => n.kind === "text" && textOf(n) === "Inside"));
});

test("an SVG image keeps its file next to a 2x PNG drawn from those bytes; broken ones are skipped", opts, async (t) => {
  // Chromium gives a viewBox-only SVG and a sizeless one a 300x150 natural size, so a 2:1 box suits both.
  const LOGO = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 40 20"><rect width="40" height="20" fill="#1a833b"/></svg>';
  const SIZED = '<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20"><rect width="40" height="20" fill="#1a833b"/></svg>';
  const FAR = '<svg xmlns="http://www.w3.org/2000/svg" width="30" height="30"><circle cx="15" cy="15" r="15" fill="#1a589f"/></svg>';
  const listen = (handler) =>
    new Promise((resolve) => {
      const server = createServer(handler);
      server.listen(0, "127.0.0.1", () => {
        t.after(() => new Promise((done) => server.close(done)));
        resolve(server.address().port);
      });
    });
  // A second origin that sends no CORS headers: drawn from the page it would taint the canvas.
  const farPort = await listen((req, res) => res.writeHead(200, { "content-type": "image/svg+xml" }).end(FAR));
  const html =
    "<!doctype html><html><head><style>body{margin:0}img{display:block}</style><script>" +
    "window.__slidev__ = { nav: { total: 1, slides: [{ no: 1, meta: { slide: { sourceIndex: 0 } } }] }, configs: { canvasWidth: 980, aspectRatio: 16 / 9 } };" +
    "</script></head><body>" +
    page(1,
      '<img src="/logo.svg" style="width:100px;height:50px"><img src="/sized.svg" style="width:80px;height:40px">' +
        `<img src="http://127.0.0.1:${farPort}/far.svg" style="width:30px;height:30px">` +
        '<img src="/gone.svg" style="width:60px;height:30px"><img src="/bad.svg" style="width:60px;height:30px"><img src="/empty.svg" style="width:60px;height:30px">') +
    "</body></html>";
  const files = { "/logo.svg": LOGO, "/sized.svg": SIZED, "/bad.svg": "this is not an svg", "/empty.svg": "" };
  const port = await listen((req, res) => {
    const path = new URL(req.url, "http://localhost").pathname;
    if (path === "/1") return void res.writeHead(200, { "content-type": "text/html" }).end(html);
    if (path in files) return void res.writeHead(200, { "content-type": "image/svg+xml" }).end(files[path]);
    res.writeHead(404).end("not found");
  });

  const { slides: [slide] } = await captureDeck({ baseUrl: `http://127.0.0.1:${port}`, chromium, timeout: 15000 });
  assert.equal(slide.error, undefined);
  const size = (png) => [png.readUInt32BE(16), png.readUInt32BE(20)];
  const [logo, sized, far, ...broken] = slide.nodes;
  assert.deepEqual([logo.kind, logo.format, logo.ext, logo.natural, logo.fit, logo.svg.toString()], ["image", "svg", "png", { w: 300, h: 150 }, "fill", LOGO]);
  assert.deepEqual(size(logo.data), [200, 100]);
  assert.deepEqual([sized.kind, sized.natural, sized.svg.toString()], ["image", { w: 40, h: 20 }, SIZED]);
  assert.deepEqual(size(sized.data), [160, 80]);
  assert.deepEqual([far.kind, far.svg.toString()], ["image", FAR]);
  assert.deepEqual(size(far.data), [60, 60]);
  assert.equal(far.data.subarray(0, 4).toString("hex"), "89504e47");
  assert.deepEqual(broken.map((n) => [n.kind, n.src.split("/").pop()]), [["missing", "gone.svg"], ["missing", "bad.svg"], ["missing", "empty.svg"]]);
});

test("an SVG that loaded in the page but whose fetched bytes won't draw is a picture, with a one-line reason", opts, async (t) => {
  // The page's own <img> request (Accept: image/...) gets a real SVG; capture's fetch (Accept: */*) gets garbage.
  const SVG = '<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20"><rect width="40" height="20" fill="#1a833b"/></svg>';
  const html =
    "<!doctype html><html><head><style>body{margin:0}img{display:block}</style><script>" +
    "window.__slidev__ = { nav: { total: 1, slides: [{ no: 1, meta: { slide: { sourceIndex: 0 } } }] }, configs: { canvasWidth: 980, aspectRatio: 16 / 9 } };" +
    "</script></head><body>" + page(1, '<img src="/flaky.svg" style="width:80px;height:40px">') + "</body></html>";
  const server = createServer((req, res) => {
    const path = new URL(req.url, "http://localhost").pathname;
    if (path === "/1") return void res.writeHead(200, { "content-type": "text/html" }).end(html);
    if (path === "/flaky.svg") return void res.writeHead(200, { "content-type": "image/svg+xml" }).end(/image\//.test(req.headers.accept || "") ? SVG : "<svg\nnot closed");
    res.writeHead(404).end("not found");
  });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  t.after(() => new Promise((resolve) => server.close(resolve)));
  const { slides: [slide] } = await captureDeck({ baseUrl: `http://127.0.0.1:${server.address().port}`, chromium, timeout: 15000 });
  const [node] = slide.nodes;
  assert.equal(node.kind, "raster");
  assert.match(node.reason, /^image not fetched: /);
  assert.ok(!node.reason.includes("\n"), node.reason);
  assert.equal(node.svg, undefined);
});
