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
  page(3, '<div style="width:200px;height:100px;background-image:url(/404.png);background-size:cover">Inside</div>') +
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
