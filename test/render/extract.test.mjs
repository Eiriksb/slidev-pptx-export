import { test, after } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { usableChromium, extractHtml } from "./helpers.mjs";

const chromium = await usableChromium();
const browser = chromium ? await chromium.launch() : null;
after(() => browser?.close());
const opts = { skip: browser ? false : "Playwright/Chromium not available - run npm i -D playwright-chromium" };
const PNG = `data:image/png;base64,${readFileSync(new URL("../fixtures/kitchen-sink/public/pic.png", import.meta.url)).toString("base64")}`;
const textOf = (node) => node.runs.map((r) => r.text).join("");

test("a card becomes a rounded box with fill and border; its heading and paragraph become text nodes", opts, async () => {
  const s = await extractHtml(browser,
    '<div style="width:276px;padding:20px;border-radius:8px;border:1px solid #D9E3DB;background:#EEF4F0">' +
      '<h3 style="margin:0;font:24px Arial">First point</h3>' +
      '<p style="margin:8px 0 0;font:17.6px Arial;color:#1f2937">Each card is a <code style="background:#f5f5f5;font-family:monospace">kv-card</code> inside a grid row of three cards.</p></div>');
  const box = s.nodes.find((n) => n.kind === "box");
  assert.deepEqual({ fill: box.fill, border: box.border, radius: box.radius, x: box.rect.x, y: box.rect.y, w: box.rect.w }, {
    fill: { hex: "EEF4F0", alpha: 1 }, border: { width: 1, color: { hex: "D9E3DB", alpha: 1 } }, radius: 8, x: 40, y: 40, w: 318,
  });
  const texts = s.nodes.filter((n) => n.kind === "text");
  assert.deepEqual(texts.map(textOf), ["First point", "Each card is a kv-card inside a grid row of three cards."]);
  assert.ok(s.nodes.indexOf(box) < s.nodes.indexOf(texts[0]), "the box is painted before its contents");
  assert.equal(texts[0].singleLine, true);
  assert.ok(texts[0].rect.w < 276, "a single line hugs its text");
  assert.equal(texts[1].singleLine, false);
  assert.deepEqual({ x: texts[1].rect.x, w: texts[1].rect.w }, { x: 61, w: 276 }, "wrapped text keeps its content box width");
  const code = texts[1].runs.find((r) => r.text === "kv-card");
  assert.deepEqual(code.highlight, { hex: "F5F5F5", alpha: 1 });
  assert.equal(code.font, "monospace");
});

test("an eyebrow keeps uppercase, letter spacing and weight in its run", opts, async () => {
  const s = await extractHtml(browser, '<div style="font:800 12px Arial;letter-spacing:1.68px;text-transform:uppercase;color:#1a833b">Overview</div>');
  assert.deepEqual(s.nodes.map((n) => n.kind), ["text"]);
  assert.deepEqual(s.nodes[0].runs, [{ text: "OVERVIEW", font: "Arial", size: 12, bold: true, italic: false, color: { hex: "1A833B", alpha: 1 }, spacing: 1.68, underline: false, strike: false, highlight: null }]);
});

test("the first h1 is the title; other headings are body text", opts, async () => {
  const s = await extractHtml(browser, '<h1 style="margin:0">Card <em>row</em></h1><h2>Sub</h2><h1>Second</h1>');
  assert.deepEqual(s.nodes.map((n) => [n.role, textOf(n)]), [["title", "Card row"], ["body", "Sub"], ["body", "Second"]]);
  assert.equal(s.nodes[0].runs[1].italic, true);
});

test("lists: items, nesting levels, ordered start, marker color and measured indent", opts, async () => {
  const s = await extractHtml(browser,
    '<ul style="margin:0;padding-left:24px;font:16px/24px Arial"><li>One</li><li>Two<ol start="3" style="padding-left:24px"><li>Three</li></ol></li></ul>',
    { css: "li::marker{color:#1a833b}" });
  const list = s.nodes.find((n) => n.kind === "list");
  assert.deepEqual(list.items.map((i) => [i.level, i.marker, i.start ?? null, textOf(i), i.markerColor.hex, i.indent, i.lineHeight]), [
    [0, "bullet", null, "One", "1A833B", 24, 24],
    [0, "bullet", null, "Two", "1A833B", 24, 24],
    [1, "number", 3, "Three", "1A833B", 48, 24],
  ]);
  assert.equal(s.nodes.filter((n) => n.kind === "text").length, 0, "list text lives only in the list node");
});

test("svg, canvas and video become raster nodes tagged for screenshots, without walking their contents", opts, async () => {
  const s = await extractHtml(browser,
    '<svg width="50" height="50"><text x="0" y="20">svg text</text></svg><canvas width="40" height="40" style="display:block"></canvas><video width="40" height="30" style="display:block"></video>');
  assert.deepEqual(s.nodes.map((n) => [n.kind, n.reason, n.label]), [["raster", "inline SVG", "svg"], ["raster", "canvas", "canvas"], ["raster", "video", "video"]]);
  assert.deepEqual(s.nodes.map((n) => n.id), [0, 1, 2]);
});

test("inline padding does not change how lines are counted or where the text sits", opts, async () => {
  const para = (codeStyle) => `<p style="margin:0;width:200px;font:17.6px/normal Arial">Some words before <code style="${codeStyle}">code</code> and more words after it that wrap</p>`;
  const padded = (await extractHtml(browser, para("padding:.2em .4em;background:#f5f5f5"))).nodes.find((n) => n.kind === "text");
  const plain = (await extractHtml(browser, para("background:#f5f5f5"))).nodes.find((n) => n.kind === "text");
  assert.equal(plain.singleLine, false, "the paragraph wraps");
  assert.equal(padded.singleLine, false, "the padded code's box must not merge two lines into one");
  assert.ok(Math.abs(padded.rect.y - plain.rect.y) < 0.5, `y ${padded.rect.y} vs ${plain.rect.y}`);
  assert.ok(Math.abs(padded.rect.h - plain.rect.h) < 0.5, `h ${padded.rect.h} vs ${plain.rect.h}`);
});

test("an inline element with opacity 0 contributes no text", opts, async () => {
  const s = await extractHtml(browser, '<p style="margin:0">Visible <span style="opacity:0">HIDDEN</span> tail</p>');
  const text = s.nodes.filter((n) => n.kind === "text").map(textOf).join("");
  assert.ok(text.includes("Visible") && text.includes("tail"), text);
  assert.ok(!text.includes("HIDDEN"), text);
});

test("images: content box, natural size, object-fit and border; a broken image is reported missing", opts, async () => {
  const s = await extractHtml(browser,
    `<img src="${PNG}" alt="Pic" style="display:block;width:100px;height:100px;object-fit:cover;padding:5px;border:1px solid #000">` +
      '<img src="http://127.0.0.1:9/nope.png" alt="x" style="width:20px;height:20px">');
  const [img, missing] = s.nodes;
  assert.deepEqual({ kind: img.kind, rect: img.rect, natural: img.natural, fit: img.fit, format: img.format, alt: img.alt, border: img.border }, {
    kind: "image", rect: { x: 46, y: 46, w: 100, h: 100 }, natural: { w: 4, h: 2 }, fit: "cover", format: "png", alt: "Pic", border: { width: 1, color: { hex: "000000", alpha: 1 } },
  });
  assert.equal(missing.kind, "missing");
  assert.equal(missing.src, "http://127.0.0.1:9/nope.png");
});

test("a translucent background image is a picture, like a translucent img", opts, async () => {
  const s = await extractHtml(browser,
    `<div style="width:100px;height:60px;opacity:0.5;background-image:url(${PNG});background-size:cover"></div>`);
  assert.deepEqual(s.nodes.map((n) => [n.kind, n.reason]), [["raster", "image opacity"]]);
});

test("hidden, transparent and zero-size elements are skipped; opacity multiplies down; one-sided borders are sides", opts, async () => {
  const s = await extractHtml(browser,
    '<div style="display:none">a</div><div style="visibility:hidden">b</div><div style="opacity:0">c</div><div style="width:0;height:0;background:red"></div>' +
      '<div style="opacity:0.5"><div style="background:#000;opacity:0.5;width:10px;height:10px"></div></div>' +
      '<blockquote style="margin:0;border-left:4px solid #1a833b;padding-left:8px">quoted</blockquote>');
  assert.deepEqual(s.nodes.filter((n) => n.kind === "text").map(textOf), ["quoted"]);
  const [faded, quote] = s.nodes.filter((n) => n.kind === "box");
  assert.equal(faded.opacity, 0.25);
  assert.deepEqual({ fill: quote.fill, border: quote.border, sides: quote.sides }, { fill: null, border: null, sides: [{ side: "l", width: 4, color: { hex: "1A833B", alpha: 1 } }] });
});

test("a rotated text-only element keeps its own box and a rotation; skewed elements are rasters", opts, async () => {
  const s = await extractHtml(browser, '<div style="width:100px;transform:rotate(90deg)">Turned</div><div style="transform:skewX(10deg)">Skewed</div>');
  const turned = s.nodes.find((n) => n.kind === "text");
  assert.equal(turned.rotate, 90);
  assert.equal(turned.rect.w, 100);
  assert.deepEqual(s.nodes.filter((n) => n.kind === "raster").map((n) => n.reason), ["transform"]);
});

test("backgrounds: the layout root's color is the slide background; a gradient asks for a picture background", opts, async () => {
  assert.deepEqual((await extractHtml(browser, "<p>x</p>", { layoutStyle: "background:#f8fafc" })).background, { color: { hex: "F8FAFC", alpha: 1 } });
  const g = await extractHtml(browser, "<p>x</p>", { layoutStyle: "background:radial-gradient(#fff, #000)" });
  assert.deepEqual(g.background, { raster: true, reason: "gradient background" });
  assert.equal(g.layout, "default");
});

test("a custom layout without the slidev-layout class still extracts, on white", opts, async () => {
  const s = await extractHtml(browser, '<h1 style="margin:0">Custom</h1>', { layoutClass: "my-layout" });
  assert.equal(s.layout, "default");
  assert.deepEqual(s.background, { color: { hex: "FFFFFF", alpha: 1 } });
  assert.deepEqual(s.nodes.map((n) => [n.kind, n.role]), [["text", "title"]]);
});
