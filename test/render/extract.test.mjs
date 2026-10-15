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

test("hidden list items and hidden paragraphs inside list items contribute no text", opts, async () => {
  const s = await extractHtml(browser,
    '<ul style="margin:0;font:16px/24px Arial"><li>Shown</li><li style="opacity:0">HIDDEN-LI</li><li><p style="display:none">HIDDEN-P</p>Visible</li></ul>');
  const list = s.nodes.find((n) => n.kind === "list");
  assert.deepEqual(list.items.map(textOf), ["Shown", "Visible"]);
  const allText = s.nodes.flatMap((n) => (n.kind === "list" ? n.items : n.runs ? [n] : [])).map(textOf).join("|");
  assert.ok(!allText.includes("HIDDEN"), allText);
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

test("tables become table nodes: a grid with spans, measured columns and rows, per-cell fill, borders, padding and alignment", opts, async () => {
  const s = await extractHtml(browser,
    '<table style="border-collapse:collapse;font:14px/20px Arial">' +
      '<tr><th style="background:#eef4f0;border-bottom:2px solid #1a833b;padding:4px 8px;text-align:left">Option</th><th style="padding:4px 8px" colspan="2">Pros and cons</th></tr>' +
      '<tr><td style="padding:4px 8px;vertical-align:middle" rowspan="2">A</td><td style="padding:4px 8px">Simple</td><td style="padding:4px 8px">Slower</td></tr>' +
      '<tr><td style="padding:4px 8px">Fast</td><td style="padding:4px 8px;text-align:right">More <b>setup</b></td></tr></table>');
  const t = s.nodes.find((n) => n.kind === "table");
  assert.ok(t, JSON.stringify(s.nodes.map((n) => n.kind)));
  assert.equal(s.nodes.filter((n) => n.kind === "text").length, 0, "cell text lives only in the table node");
  assert.deepEqual([t.colW.length, t.rowH.length], [3, 3]);
  assert.ok(Math.abs(t.colW.reduce((a, b) => a + b, 0) - t.rect.w) < 1);
  const [head, span, covered] = t.rows[0];
  assert.deepEqual([textOf(head), head.fill, head.borders.b, head.padding, head.align], [
    "Option", { hex: "EEF4F0", alpha: 1 }, { width: 2, color: { hex: "1A833B", alpha: 1 } }, { t: 4, r: 8, b: 4, l: 8 }, "left",
  ]);
  assert.deepEqual([span.colspan, span.align], [2, "center"]);
  assert.deepEqual(covered, { covered: "h" });
  assert.deepEqual([t.rows[1][0].rowspan, t.rows[1][0].valign, t.rows[2][0]], [2, "middle", { covered: "v" }]);
  assert.deepEqual([textOf(t.rows[2][2]), t.rows[2][2].align, t.rows[2][2].runs[1].bold], ["More setup", "right", true]);
});

test("code blocks become a panel box plus one line of runs per source line, keeping token colors and leading spaces", opts, async () => {
  const s = await extractHtml(browser,
    '<pre style="margin:0;background:#eef4f0;padding:10px;font:13px/18px monospace"><code><span style="color:#1a589f">const</span> x = 1;\n<span style="color:#1a833b">  return</span> x;\n</code></pre>');
  const [panel, code] = s.nodes;
  assert.deepEqual([panel.kind, panel.fill.hex], ["box", "EEF4F0"]);
  assert.equal(code.kind, "code");
  assert.equal(code.lineHeight, 18);
  assert.deepEqual({ x: code.rect.x, y: code.rect.y }, { x: 50, y: 50 });
  assert.deepEqual(code.lines.map((l) => l.map((r) => [r.text, r.color.hex])), [
    [["const", "1A589F"], [" x = 1;", "000000"]],
    [["  return", "1A833B"], [" x;", "000000"]],
  ]);
});

test("collapsed-border tables: a row's rules land on its cells, the wider rule wins, and a row group's background fills its cells", opts, async () => {
  const s = await extractHtml(browser,
    '<table style="border-collapse:collapse;font:14px/20px Arial"><thead style="background:#eef4f0"><tr style="border-bottom:2px solid #1a833b"><th>H1</th><th>H2</th></tr></thead>' +
      '<tbody><tr style="border-bottom:1px solid #999"><td rowspan="2">A</td><td style="border-bottom:3px solid #1a589f">B</td></tr><tr style="border-bottom:1px solid #999"><td>C</td></tr></tbody></table>');
  const t = s.nodes.find((n) => n.kind === "table");
  const green = { width: 2, color: { hex: "1A833B", alpha: 1 } };
  assert.deepEqual([t.rows[0][0].borders, t.rows[0][1].borders], [{ b: green }, { b: green }]);
  assert.deepEqual([t.rows[0][0].fill, t.rows[0][1].fill], [{ hex: "EEF4F0", alpha: 1 }, { hex: "EEF4F0", alpha: 1 }]);
  assert.equal(t.rows[1][0].fill, null);
  assert.deepEqual(t.rows[1][1].borders.b, { width: 3, color: { hex: "1A589F", alpha: 1 } }, "the cell's own wider rule beats its row's");
  assert.deepEqual(t.rows[1][0].borders.b, { width: 1, color: { hex: "999999", alpha: 1 } }, "a row-spanning cell takes the rule of the last row it spans");
});

test("table cells: hidden paragraphs add no text; a list inside a cell and a caption become nodes of their own on top of the table", opts, async () => {
  const s = await extractHtml(browser,
    '<table style="font:14px/20px Arial"><caption>Caption text</caption><tr><td><p style="display:none">HIDDEN</p><p>First</p><p>Second</p></td><td><ul style="margin:0"><li>Nested item</li></ul></td></tr></table>');
  const kinds = s.nodes.map((n) => n.kind);
  assert.equal(kinds[0], "table");
  const t = s.nodes[0];
  assert.deepEqual([t.rows[0][0].runs.map((r) => r.text), t.rows[0][1].runs], [["First\nSecond"], []], "paragraphs join with a break, like list items");
  assert.ok(kinds.includes("list"), kinds.join());
  assert.ok(s.nodes.some((n) => n.kind === "text" && textOf(n) === "Caption text"), "the caption is drawn as text");
  assert.ok(!JSON.stringify(s.nodes).includes("HIDDEN"));
});

test("code blocks: the token spans' own decorations count, hidden spans are skipped, and a blank line in the middle stays a line", opts, async () => {
  const s = await extractHtml(browser,
    '<pre style="margin:0;font:13px/18px monospace;text-decoration:underline"><code>a\n\n<span style="opacity:0">GONE</span>b\n</code></pre>');
  const code = s.nodes.find((n) => n.kind === "code");
  assert.deepEqual(code.lines.map((l) => l.map((r) => r.text)), [["a"], [], ["b"]]);
  assert.equal(code.lines[0][0].underline, true);
});
