import { test, after } from "node:test";
import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { ROOT, usableChromium, extractHtml } from "./helpers.mjs";

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

test("a translucent inline highlight is the color the page shows: composited over the background beneath it", opts, async () => {
  const s = await extractHtml(browser,
    '<p style="margin:0;font:16px Arial">On white <code style="background:rgba(26, 131, 59, 0.12)">tinted</code></p>' +
      '<p style="margin:0;font:16px Arial;background:#000;color:#fff">On black <code style="background:rgba(255, 255, 255, 0.5)">gray</code> <b style="background:#ff0000">red</b></p>');
  const highlights = s.nodes.filter((n) => n.kind === "text").flatMap((n) => n.runs.filter((r) => r.highlight).map((r) => [r.text, r.highlight]));
  assert.deepEqual(highlights, [["tinted", { hex: "E4F0E7", alpha: 1 }], ["gray", { hex: "808080", alpha: 1 }], ["red", { hex: "FF0000", alpha: 1 }]]);
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

test("a translucent background image is a picture behind the element's contents, which stay native with the element's opacity", opts, async () => {
  const s = await extractHtml(browser,
    `<div class="hero" style="width:200px;height:60px;opacity:0.5;background-image:url(${PNG});background-size:cover;background-position:center"><p style="margin:0;font:20px Arial;color:#000">Over the photo</p></div>`);
  assert.deepEqual(s.nodes.map((n) => [n.kind, n.reason, n.label, n.contentsHidden]), [
    ["raster", "image opacity", "div.hero background", true],
    ["text", undefined, undefined, undefined],
  ]);
  const [picture, text] = s.nodes;
  assert.deepEqual(picture.rect, { x: 40, y: 40, w: 200, h: 60 });
  assert.deepEqual([textOf(text), text.opacity], ["Over the photo", 0.5]);
});

test("a card whose background the box can't draw (an SVG pattern) is its box, a picture of that background, then its native contents", opts, async () => {
  const pattern = `data:image/svg+xml;charset=utf-8,${encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="8" height="8"><circle cx="4" cy="4" r="1" fill="#1a833b"/></svg>')}`;
  const s = await extractHtml(browser,
    `<div class="kv-card" style="width:276px;padding:20px;border-radius:8px;border:1px solid #D9E3DB;background:#EEF4F0 url(&quot;${pattern}&quot;);box-shadow:0 2px 4px rgba(0,0,0,0.2)">` +
      '<h3 style="margin:0;font:24px Arial">First point</h3><ul style="margin:0"><li>Native list</li></ul></div>');
  assert.deepEqual(s.nodes.map((n) => [n.kind, n.reason ?? null]), [["box", null], ["raster", "background image"], ["text", null], ["list", null]]);
  const [box, picture] = s.nodes;
  assert.deepEqual([box.fill, box.border.color.hex, box.radius, !!box.shadow], [{ hex: "EEF4F0", alpha: 1 }, "D9E3DB", 8, true]);
  assert.deepEqual([picture.label, picture.contentsHidden, picture.rect], ["div.kv-card background", true, box.rect]);
  // Element-level reasons still picture the whole element, contents included.
  const blurred = await extractHtml(browser, `<div style="filter:blur(1px);background:url(&quot;${pattern}&quot;)"><p>Inside</p></div>`);
  assert.deepEqual(blurred.nodes.map((n) => [n.kind, n.reason]), [["raster", "CSS filter"]]);
  const layered = await extractHtml(browser, '<div style="width:100px;height:50px;background:linear-gradient(#f00, #00f), linear-gradient(#0f0, #000)">Text</div>');
  assert.deepEqual(layered.nodes.map((n) => [n.kind, n.reason ?? textOf(n)]), [["raster", "background layers"], ["text", "Text"]]);
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

test("backgrounds: the layout root's color is the slide background; an image background is a picture", opts, async () => {
  assert.deepEqual((await extractHtml(browser, "<p>x</p>", { layoutStyle: "background:#f8fafc" })).background, { color: { hex: "F8FAFC", alpha: 1 } });
  const g = await extractHtml(browser, "<p>x</p>", { layoutStyle: "background:url(data:image/gif;base64,R0lGODlhAQABAAAAACw=), #fff" });
  assert.deepEqual(g.background, { raster: true, reason: "background image" });
  assert.equal(g.layout, "default");
});

test("a layout that paints no background takes the color Slidev paints around the page (dark mode); translucent ones are composited onto it", opts, async () => {
  assert.deepEqual((await extractHtml(browser, "<p>x</p>", { contentStyle: "background:#121212" })).background, { color: { hex: "121212", alpha: 1 } });
  const veiled = await extractHtml(browser, "<p>x</p>", { contentStyle: "background:#000000", layoutStyle: "background:rgba(255, 255, 255, 0.5)" });
  assert.deepEqual(veiled.background, { color: { hex: "808080", alpha: 1 } });
  const glow = await extractHtml(browser, "<p>x</p>", { contentStyle: "background:#121212", layoutStyle: "background:linear-gradient(#ff0000, transparent)" });
  assert.deepEqual(glow.background.base, { hex: "121212", alpha: 1 }, "a gradient with no color of its own sits on the color beneath it");
});

test("gradients: a radial glow over a color is a native background gradient; a linear gradient fills a box; layered backgrounds stay pictures", opts, async () => {
  const glow = await extractHtml(browser, "<p>x</p>", { layoutStyle: "background:radial-gradient(60% 50% at 50% 38%, rgba(26, 131, 59, 0.12) 0%, transparent 70%), #ffffff" });
  assert.deepEqual(glow.background, {
    gradient: { type: "radial", cx: 0.5, cy: 0.38, rx: 588, ry: 276, stops: [{ pos: 0, color: { hex: "1A833B", alpha: 0.12 } }, { pos: 0.7, color: { hex: "000000", alpha: 0 } }] },
    base: { hex: "FFFFFF", alpha: 1 },
  });
  const lin = await extractHtml(browser, '<div style="width:100px;height:50px;background:linear-gradient(to right, #ff0000, #0000ff)"></div>');
  assert.deepEqual(lin.nodes.find((n) => n.kind === "box").fill, {
    gradient: { type: "linear", angle: 90, stops: [{ pos: 0, color: { hex: "FF0000", alpha: 1 } }, { pos: 1, color: { hex: "0000FF", alpha: 1 } }] },
    base: null,
  });
  const layered = await extractHtml(browser, '<div style="width:100px;height:50px;background:linear-gradient(#f00, #00f), linear-gradient(#0f0, #000)"></div>');
  assert.deepEqual(layered.nodes.map((n) => [n.kind, n.reason]), [["raster", "background layers"]]);
});

test("gradient geometry: corners aim across the box, an `at` position alone works, stop positions stay in order; patterns and unsupported forms stay pictures", opts, async () => {
  const fillOf = async (css, size = "") => {
    const s = await extractHtml(browser, `<div style="width:200px;height:100px;background:${css};${size}"></div>`);
    return s.nodes[0].kind === "box" ? s.nodes[0].fill.gradient : s.nodes[0].reason;
  };
  assert.equal((await fillOf("linear-gradient(to bottom right, #fff, #00f)")).angle, 153.43);
  assert.equal((await fillOf("linear-gradient(to left, #fff, #00f)")).angle, 270);
  const at = await fillOf("radial-gradient(at 30% 40%, #f00, #00f)");
  assert.deepEqual([at.cx, at.cy], [0.3, 0.4]);
  const px = await fillOf("linear-gradient(rgba(0,0,0,.5) 10px, transparent 40px)");
  assert.deepEqual(px.stops.map((x) => x.pos), [0.1, 0.4]);
  assert.deepEqual((await fillOf("linear-gradient(red 50%, blue 30%)")).stops.map((x) => x.pos), [0.5, 0.5]);
  for (const css of ["repeating-linear-gradient(red, blue 20px)", "conic-gradient(red, blue)", "linear-gradient(red, 30%, blue)", "linear-gradient(in oklch, red, blue)"]) {
    assert.equal(await fillOf(css), "gradient", css);
  }
  assert.equal(await fillOf("linear-gradient(red, blue)", "background-size:20px 20px"), "gradient");
});

test("radial gradients keep their shape: circle and ellipse get their own radii; a shape the box can't show, or a focus outside it, is a picture", opts, async () => {
  const radial = async (css, size = "width:200px;height:150px") => {
    const s = await extractHtml(browser, `<div style="${size};background:${css}"></div>`);
    const n = s.nodes[0];
    return n.kind === "box" ? n.fill.gradient : n.reason;
  };
  const radii = (g) => [g.rx, g.ry];
  assert.deepEqual(radii(await radial("radial-gradient(#fff, #000)")), [141.42, 106.07]);
  assert.deepEqual(radii(await radial("radial-gradient(circle, #fff, #000)")), [125, 125]);
  assert.deepEqual(radii(await radial("radial-gradient(circle closest-side, #fff, #000)")), [75, 75]);
  assert.deepEqual(radii(await radial("radial-gradient(ellipse closest-side, #fff, #000)", "width:400px;height:200px")), [200, 100]);
  assert.deepEqual(radii(await radial("radial-gradient(ellipse farthest-side at 40% 50%, #fff, #000)", "width:400px;height:200px")), [240, 100]);
  assert.deepEqual(radii(await radial("radial-gradient(ellipse closest-corner, #fff, #000)", "width:400px;height:200px")), [282.84, 141.42]);
  assert.deepEqual(radii(await radial("radial-gradient(60px, #fff, #000)")), [60, 60]);
  assert.deepEqual(radii(await radial("radial-gradient(120px 80px at 50% 50%, #fff, #000)")), [120, 80]);
  // A circle on a wide box, an ellipse far from the box's aspect and a focus outside the box.
  assert.equal(await radial("radial-gradient(circle, #fff, #000)", "width:400px;height:200px"), "radial gradient shape");
  assert.equal(await radial("radial-gradient(circle, #fff, #000)", "width:320px;height:180px"), "radial gradient shape");
  assert.equal(await radial("radial-gradient(100px 20px, #fff, #000)"), "radial gradient shape");
  assert.equal(await radial("radial-gradient(circle at 150% 50%, #fff, #000)"), "radial gradient focus");
  const slide = await extractHtml(browser, "<p>x</p>", { layoutStyle: "background:radial-gradient(circle, #fff, #000), #fff" });
  assert.deepEqual(slide.background, { raster: true, reason: "radial gradient shape" });
});

test("url() backgrounds are native only when stretched or cropped around the center; sized, tiled or off-center ones are pictures", opts, async () => {
  const bg = async (css, size = "width:200px;height:100px") => {
    const s = await extractHtml(browser, `<div style="${size};background:${css}"></div>`);
    const image = s.nodes.find((n) => n.kind === "image");
    return image ? image.fit : s.nodes.find((n) => n.kind === "raster").reason;
  };
  // The longhands list one value per layer ("cover, auto"); only the image's own layer counts.
  assert.equal(await bg(`url(${PNG}) center/cover no-repeat, #fff`), "cover");
  assert.equal(await bg(`url(${PNG}) center/cover`), "cover");
  assert.equal(await bg(`url(${PNG}) center/contain no-repeat`), "contain");
  assert.equal(await bg(`url(${PNG}) 0 0/100% 100%`), "fill");
  assert.equal(await bg(`url(${PNG}) center/contain`), "background image size");
  assert.equal(await bg(`url(${PNG}) 0 0/20px 20px`), "background image size");
  assert.equal(await bg(`url(${PNG})`), "background image size");
  assert.equal(await bg(`url(${PNG}) 0 0/cover no-repeat`), "background image position");
  assert.equal(await bg(`url(${PNG}) top/cover`), "background image position");
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

test("a code block that scrolls inside its panel is a picture; with overflow visible it stays native", opts, async () => {
  const long = "x".repeat(120);
  const pre = (style, body = long) => `<pre style="margin:0;font:13px/18px monospace;${style}">${body}</pre>`;
  const wide = await extractHtml(browser, pre("width:300px;overflow:auto"));
  assert.deepEqual(wide.nodes.map((n) => [n.kind, n.reason]), [["raster", "code overflows its panel"]]);
  const tall = await extractHtml(browser, pre("max-height:36px;overflow:auto", "1\n2\n3\n4\n5"));
  assert.deepEqual(tall.nodes.map((n) => [n.kind, n.reason]), [["raster", "code overflows its panel"]]);
  const visible = await extractHtml(browser, pre("width:300px;overflow:visible"));
  assert.deepEqual(visible.nodes.map((n) => n.kind), ["code"]);
  const fits = await extractHtml(browser, pre("width:300px;overflow:auto", "short"));
  assert.deepEqual(fits.nodes.map((n) => n.kind), ["code"]);
});

test("a pre-wrap code block that wraps is a picture; one that fits stays native", opts, async () => {
  const pre = (body) => `<pre style="margin:0;width:100px;white-space:pre-wrap;font:13px/18px monospace">${body}</pre>`;
  const wraps = await extractHtml(browser, pre("aaaa bbbb cccc dddd eeee ffff"));
  assert.deepEqual(wraps.nodes.map((n) => [n.kind, n.reason]), [["raster", "wrapped code"]]);
  const blankLine = await extractHtml(browser, pre("aaaa\n\nbbbb cccc dddd eeee ffff"));
  assert.deepEqual(blankLine.nodes.map((n) => [n.kind, n.reason]), [["raster", "wrapped code"]], "a blank line must not hide the wrap");
  const fits = await extractHtml(browser, pre("aaaa\n\nbbbb"));
  assert.deepEqual(fits.nodes.map((n) => n.kind), ["code"]);
});

test("code with line-height normal gets the browser's line pitch, not a guessed multiple of the font size", opts, async () => {
  const s = await extractHtml(browser, '<pre style="margin:0;font:13px monospace;line-height:normal">a\nb\n\nd</pre>');
  const code = s.nodes.find((n) => n.kind === "code");
  assert.equal(code.lines.length, 4);
  assert.ok(Math.abs(code.lineHeight - code.rect.h / 4) <= 1, `lineHeight ${code.lineHeight} vs ${code.rect.h / 4}`);
});

test("tokens inside dimmed spans carry the spans' opacity in their color alpha", opts, async () => {
  const s = await extractHtml(browser,
    '<pre style="margin:0;font:13px/18px monospace"><code><span style="opacity:.3"><span style="color:#1a589f">dim</span></span> bright\n</code></pre>');
  const code = s.nodes.find((n) => n.kind === "code");
  assert.deepEqual(code.lines[0].map((r) => [r.text, r.color.hex, r.color.alpha]), [["dim", "1A589F", 0.3], [" bright", "000000", 1]]);
});

test("a display:none cell is left out of the grid and its text out of the table", opts, async () => {
  const s = await extractHtml(browser, '<table style="font:14px/20px Arial"><tr><td>a</td><td style="display:none">HID</td><td>c</td></tr></table>');
  const t = s.nodes.find((n) => n.kind === "table");
  assert.equal(t.colW.length, 2);
  assert.ok(t.colW.every((w) => w > 0), t.colW.join());
  assert.deepEqual(t.rows[0].map(textOf), ["a", "c"]);
  assert.ok(!JSON.stringify(s.nodes).includes("HID"));
});

test("tabs in code expand to the next tab stop, counted across the spans of a line", opts, async () => {
  const s = await extractHtml(browser,
    '<pre style="margin:0;font:13px/18px monospace;tab-size:2">\ta\n<span>b</span>\t<span>c</span>\n\t\td</pre>');
  const code = s.nodes.find((n) => n.kind === "code");
  assert.deepEqual(code.lines.map((l) => l.map((r) => r.text).join("")), ["  a", "b c", "    d"]);
  const dflt = await extractHtml(browser, '<pre style="margin:0;font:13px/18px monospace">\ta</pre>');
  assert.equal(dflt.nodes.find((n) => n.kind === "code").lines[0][0].text, "        a");
});

test("a column no cell starts in shares the gap, and an empty cell still gets a line height from its font", opts, async () => {
  const s = await extractHtml(browser,
    '<table style="font:10px Arial"><tr><td colspan="2">A</td><td>B</td></tr><tr><td colspan="2"></td><td>D</td></tr></table>');
  const t = s.nodes.find((n) => n.kind === "table");
  assert.equal(t.colW.length, 3);
  assert.ok(t.colW.every((w) => w > 0), t.colW.join());
  assert.ok(Math.abs(t.colW.reduce((a, b) => a + b, 0) - t.rect.w) < 0.1, `${t.colW.join()} vs ${t.rect.w}`);
  assert.ok(Math.abs(t.colW[0] - t.colW[1]) < 0.1, "the two spanned columns split evenly");
  assert.deepEqual([t.rows[1][0].runs, t.rows[1][0].lineHeight], [[], 12]);
});

// Charts: real Chart.js (a devDependency) drawing real canvases, found through Chart.getChart.
const CHARTJS_FILE = join(ROOT, "node_modules/chart.js/dist/chart.umd.js");
const CHARTJS = existsSync(CHARTJS_FILE) ? readFileSync(CHARTJS_FILE, "utf8") : null;
const chartOpts = { skip: !browser ? opts.skip : CHARTJS ? false : "chart.js is not installed - run npm install" };
const LIB = `<script>${CHARTJS}</script><script>Chart.defaults.animation=false;Chart.defaults.responsive=false;</script>`;
// A canvas and the script that draws `config` (JS source) on it.
const canvasOf = (config, style = "display:block") =>
  `<canvas width="400" height="200" style="${style}"></canvas><script>new Chart(document.currentScript.previousElementSibling,${config})</script>`;
const bar = (datasets, options = "{}", labels = "['Q1','Q2']") => `{type:'bar',data:{labels:${labels},datasets:[${datasets}]},options:${options}}`;
const GREEN = "{label:'A',data:[22,31],backgroundColor:'#1A833B'}";
const chartPage = async (config, style) => (await extractHtml(browser, LIB + canvasOf(config, style))).nodes;
const C = (hex, alpha = 1) => ({ hex, alpha });

test("a Chart.js bar chart becomes a chart node with its data, colors, alpha, axis range, legend and tick format", chartOpts, async () => {
  const nodes = await chartPage(bar(
    "{label:'Before',data:[22,31],backgroundColor:'rgba(26,131,59,0.4)'},{label:'After',data:[41,null],backgroundColor:['#9CA3AF','#9CA3AF']},{label:'Hidden',data:[1,2],hidden:true}",
    "{scales:{y:{max:100,ticks:{color:'#5b6b63',callback:(v)=>v+'%'},grid:{color:'rgba(20,24,26,0.1)'}},x:{grid:{display:false}}},plugins:{legend:{position:'bottom',labels:{color:'#1f2937'}}}}",
  ));
  assert.deepEqual(nodes.map((n) => n.kind), ["chart"]);
  assert.deepEqual(nodes[0].rect, { x: 40, y: 40, w: 400, h: 200 });
  assert.deepEqual(nodes[0].components, []);
  assert.deepEqual(nodes[0].chart, {
    type: "column", categories: ["Q1", "Q2"],
    series: [{ name: "Before", values: [22, 31], color: C("1A833B", 0.4) }, { name: "After", values: [41, null], color: C("9CA3AF") }],
    min: 0, max: 100, stacked: false, valueAxis: true, legend: true, legendPos: "b", numFmt: '0"%"',
    gridColor: C("14181A", 0.1), labelColor: C("5B6B63"), legendColor: C("1F2937"),
  });
});

test("a series without a color takes the one Chart.js drew with (the Colors plugin's, alpha included)", chartOpts, async () => {
  const [node] = await chartPage(bar("{label:'A',data:[1,2]},{label:'B',data:[3,4]}"));
  assert.equal(node.kind, "chart");
  assert.deepEqual(node.chart.series.map((s) => s.color), [C("36A2EB", 0.5), C("FF6384", 0.5)]);
});

test("a bar chart's value axis starts at zero; a suggested bound is where the drawn axis ended; Chart.js's own tick text stays native", chartOpts, async () => {
  const [narrow] = await chartPage(bar("{label:'A',data:[90,95,99],backgroundColor:'#1A833B'}", "{}", "['Q1','Q2','Q3']"));
  assert.deepEqual([narrow.kind, narrow.chart.min, narrow.chart.max, narrow.chart.numFmt], ["chart", 0, null, null]);
  const [suggested] = await chartPage(bar("{label:'A',data:[1,2],backgroundColor:'#1A833B'}", "{scales:{y:{suggestedMax:50}}}"));
  assert.deepEqual([suggested.kind, suggested.chart.min, suggested.chart.max], ["chart", 0, 50]);
  const [explicit] = await chartPage(bar("{label:'A',data:[1,2],backgroundColor:'#1A833B'}", "{scales:{y:{min:-5,max:10,suggestedMax:50}}}"));
  assert.deepEqual([explicit.chart.min, explicit.chart.max], [-5, 10]);
  // Fractional ticks (0.5, 1.0) and thousands (1,000) are what the native axis shows too.
  const [fraction] = await chartPage(bar("{label:'A',data:[0.5,2],backgroundColor:'#1A833B'}"));
  const [thousands] = await chartPage(bar("{label:'A',data:[900,2400],backgroundColor:'#1A833B'}"));
  assert.deepEqual([fraction.kind, thousands.kind], ["chart", "chart"]);
  const [dollars] = await chartPage(bar("{label:'A',data:[1,2],backgroundColor:'#1A833B'}", "{scales:{y:{ticks:{callback:(v)=>'$'+v}}}}"));
  assert.deepEqual([dollars.kind, dollars.chart.numFmt], ["chart", '"$"0']);
  const [fixed] = await chartPage(bar("{label:'A',data:[10,20],backgroundColor:'#1A833B'}", "{scales:{y:{ticks:{callback:(v)=>'$'+v.toFixed(1)}}}}"));
  assert.deepEqual([fixed.kind, fixed.chart.numFmt], ["chart", '"$"0.0']);
});

test("indexAxis y is a horizontal chart; a hidden value axis is kept hidden and takes its gridlines along", chartOpts, async () => {
  const [horizontal] = await chartPage(bar(GREEN, "{indexAxis:'y'}"));
  assert.deepEqual([horizontal.kind, horizontal.chart.type, horizontal.chart.categories, horizontal.chart.series[0].values], ["chart", "bar", ["Q1", "Q2"], [22, 31]]);
  const [hidden] = await chartPage(bar(GREEN, "{scales:{y:{display:false}}}"));
  assert.deepEqual([hidden.kind, hidden.chart.valueAxis, hidden.chart.gridColor], ["chart", false, null]);
});

test("a stacked chart (both axes stacked, one stack) is native and stacked", chartOpts, async () => {
  const [node] = await chartPage(bar(`${GREEN},{label:'B',data:[5,6],backgroundColor:'#9CA3AF'}`, "{scales:{x:{stacked:true},y:{stacked:true}}}"));
  assert.deepEqual([node.kind, node.chart.stacked], ["chart", true]);
});

test("the category axis's own labels option wins over data.labels", chartOpts, async () => {
  const [node] = await chartPage(bar(GREEN, "{scales:{x:{labels:['one','two']}}}"));
  assert.deepEqual([node.kind, node.chart.categories], ["chart", ["one", "two"]]);
});

test("a canvas with a background, border and padding is a box plus a chart in its content box", chartOpts, async () => {
  const nodes = await chartPage(bar(GREEN), "display:block;padding:10px;border:2px solid #1A833B;background:#EEF4F0");
  assert.deepEqual(nodes.map((n) => n.kind), ["box", "chart"]);
  // Chart.js sizes the canvas border-box: 400x200 including the 10px padding and 2px border.
  assert.deepEqual({ rect: nodes[0].rect, fill: nodes[0].fill, border: nodes[0].border }, { rect: { x: 40, y: 40, w: 400, h: 200 }, fill: C("EEF4F0"), border: { width: 2, color: C("1A833B") } });
  assert.deepEqual(nodes[1].rect, { x: 52, y: 52, w: 376, h: 176 });
});

// A chart is native only when it can be drawn faithfully; everything else is a picture.
const PICTURES = [
  ["a mixed chart", bar(`${GREEN},{type:'line',label:'B',data:[5,6],borderColor:'#9CA3AF'}`), "display:block", "chart mixed chart types"],
  ["floating bars", bar("{label:'A',data:[[2,5],[3,9]],backgroundColor:'#1A833B'}"), "display:block", "chart data format"],
  ["{x, y} data", bar("{label:'A',data:[{x:'Q2',y:5},{x:'Q1',y:7}],backgroundColor:'#1A833B'}"), "display:block", "chart data format"],
  ["object-keyed data", bar("{label:'A',data:{Q1:5,Q2:7},backgroundColor:'#1A833B'}"), "display:block", "chart data format"],
  ["a custom parsing key", bar("{label:'A',data:[{m:'Q1',v:5},{m:'Q2',v:7}],parsing:{xAxisKey:'m',yAxisKey:'v'},backgroundColor:'#1A833B'}"), "display:block", "chart custom parsing"],
  ["a logarithmic axis", bar(GREEN, "{scales:{y:{type:'logarithmic'}}}"), "display:block", "chart value axis type"],
  ["a second value axis", bar("{label:'A',data:[22,31],yAxisID:'y',backgroundColor:'#1A833B'},{label:'B',data:[1,2],yAxisID:'y1',backgroundColor:'#9CA3AF'}", "{scales:{y1:{position:'right'}}}"), "display:block", "chart second value axis"],
  ["a value axis beside the one the datasets use", bar(`${GREEN},{label:'B',data:[1,2],yAxisID:'y1',backgroundColor:'#9CA3AF'}`, "{scales:{y1:{position:'right'}}}"), "display:block", "chart second value axis"],
  ["stack groups", bar(`{label:'A',data:[22,31],stack:'s1',backgroundColor:'#1A833B'},{label:'B',data:[5,6],stack:'s2',backgroundColor:'#9CA3AF'}`, "{scales:{x:{stacked:true},y:{stacked:true}}}"), "display:block", "chart stack groups"],
  ["one stacked axis", bar(GREEN, "{scales:{y:{stacked:true}}}"), "display:block", "chart stack groups"],
  ["a hidden category axis", bar(GREEN, "{scales:{x:{display:false}}}"), "display:block", "chart hidden category axis"],
  ["a reversed axis", bar(GREEN, "{scales:{y:{reverse:true}}}"), "display:block", "chart reversed axis"],
  ["a title", bar(GREEN, "{plugins:{title:{display:true,text:'Results'}}}"), "display:block", "chart title"],
  ["per-bar colors", bar("{label:'A',data:[22,31],backgroundColor:['#ff0000','#00ff00']}"), "display:block", "chart per-bar colors"],
  ["a scripted color", bar("{label:'A',data:[22,31],backgroundColor:(c)=>c.dataIndex?'#ff0000':'#00ff00'}"), "display:block", "chart scripted colors"],
  ["a gradient fill", bar("{label:'A',data:[22,31],backgroundColor:(()=>{const g=document.currentScript.previousElementSibling.getContext('2d').createLinearGradient(0,0,1,1);g.addColorStop(0,'red');return g})()}"), "display:block", "chart scripted colors"],
  ["tick labels a number format can't reproduce", bar(GREEN, "{scales:{y:{ticks:{callback:(v)=>(v/10).toFixed(1)}}}}"), "display:block", "chart tick labels"],
  ["no visible dataset", bar("{label:'A',data:[1,2],hidden:true}"), "display:block", "chart no visible datasets"],
  ["a rotated canvas", bar(GREEN), "display:block;transform:rotate(10deg)", "rotated chart"],
  ["a translucent canvas", bar(GREEN), "display:block;opacity:.5", "chart opacity"],
];
for (const [name, config, style, reason] of PICTURES) {
  test(`${name} is a picture (${reason})`, chartOpts, async () => {
    const nodes = await chartPage(config, style);
    assert.deepEqual(nodes.map((n) => [n.kind, n.reason]), [["raster", reason]]);
    assert.equal(nodes[0].label, "canvas");
  });
}

test("a chart found through its Vue component names the components around it; other canvases stay pictures", chartOpts, async () => {
  const setup = () => {
    const [barEl, pieEl] = document.querySelectorAll("canvas");
    const data = {
      labels: ["Q1", "Q2"],
      datasets: [
        { label: "Before", data: [22, 31], backgroundColor: "#9CA3AF" },
        { label: "After", data: [41, null], backgroundColor: ["#1A833B", "#1A833B"] },
        { label: "Hidden", data: [1, 2], hidden: true },
      ],
    };
    const options = {
      scales: { y: { min: 0, max: 100, ticks: { color: "#5b6b63", callback: (v) => v + "%" }, grid: { color: "rgba(20, 24, 26, 0.1)" } }, x: { grid: { display: false } } },
      plugins: { legend: { position: "bottom", labels: { color: "#1f2937" } } },
    };
    const barChart = new Chart(barEl, { type: "bar", data, options });
    const pieChart = new Chart(pieEl, { type: "pie", data: { labels: ["a", "b"], datasets: [{ data: [1, 2] }] } });
    // Only the component tree may lead to the chart here.
    Chart.getChart = () => undefined;
    barEl.__vueParentComponent = { type: {}, exposed: { chart: { value: barChart } }, parent: { type: { __name: "ExampleChart" }, parent: null } };
    pieEl.__vueParentComponent = { type: {}, exposed: { chart: { value: pieChart } }, parent: null };
  };
  const s = await extractHtml(browser,
    `${LIB}<canvas width="400" height="200" style="display:block"></canvas><canvas width="100" height="100" style="display:block"></canvas><canvas width="50" height="50" style="display:block"></canvas>`,
    { setup });
  const [chart, pie, plain] = s.nodes;
  assert.equal(chart.kind, "chart");
  assert.deepEqual(chart.components, ["ExampleChart"]);
  assert.deepEqual(chart.chart, {
    type: "column", categories: ["Q1", "Q2"],
    series: [{ name: "Before", values: [22, 31], color: C("9CA3AF") }, { name: "After", values: [41, null], color: C("1A833B") }],
    min: 0, max: 100, stacked: false, valueAxis: true, legend: true, legendPos: "b", numFmt: '0"%"',
    gridColor: C("14181A", 0.1), labelColor: C("5B6B63"), legendColor: C("1F2937"),
  });
  assert.deepEqual([pie.kind, pie.reason], ["raster", "chart type pie"]);
  assert.deepEqual([plain.kind, plain.reason], ["raster", "canvas"]);
});

test("a canvas that becomes a picture still names the Vue components around it, and where a chart would go", chartOpts, async () => {
  const setup = () => {
    const [pieEl, brokenEl, plainEl, blurredEl] = document.querySelectorAll("canvas");
    const pie = new Chart(pieEl, { type: "pie", data: { labels: ["a", "b"], datasets: [{ data: [1, 2] }] } });
    Chart.getChart = () => undefined;
    const owner = (name, exposed = {}) => ({ type: {}, exposed, parent: { type: { __name: name }, parent: { type: { name: "Slide" }, parent: null } } });
    pieEl.__vueParentComponent = owner("GatedChart", { chart: { value: pie } });
    // A chart whose options can't be read.
    const broken = { config: { type: "bar" }, data: { datasets: [] }, canvas: brokenEl, get options() { throw new Error("options gone\nsecond line"); } };
    brokenEl.__vueParentComponent = owner("BrokenChart", { chart: { value: broken } });
    plainEl.__vueParentComponent = owner("Sketch");
    blurredEl.__vueParentComponent = owner("Blurred");
  };
  const s = await extractHtml(browser,
    `${LIB}<canvas width="200" height="100" style="display:block;padding:10px;border:2px solid #000"></canvas><canvas width="100" height="50" style="display:block"></canvas><canvas width="50" height="50" style="display:block"></canvas><canvas width="50" height="50" style="display:block;filter:blur(1px)"></canvas>`,
    { setup });
  assert.deepEqual(s.nodes.map((n) => [n.kind, n.reason, n.components]), [
    ["raster", "chart type pie", ["GatedChart", "Slide"]],
    ["raster", "chart unreadable: options gone", ["BrokenChart", "Slide"]],
    ["raster", "canvas", ["Sketch", "Slide"]],
    ["raster", "CSS filter", ["Blurred", "Slide"]],
  ]);
  const [pie] = s.nodes;
  assert.deepEqual(pie.chartRect, { x: pie.rect.x + 12, y: pie.rect.y + 12, w: pie.rect.w - 24, h: pie.rect.h - 24 }, "the chart goes in the canvas's content box");
});

test("a rounded image is a picture: its own corners, or a rounded box that clips it where it reaches into a corner", opts, async () => {
  const img = (style) => `<img src="${PNG}" alt="Photo" style="display:block;width:120px;height:80px;${style}">`;
  const kind = async (html) => (await extractHtml(browser, html)).nodes.filter((n) => n.alt === "Photo").map((n) => [n.kind, n.reason ?? null]);
  assert.deepEqual(await kind(img("border-radius:12px")), [["raster", "rounded image"]]);
  assert.deepEqual(await kind(img("border-bottom-right-radius:50%")), [["raster", "rounded image"]]);
  assert.deepEqual(await kind(img("")), [["image", null]]);
  const box = (style, inner) => `<div style="width:200px;height:120px;border-radius:16px;${style}">${inner}</div>`;
  assert.deepEqual(await kind(box("overflow:hidden", img(""))), [["raster", "rounded image"]], "the image sits in the box's top-left corner");
  assert.deepEqual(await kind(box("overflow:hidden;padding:20px;box-sizing:border-box", img(""))), [["image", null]], "padding keeps it clear of the corners");
  assert.deepEqual(await kind(box("", img(""))), [["image", null]], "a box that doesn't clip");
  assert.deepEqual(await kind(box("overflow:hidden;border-radius:0", img(""))), [["image", null]], "a square box");
});

test("an SVG image is an image node, not a picture of an element", opts, async () => {
  const svg = `data:image/svg+xml;charset=utf-8,${encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20"><rect width="40" height="20" fill="#1a833b"/></svg>')}`;
  const s = await extractHtml(browser, `<img src="${svg}" alt="Logo" style="display:block;width:80px;height:40px">`);
  assert.deepEqual(s.nodes.map((n) => [n.kind, n.format, n.alt]), [["image", "svg", "Logo"]]);
});

const svgUrl = (w, h) => `data:image/svg+xml;charset=utf-8,${encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}"><rect width="${w}" height="${h}" fill="#1a833b"/></svg>`)}`;

test("an SVG image in a box of another aspect is a picture of the page's drawing", opts, async () => {
  // A 2:1 SVG in a square box: the page letterboxes, crops or leaves it at natural size depending on object-fit.
  for (const fit of ["fill", "contain", "cover"]) {
    const s = await extractHtml(browser, `<img src="${svgUrl(40, 20)}" style="display:block;width:100px;height:100px;object-fit:${fit}">`);
    assert.deepEqual(s.nodes.map((n) => [n.kind, n.reason]), [["raster", "SVG image fit"]], fit);
  }
});

test("object-fit: none or scale-down on an SVG image is a picture even in a box of its own aspect", opts, async () => {
  for (const fit of ["none", "scale-down"]) {
    const s = await extractHtml(browser, `<img src="${svgUrl(40, 20)}" style="display:block;width:80px;height:40px;object-fit:${fit}">`);
    assert.deepEqual(s.nodes.map((n) => [n.kind, n.reason]), [["raster", "SVG image fit"]], fit);
  }
});

test("an SVG image in a box of its own aspect stays an image node for fill, contain and cover; padding counts out of the box", opts, async () => {
  for (const fit of ["fill", "contain", "cover"]) {
    const s = await extractHtml(browser, `<img src="${svgUrl(40, 20)}" style="display:block;box-sizing:border-box;width:84px;height:44px;padding:2px;object-fit:${fit}">`);
    assert.deepEqual(s.nodes.map((n) => [n.kind, n.format, n.fit, n.rect.w, n.rect.h]), [["image", "svg", fit, 80, 40]], fit);
  }
});

test("a rotated SVG image is judged by its own box, not the rotated hull", opts, async () => {
  const s = await extractHtml(browser, `<img src="${svgUrl(40, 20)}" style="display:block;width:80px;height:40px;transform:rotate(30deg)">`);
  assert.deepEqual(s.nodes.map((n) => [n.kind, n.format, n.rotate]), [["image", "svg", 30]]);
});

test("a broken SVG image is skipped as missing, like any other image that didn't load", opts, async () => {
  const bad = "data:image/svg+xml;charset=utf-8,this%20is%20not%20an%20svg";
  const s = await extractHtml(browser, `<img src="${bad}" alt="Logo" style="display:block;width:80px;height:40px">`);
  assert.deepEqual(s.nodes.map((n) => [n.kind, n.src === bad]), [["missing", true]]);
});

test("global layers (global-bottom.vue / global-top.vue) are extracted in DOM order, in the page's coordinates; Slidev's own UI is not", opts, async () => {
  const s = await extractHtml(browser, '<p style="margin:0;font:20px Arial">Page text</p>', {
    globalBottom: '<div style="position:absolute;left:0;top:0;width:980px;height:8px;background:#1a833b"></div>',
    globalTop: '<div style="position:absolute;right:24px;bottom:16px;font:12px Arial">Global footer</div><svg class="w-full h-full absolute top-0" style="position:absolute;top:0;left:0;width:100%;height:100%"></svg><svg class="drawn" width="40" height="40" style="position:absolute;left:0;top:0"><rect width="40" height="40"/></svg>',
    setup: () => {
      // The drawing layer is a component of @slidev/client.
      document.querySelector("svg.drawn").__vueParentComponent = { type: { __file: "/deck/node_modules/@slidev/client/internals/DrawingLayer.vue" } };
    },
  });
  assert.deepEqual(s.nodes.map((n) => [n.kind, n.kind === "text" ? textOf(n) : n.fill.hex]), [["box", "1A833B"], ["text", "Page text"], ["text", "Global footer"]]);
  assert.deepEqual(s.nodes[0].rect, { x: 0, y: 0, w: 980, h: 8 });
  const footer = s.nodes[2];
  assert.ok(Math.abs(footer.rect.x + footer.rect.w - 956) < 1 && Math.abs(footer.rect.y + footer.rect.h - 536) < 2, JSON.stringify(footer.rect));
  assert.equal(footer.role, "body");
  const plain = await extractHtml(browser, '<p style="margin:0;font:20px Arial">Page text</p>');
  assert.deepEqual(plain.nodes.map((n) => n.kind), ["text"]);
});

test("the global bottom layer is left out where the page paints over it: an opaque color or picture, not a fading gradient", opts, async () => {
  const band = '<div style="position:absolute;left:0;top:0;width:980px;height:8px;background:#1a833b"></div>';
  const kinds = async (options) => (await extractHtml(browser, "<p>x</p>", { globalBottom: band, ...options })).nodes.map((n) => n.kind);
  assert.deepEqual(await kinds({ layoutStyle: "background:#ffffff" }), ["text"]);
  assert.deepEqual(await kinds({ pageStyle: "background:#ffffff" }), ["text"]);
  assert.deepEqual(await kinds({ layoutStyle: `background:url(${PNG}) center / cover` }), ["text"]);
  assert.deepEqual(await kinds({ layoutStyle: "background:rgba(255, 255, 255, 0.5)" }), ["box", "text"]);
  assert.deepEqual(await kinds({ layoutStyle: "background:linear-gradient(#ffffff, transparent)" }), ["box", "text"]);
  assert.deepEqual(await kinds({ contentStyle: "background:#ffffff" }), ["box", "text"], "Slidev's own background is beneath the layer");
});

// Slidev's `zoom:` frontmatter: the page is 1/zoom as large and scaled down with the `scale` property.
const ZOOMED = "width:1960px;height:1104px;scale:0.5;transform-origin:top left";

test("a zoomed slide (zoom: 0.5) is measured in canvas px: fonts, line height, spacing, padding, borders and radii scale with the rects", opts, async () => {
  const card = (font, lh, pad, border, radius, spacing) =>
    `<div style="width:fit-content;padding:${pad}px;border:${border}px solid #1a833b;border-radius:${radius}px;background:#eef4f0"><p style="margin:0;font:${font}px/${lh}px Arial;letter-spacing:${spacing}px">Zoomed card</p></div>`;
  const zoomed = await extractHtml(browser, card(20, 30, 10, 2, 8, 2), { pageStyle: ZOOMED, globalTop: '<div style="position:absolute;left:10px;top:500px;font:12px Arial">Global</div>' });
  const plain = await extractHtml(browser, card(10, 15, 5, 1, 4, 1), { layoutStyle: "padding:20px" });
  const pick = (s) => {
    const box = s.nodes.find((n) => n.kind === "box");
    const text = s.nodes.find((n) => n.kind === "text" && textOf(n) === "Zoomed card");
    return { box: [box.border.width, box.radius], run: [text.runs[0].size, text.runs[0].spacing], lineHeight: text.lineHeight, boxRect: box.rect, textRect: text.rect };
  };
  const z = pick(zoomed);
  const p = pick(plain);
  assert.deepEqual([z.box, z.run, z.lineHeight], [p.box, p.run, p.lineHeight]);
  assert.deepEqual(z.run, [10, 1]);
  for (const k of ["x", "y", "w", "h"]) {
    assert.ok(Math.abs(z.boxRect[k] - p.boxRect[k]) < 1, `box ${k}: ${z.boxRect[k]} vs ${p.boxRect[k]}`);
    assert.ok(Math.abs(z.textRect[k] - p.textRect[k]) < 1.5, `text ${k}: ${z.textRect[k]} vs ${p.textRect[k]}`);
  }
  // The global layers sit outside the zoomed page: their text keeps its size.
  assert.equal(zoomed.nodes.find((n) => n.kind === "text" && textOf(n) === "Global").runs[0].size, 12);
});

test("the individual rotate, scale and translate properties count like transform", opts, async () => {
  const s = await extractHtml(browser,
    '<div style="width:100px;rotate:10deg">Turned</div><div style="width:100px;scale:0.5">Scaled</div><div style="width:100px;translate:30px 5px;margin-top:40px">Moved</div><div style="rotate:x 30deg">Tilted</div>');
  const [turned, moved] = s.nodes.filter((n) => n.kind === "text");
  assert.deepEqual([textOf(turned), turned.rotate, turned.rect.w], ["Turned", 10, 100]);
  assert.deepEqual(s.nodes.filter((n) => n.kind === "raster").map((n) => n.reason), ["transform", "transform"]);
  assert.equal(textOf(moved), "Moved");
  assert.ok(moved.rect.x >= 69.5, `moved x ${moved.rect.x}`);
  assert.equal(moved.rotate, 0);
});
