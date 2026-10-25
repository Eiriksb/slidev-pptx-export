// Render mode: Scene (what extract.js measured in the browser, in canvas px) -> SlideSpec (what
// the OOXML writer draws, in inches and points). Pure - no browser, no file access - so every
// mapping rule is unit-tested with hand-written scenes.

import { toColor } from "../colors.mjs";
import { PALETTE } from "../layout.mjs";

const ALIGN = { left: "l", start: "l", center: "ctr", right: "r", end: "r", justify: "just" };

const round = (v, digits = 4) => Math.round(v * 10 ** digits) / 10 ** digits;

// 1 canvas px = k inches, where k = slide width / canvas width.
export function unitsFor(scene, geometry) {
  const k = geometry.width / scene.canvas.w;
  return {
    inch: (px) => round(px * k),
    pt: (px) => round(px * k * 72, 2),
    box: (r) => ({ x: round(r.x * k), y: round(r.y * k), w: round(r.w * k), h: round(r.h * k) }),
  };
}

function colorOf(c, style) {
  return toColor(c.hex, style.theme);
}

// The alpha to write after the node's cumulative opacity; undefined when fully opaque.
function alphaOf(c, opacity = 1) {
  const a = round((c.alpha ?? 1) * opacity, 3);
  return a < 1 ? a : undefined;
}

// Fonts that are the theme's heading/body/mono fonts become theme font references.
function fontOf(family, fonts) {
  const f = String(family || "").toLowerCase();
  for (const role of ["heading", "body", "mono"]) {
    if (fonts[role] && fonts[role].toLowerCase() === f) return { font: role };
  }
  return family ? { typeface: family } : { font: "body" };
}

function runOf(r, u, style, opacity) {
  const run = { text: r.text, size: u.pt(r.size), ...fontOf(r.font, style.fonts) };
  if (r.bold) run.b = true;
  if (r.italic) run.i = true;
  if (r.underline) run.u = true;
  if (r.strike) run.strike = true;
  if (r.spacing) run.spacing = u.pt(r.spacing);
  if (r.color) {
    run.color = colorOf(r.color, style);
    const a = alphaOf(r.color, opacity);
    if (a !== undefined) run.alpha = a;
  }
  if (r.highlight) run.highlight = colorOf(r.highlight, style);
  return run;
}

function textShape(node, u, style) {
  const para = { runs: node.runs.map((r) => runOf(r, u, style, node.opacity ?? 1)) };
  const align = ALIGN[node.align];
  if (align && align !== "l") para.align = align;
  if (node.lineHeight) para.lineSpacing = u.pt(node.lineHeight);
  const shape = { kind: "text", ...u.box(node.rect), geometry: "rect", anchor: "t", inset: 0, paras: [para] };
  if (node.singleLine) shape.wrap = false;
  if (node.rotate) shape.rot = node.rotate;
  return shape;
}

function listShape(node, u, style) {
  const paras = node.items.map((it) => {
    const para = { level: it.level };
    para.bullet = it.marker === "none" ? "none" : it.marker === "number" ? "number" : "char";
    if (para.bullet === "number" && it.start != null) para.number = it.start;
    if (para.bullet !== "none" && it.markerColor) para.bulletColor = colorOf(it.markerColor, style);
    para.marL = u.inch(it.indent);
    para.indent = para.bullet === "none" ? 0 : -Math.min(para.marL, 0.25);
    if (it.lineHeight) para.lineSpacing = u.pt(it.lineHeight);
    if (it.spaceBefore > 0) para.spaceBefore = u.pt(it.spaceBefore);
    const align = ALIGN[it.align];
    if (align && align !== "l") para.align = align;
    para.runs = it.runs.map((r) => runOf(r, u, style, node.opacity ?? 1));
    return para;
  });
  return { kind: "text", ...u.box(node.rect), geometry: "rect", anchor: "t", inset: 0, paras };
}

function shadowOf(s, u, style, opacity) {
  const shadow = {
    blur: u.inch(s.blur),
    dist: u.inch(Math.hypot(s.x, s.y)),
    dir: round(((Math.atan2(s.y, s.x) * 180) / Math.PI + 360) % 360, 2),
    color: colorOf(s.color, style),
  };
  const a = alphaOf(s.color, opacity);
  if (a !== undefined) shadow.alpha = a;
  return shadow;
}

// A border on only some sides (a blockquote's left rule, a heading's underline) is a thin bar.
// CSS border styles PowerPoint draws as a dash pattern.
const DASH = { dashed: "dash", dotted: "sysDot" };

function sideBar(rect, side, u, style, opacity) {
  const t = side.width;
  if (side.dash) {
    // a dashed or dotted side is a line along the side's middle, in that pattern
    const half = t / 2;
    const r =
      side.side === "l" ? { x: rect.x + half, y: rect.y, w: 0, h: rect.h }
      : side.side === "r" ? { x: rect.x + rect.w - half, y: rect.y, w: 0, h: rect.h }
      : side.side === "t" ? { x: rect.x, y: rect.y + half, w: rect.w, h: 0 }
      : { x: rect.x, y: rect.y + rect.h - half, w: rect.w, h: 0 };
    const line = { width: u.pt(t), color: colorOf(side.color, style), dash: DASH[side.dash] };
    const a = alphaOf(side.color, opacity);
    if (a !== undefined) line.alpha = a;
    return { kind: "text", ...u.box(r), geometry: "line", line, paras: [] };
  }
  const r =
    side.side === "l" ? { x: rect.x, y: rect.y, w: t, h: rect.h }
    : side.side === "r" ? { x: rect.x + rect.w - t, y: rect.y, w: t, h: rect.h }
    : side.side === "t" ? { x: rect.x, y: rect.y, w: rect.w, h: t }
    : { x: rect.x, y: rect.y + rect.h - t, w: rect.w, h: t };
  const bar = { kind: "text", ...u.box(r), geometry: "rect", fill: colorOf(side.color, style), paras: [] };
  const a = alphaOf(side.color, opacity);
  if (a !== undefined) bar.fillAlpha = a;
  return bar;
}

const channels = (c) => [0, 2, 4].map((i) => parseInt(c.hex.slice(i, i + 2), 16));
const hexOf = (rgb) => rgb.map((v) => Math.round(v).toString(16).padStart(2, "0")).join("").toUpperCase();

// Source-over: color `top` painted over color `bottom` (each { hex, alpha }).
function composite(top, bottom) {
  const at = top.alpha ?? 1;
  const ab = bottom.alpha ?? 1;
  const a = ab >= 1 ? 1 : at + ab * (1 - at);
  if (!(a > 0)) return { hex: top.hex, alpha: 0 };
  const t = channels(top);
  const b = channels(bottom);
  return { hex: hexOf(t.map((v, i) => (v * at + b[i] * ab * (1 - at)) / a)), alpha: a };
}

// The color u of the way from c1 to c2, mixed the way CSS does (premultiplied by alpha).
function mix(c1, c2, u) {
  const a1 = c1.alpha ?? 1;
  const a2 = c2.alpha ?? 1;
  const a = a1 * (1 - u) + a2 * u;
  if (!(a > 0)) return { hex: c1.hex, alpha: 0 };
  const x = channels(c1);
  const y = channels(c2);
  return { hex: hexOf(x.map((v, i) => (v * a1 * (1 - u) + y[i] * a2 * u) / a)), alpha: a };
}

// Stops (positions ascending) cut to 0..1: the color at the cut is interpolated, the stops
// beyond it dropped - CSS paints only that part of the gradient.
function clipStops(stops) {
  if (stops[0].pos < 0) {
    let i = 0;
    while (i + 1 < stops.length && stops[i + 1].pos <= 0) i++;
    const next = stops[i + 1];
    const color = next ? mix(stops[i].color, next.color, (0 - stops[i].pos) / (next.pos - stops[i].pos)) : stops[i].color;
    stops = [{ pos: 0, color }, ...stops.slice(i + 1)];
  }
  if (stops[stops.length - 1].pos > 1) {
    let j = stops.length - 1;
    while (j > 0 && stops[j - 1].pos >= 1) j--;
    const prev = stops[j - 1];
    const color = prev ? mix(prev.color, stops[j].color, (1 - prev.pos) / (stops[j].pos - prev.pos)) : stops[j].color;
    stops = [...stops.slice(0, j), { pos: 1, color }];
  }
  return stops;
}

// A CSS gradient (and the color under it) -> a DrawingML gradient. Stops are composited over the
// base color (source-over, so a translucent base survives); transparent stops borrow their
// neighbour's hue (a fade to "transparent" would otherwise pass through black). Radial stop
// positions are rescaled from the CSS ellipse to DrawingML's focus-to-farthest-corner radius:
// stops are scaled by sqrt(rx * ry), the radius of the circle with the ellipse's area, over that
// distance - the extractor only passes ellipses within 0.7-1.4 of the box's aspect, so this keeps
// the glow's size right whether the renderer draws DrawingML's circle round or stretched to the
// box. Linear angles move from CSS (0 = up) to DrawingML (0 = left to right).
function gradientOf(g, base, rect, style, opacity = 1) {
  let stops = g.stops.map((s) => ({ pos: s.pos, color: s.color }));
  stops = stops.map((s, i) => {
    if (s.color.alpha > 0) return s;
    const near = stops.slice(i + 1).concat(stops.slice(0, i).reverse()).find((x) => x.color.alpha > 0);
    return near ? { pos: s.pos, color: { hex: near.color.hex, alpha: 0 } } : s;
  });
  if (base && (base.alpha ?? 1) > 0) stops = stops.map((s) => ({ pos: s.pos, color: composite(s.color, base) }));
  let out;
  if (g.type === "radial") {
    const cx = g.cx * rect.w;
    const cy = g.cy * rect.h;
    const far = Math.max(...[[0, 0], [rect.w, 0], [0, rect.h], [rect.w, rect.h]].map(([x, y]) => Math.hypot(x - cx, y - cy)));
    const k = far > 0 ? Math.sqrt(g.rx * (g.ry ?? g.rx)) / far : 1;
    stops = stops.map((s) => ({ ...s, pos: s.pos * k }));
    out = { type: "radial", cx: g.cx, cy: g.cy };
  } else {
    out = { type: "linear", angle: (((g.angle - 90) % 360) + 360) % 360 };
  }
  stops = clipStops(stops);
  // Past either end CSS keeps painting the end color; DrawingML wants stops at 0 and 1.
  const first = stops[0];
  if (first.pos > 0) stops.unshift({ pos: 0, color: first.color });
  const last = stops[stops.length - 1];
  if (last.pos < 1) stops.push({ pos: 1, color: last.color });
  out.stops = stops.map((s) => {
    const stop = { pos: round(s.pos), color: colorOf(s.color, style) };
    const a = alphaOf(s.color, opacity);
    if (a !== undefined) stop.alpha = a;
    return stop;
  });
  return out;
}

function boxShapes(node, u, style) {
  const op = node.opacity ?? 1;
  const shapes = [];
  if (node.fill || node.border || node.shadow) {
    const shape = { kind: "text", ...u.box(node.rect), geometry: node.radius > 0 ? "roundRect" : "rect", paras: [] };
    if (node.radius > 0) shape.radius = u.inch(node.radius);
    if (node.fill?.gradient) shape.fill = gradientOf(node.fill.gradient, node.fill.base, node.rect, style, op);
    else if (node.fill) {
      shape.fill = colorOf(node.fill, style);
      const a = alphaOf(node.fill, op);
      if (a !== undefined) shape.fillAlpha = a;
    }
    if (node.border) {
      shape.line = { width: u.pt(node.border.width), color: colorOf(node.border.color, style), ...(node.border.dash ? { dash: DASH[node.border.dash] } : {}) };
      const a = alphaOf(node.border.color, op);
      if (a !== undefined) shape.line.alpha = a;
    }
    if (node.shadow) shape.shadow = shadowOf(node.shadow, u, style, op);
    if (node.rotate) shape.rot = node.rotate;
    shapes.push(shape);
  }
  for (const side of node.sides || []) shapes.push(sideBar(node.rect, side, u, style, op));
  return shapes;
}

// object-fit: the picture's box and crop for an image of `natural` size drawn into `rect`.
export function fitImage(rect, natural, fit) {
  if (!natural || !natural.w || !natural.h || fit === "fill") return { rect };
  if (fit === "cover") {
    const s = Math.max(rect.w / natural.w, rect.h / natural.h);
    const fx = Math.min(1, rect.w / (natural.w * s));
    const fy = Math.min(1, rect.h / (natural.h * s));
    const crop = { l: round((1 - fx) / 2), t: round((1 - fy) / 2), r: round((1 - fx) / 2), b: round((1 - fy) / 2) };
    return { rect, crop: crop.l || crop.t ? crop : undefined };
  }
  // contain, scale-down and none are all drawn letterboxed inside the box.
  let s = Math.min(rect.w / natural.w, rect.h / natural.h);
  if (fit === "scale-down" || fit === "none") s = Math.min(1, s);
  const w = natural.w * s;
  const h = natural.h * s;
  return { rect: { x: rect.x + (rect.w - w) / 2, y: rect.y + (rect.h - h) / 2, w, h } };
}

function imageShape(node, u, style) {
  const { rect, crop } = fitImage(node.rect, node.natural, node.fit);
  const shape = { kind: "image", ...u.box(rect), data: node.data, ext: node.ext, alt: node.alt || "" };
  if (crop) shape.crop = crop;
  if (node.radius > 0) shape.radius = u.inch(node.radius);
  if (node.svg) shape.svg = node.svg;
  const op = node.opacity ?? 1;
  if (node.border) shape.line = { width: u.pt(node.border.width), color: colorOf(node.border.color, style), ...(node.border.dash ? { dash: DASH[node.border.dash] } : {}) };
  if (node.shadow) shape.shadow = shadowOf(node.shadow, u, style, op);
  if (node.rotate) shape.rot = node.rotate;
  return shape;
}

// A video: placed and cropped as the page draws it (object-fit), its poster as the picture shown
// until it plays, starting by itself, looping and muted as the element does.
function videoShape(node, u) {
  const { rect, crop } = fitImage(node.rect, node.natural, node.fit);
  const shape = {
    kind: "video", ...u.box(rect), data: node.data, ext: node.ext, poster: node.posterImage, alt: node.alt || "",
    autoplay: !!node.autoplay, loop: !!node.loop, muted: !!node.muted, duration: node.duration,
  };
  if (crop) shape.crop = crop;
  return shape;
}

const VALIGN = { middle: "ctr", bottom: "b" };

// A table: one native table with no style, every cell carrying its own fill, borders, margins and
// anchor. A span's start cell carries gridSpan/rowSpan; the slots it covers are merge markers.
function tableShape(node, u, style) {
  const op = node.opacity ?? 1;
  const cellOf = (c) => {
    if (c.covered) return { paras: [], hMerge: c.covered.includes("h"), vMerge: c.covered.includes("v") };
    const para = { runs: c.runs.map((r) => runOf(r, u, style, op)) };
    const align = ALIGN[c.align];
    if (align && align !== "l") para.align = align;
    if (c.lineHeight) para.lineSpacing = u.pt(c.lineHeight);
    const out = {
      paras: [para],
      margin: { l: u.inch(c.padding.l), r: u.inch(c.padding.r), t: u.inch(c.padding.t), b: u.inch(c.padding.b) },
      anchor: VALIGN[c.valign] || "t",
      borders: {},
    };
    for (const [side, b] of Object.entries(c.borders || {})) {
      out.borders[side] = { width: u.pt(b.width), color: colorOf(b.color, style) };
      const a = alphaOf(b.color, op);
      if (a !== undefined) out.borders[side].alpha = a;
    }
    if (c.fill) {
      out.fill = colorOf(c.fill, style);
      const a = alphaOf(c.fill, op);
      if (a !== undefined) out.fillAlpha = a;
    }
    if (c.colspan > 1) out.gridSpan = c.colspan;
    if (c.rowspan > 1) out.rowSpan = c.rowspan;
    return out;
  };
  return {
    kind: "table", ...u.box(node.rect), variant: "plain",
    colW: node.colW.map((w) => u.inch(w)), rowH: node.rowH.map((h) => u.inch(h)),
    rows: node.rows.map((row) => row.map(cellOf)),
  };
}

// A code block's text: one paragraph per source line at the measured line height, never wrapped,
// over the panel box the `pre` itself produced.
function codeShape(node, u, style) {
  const op = node.opacity ?? 1;
  const paras = node.lines.map((runs) => ({ lineSpacing: u.pt(node.lineHeight), runs: runs.map((r) => runOf(r, u, style, op)) }));
  return { kind: "text", ...u.box(node.rect), geometry: "rect", anchor: "t", inset: 0, wrap: false, paras };
}

function chartShape(node, u, style) {
  if (node.resolvedChart) return { kind: "chart", ...u.box(node.rect), chart: node.resolvedChart };
  const c = node.chart;
  const color = (col, fallback) => (col ? colorOf(col, style) : fallback);
  const chart = {
    type: c.type, categories: c.categories,
    series: c.series.map((s, j) => {
      const series = { name: s.name, values: s.values.map((v) => (typeof v === "number" ? v : NaN)), color: color(s.color, PALETTE[j % PALETTE.length]) };
      const a = s.color ? alphaOf(s.color) : undefined;
      if (a !== undefined) series.alpha = a;
      return series;
    }),
    valueAxis: c.valueAxis !== false, legend: c.legend, legendPos: c.legendPos, stacked: c.stacked, gapWidth: 60, background: null,
    catLabelColor: color(c.labelColor, "dk1"), catLabelSize: u.pt(12),
    valLabelColor: color(c.labelColor, "accent3"), valLabelSize: u.pt(12),
    legendColor: color(c.legendColor, "dk1"), legendSize: u.pt(12),
  };
  // Chart.js plots a horizontal chart's first category at the top; the chart part plots it at the bottom.
  if (c.type === "bar") chart.reverseCategories = true;
  if (c.min != null) chart.min = c.min;
  if (c.max != null) chart.max = c.max;
  if (c.numFmt) chart.numFmt = c.numFmt;
  if (c.gridColor) {
    chart.gridColor = colorOf(c.gridColor, style);
    const a = alphaOf(c.gridColor);
    if (a !== undefined) chart.gridAlpha = a;
  }
  return { kind: "chart", ...u.box(node.rect), chart };
}

export function shapesOf(node, u, style) {
  switch (node.kind) {
    case "text":
      return [textShape(node, u, style)];
    case "list":
      return [listShape(node, u, style)];
    case "table":
      return [tableShape(node, u, style)];
    case "code":
      return [codeShape(node, u, style)];
    case "chart":
      return [chartShape(node, u, style)];
    case "box":
      return boxShapes(node, u, style);
    case "image":
      return node.data ? [imageShape(node, u, style)] : [];
    case "raster":
      return node.png ? [{ kind: "image", ...u.box(node.rect), data: node.png, ext: "png", alt: node.alt || "" }] : [];
    case "video":
      return node.data ? [videoShape(node, u)] : [];
    case "missing":
      return [];
    default:
      throw new Error(`Unknown scene node kind ${JSON.stringify(node.kind)}`);
  }
}

function backgroundOf(bg, style, canvas) {
  if (!bg) return null;
  if (bg.png) return { image: { data: bg.png, ext: "png" } };
  if (bg.gradient) return { gradient: gradientOf(bg.gradient, bg.base, { x: 0, y: 0, w: canvas.w, h: canvas.h }, style) };
  if (!bg.color) return null;
  const color = colorOf(bg.color, style);
  const alpha = alphaOf(bg.color);
  // Plain white is the master's background already.
  if (color === "lt1" && alpha === undefined) return null;
  return alpha === undefined ? { color } : { color, alpha };
}

// A title that fits on one line in the browser gets room to spare, so a wider substitute font
// can't wrap it (LibreOffice wraps placeholder text even with wrap="none"): left-aligned titles
// extend to the mirrored right margin, centered ones widen symmetrically, right-aligned ones
// extend to the mirrored left margin.
function titleRectOf(node, canvasW) {
  const r = node.rect;
  if (!node.singleLine) return r;
  if (node.align === "center") {
    const c = r.x + r.w / 2;
    const half = Math.max(r.w / 2, Math.min(c, canvasW - c));
    return { ...r, x: c - half, w: 2 * half };
  }
  if (node.align === "right" || node.align === "end") {
    const right = r.x + r.w;
    const x = Math.min(r.x, canvasW - right);
    return { ...r, x, w: right - x };
  }
  return { ...r, w: Math.max(r.w, canvasW - 2 * r.x) };
}

export function sceneToSpec(scene, style) {
  const u = unitsFor(scene, style.geometry);
  const spec = { layout: "blank", placeholders: {}, shapes: [] };
  if (scene.notes) spec.notes = scene.notes;
  const bg = backgroundOf(scene.background, style, scene.canvas);
  if (bg) spec.background = bg;
  let title = null;
  for (const node of scene.nodes) {
    if (!title && node.kind === "text" && node.role === "title") {
      title = node;
      // The title placeholder is drawn beneath every ordinary shape, so a shape that sits before
      // the title in DOM order (a hero band, a card around the h1) would cover it. titleZ tells
      // the writer how many shapes the title goes above.
      if (spec.shapes.length > 0) spec.placeholders.titleZ = spec.shapes.length;
      continue;
    }
    spec.shapes.push(...shapesOf(node, u, style));
  }
  if (title) {
    spec.placeholders.title = textShape(title, u, style).paras.map((para) => ({ ...para, align: para.align || "l" }));
    spec.placeholders.titleBox = u.box(titleRectOf(title, scene.canvas.w));
    // The box is measured, so none of the master's insets, bottom anchoring or shrink-to-fit
    // apply; a title that fits on one line in the browser must not wrap in PowerPoint.
    const body = { inset: 0, anchor: "t", autofit: "none" };
    spec.placeholders.titleBody = title.singleLine ? { wrap: false, ...body } : body;
  }
  const cover = scene.index === 1 || scene.layout === "cover" || scene.layout === "intro";
  spec.layout = cover ? "title" : title ? "titleOnly" : "blank";
  return spec;
}
