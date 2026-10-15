// Render mode: Scene (what extract.js measured in the browser, in canvas px) -> SlideSpec (what
// the OOXML writer draws, in inches and points). Pure - no browser, no file access - so every
// mapping rule is unit-tested with hand-written scenes.

import { toColor } from "../colors.mjs";

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
function sideBar(rect, side, u, style, opacity) {
  const t = side.width;
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

function boxShapes(node, u, style) {
  const op = node.opacity ?? 1;
  const shapes = [];
  if (node.fill || node.border || node.shadow) {
    const shape = { kind: "text", ...u.box(node.rect), geometry: node.radius > 0 ? "roundRect" : "rect", paras: [] };
    if (node.radius > 0) shape.radius = u.inch(node.radius);
    if (node.fill) {
      shape.fill = colorOf(node.fill, style);
      const a = alphaOf(node.fill, op);
      if (a !== undefined) shape.fillAlpha = a;
    }
    if (node.border) {
      shape.line = { width: u.pt(node.border.width), color: colorOf(node.border.color, style) };
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
  const op = node.opacity ?? 1;
  if (node.border) shape.line = { width: u.pt(node.border.width), color: colorOf(node.border.color, style) };
  if (node.shadow) shape.shadow = shadowOf(node.shadow, u, style, op);
  if (node.rotate) shape.rot = node.rotate;
  return shape;
}

export function shapesOf(node, u, style) {
  switch (node.kind) {
    case "text":
      return [textShape(node, u, style)];
    case "list":
      return [listShape(node, u, style)];
    case "box":
      return boxShapes(node, u, style);
    case "image":
      return node.data ? [imageShape(node, u, style)] : [];
    case "raster":
      return node.png ? [{ kind: "image", ...u.box(node.rect), data: node.png, ext: "png", alt: node.alt || "" }] : [];
    case "missing":
      return [];
    default:
      throw new Error(`Unknown scene node kind ${JSON.stringify(node.kind)}`);
  }
}

function backgroundOf(bg, style) {
  if (!bg) return null;
  if (bg.png) return { image: { data: bg.png, ext: "png" } };
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
  const bg = backgroundOf(scene.background, style);
  if (bg) spec.background = bg;
  let title = null;
  for (const node of scene.nodes) {
    if (!title && node.kind === "text" && node.role === "title") {
      title = node;
      continue;
    }
    spec.shapes.push(...shapesOf(node, u, style));
  }
  if (title) {
    spec.placeholders.title = textShape(title, u, style).paras;
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
