// Turns parsed slides (see parse.mjs) into SlideSpecs for the OOXML writer in ./ooxml/.
// Titles always go into the layout's title placeholder; text-only bodies (paragraphs, bullets,
// sub-headings) go into its body placeholder, so they behave like hand-made PowerPoint slides.
// Everything else is drawn as positioned shapes, using the geometry the pptxgenjs renderer
// used. Every color/font decision reads from `config` - nothing here is specific to one deck.

import { existsSync, readFileSync } from "node:fs";
import { join, extname } from "node:path";
import { imageSize } from "image-size";
import { mdInlineTokens } from "./parse.mjs";
import { themeColors, toColor, hexOf, blendWithWhite } from "./colors.mjs";
import { slideGeometry } from "./geometry.mjs";

// The master's level-1 body size (ooxml/master.mjs) - placeholder text inherits it.
const BODY_PT = 13;
const TEXT_ONLY = new Set(["paragraph", "bullets", "heading"]);
const IMAGE_EXTS = new Set([".png", ".jpg", ".jpeg", ".gif"]);
// Series colors for resolver datasets/panels that don't name one.
const PALETTE = ["accent1", "accent2", "accent5", "accent6", "accent3"];

export function deckStyle(config) {
  const theme = themeColors(config.colors);
  const slot = (key) => toColor(config.colors[key], theme);
  return {
    config,
    theme,
    geometry: slideGeometry(config.slide),
    color: {
      text: slot("text"), accentDark: slot("accentDark"), panel: slot("panel"), accent: slot("accent"),
      accent2: slot("accent2"), muted: slot("muted"), border: slot("border"),
    },
    hex: (h) => toColor(h, theme),
  };
}

// ---------------------------------------------------------------------------
// Height estimation (unchanged from the pptxgenjs renderer)
// ---------------------------------------------------------------------------

function estimateBlockHeight(block, width, textSize, codeStore) {
  const charsPerLine = Math.max(20, Math.floor((width * 72) / (textSize * 0.52)));
  const lineHeight = (textSize + 5) / 72;
  switch (block.type) {
    case "heading":
      return lineHeight * 1.5 + 0.08;
    case "paragraph":
      return Math.max(1, Math.ceil(block.text.length / charsPerLine)) * lineHeight + 0.1;
    case "quote":
      return Math.max(1, Math.ceil(block.text.length / charsPerLine)) * lineHeight + 0.16;
    case "bullets": {
      let lines = 0;
      for (const item of block.items) lines += Math.max(1, Math.ceil(item.length / (charsPerLine * 0.9)));
      return lines * lineHeight + 0.1;
    }
    case "code":
      return (codeStore[block.idx] || "").split("\n").length * lineHeight * 0.95 + 0.2;
    case "image":
      return 2.4;
    case "table":
      return block.rows.length * 0.32 + 0.15;
    case "deflist": {
      const perLine = Math.max(20, Math.floor((width * 0.58 * 72) / (textSize * 0.52)));
      let total = 0;
      for (const r of block.rows) total += Math.max(1, Math.ceil(r.desc.length / perLine)) * lineHeight + 0.12;
      return total + 0.05;
    }
    case "styledtext": {
      const size = block.style.size || textSize;
      const perLine = Math.max(20, Math.floor((width * 72) / (size * 0.55)));
      return Math.max(1, Math.ceil(block.text.length / perLine)) * ((size + 6) / 72) + 0.06;
    }
    case "chart": {
      const kind = block.resolved && block.resolved.kind;
      if (kind === "barchart-panels" || kind === "barchart") return 3.1;
      return 2.1; // statcards, or the unresolved placeholder
    }
    case "cardrow": {
      const n = block.cards.length;
      const cardW = (width - 0.18 * (n - 1)) / n;
      let maxH = 0.6; // floor so a one-line card isn't razor thin
      for (const card of block.cards) {
        maxH = Math.max(maxH, blocksHeight(card.blocks, cardW - 0.28, Math.max(8, textSize - 2), codeStore, 0.1) + 0.28);
      }
      return maxH;
    }
    default:
      return 0.3;
  }
}

function blocksHeight(blocks, width, textSize, codeStore, gap = 0.14) {
  let total = 0;
  for (const b of blocks) total += estimateBlockHeight(b, width, textSize, codeStore) + gap;
  return total;
}

// PowerPoint-style shrink: 2.5% steps, never below 55% (same floor as the free-shape scaling).
function shrinkScale(available, needed) {
  return Math.max(0.55, Math.floor((available / needed) * 40) / 40);
}

// ---------------------------------------------------------------------------
// Text
// ---------------------------------------------------------------------------

function inlineRuns(text, base, S) {
  return mdInlineTokens(text).map((t) => {
    if (t.bold) return { ...base, text: t.text, b: true, color: S.color.accentDark };
    if (t.code) return { ...base, text: t.text, font: "mono", color: S.color.accentDark };
    if (t.italic) return { ...base, text: t.text, i: true };
    return { ...base, text: t.text };
  });
}

// Tailwind/UnoCSS opacity-NN fades the whole element, including its bold/code runs.
function fadeRuns(runs, alpha, S) {
  return runs.map((r) => ({ ...r, color: `#${blendWithWhite(hexOf(r.color || S.color.text, S.theme), alpha)}` }));
}

// Paragraphs for text-only blocks. In a placeholder (`size` null) plain runs inherit the
// master's size and color; in a free text box they carry them explicitly.
function blockParas(blocks, size, S, align) {
  const base = size ? { font: "body", size, color: S.color.text } : { font: "body" };
  const paras = [];
  blocks.forEach((block, bi) => {
    const first = paras.length;
    const runsFor = (text) => {
      const runs = inlineRuns(text, base, S);
      return block.opacity != null ? fadeRuns(runs, block.opacity, S) : runs;
    };
    if (block.type === "paragraph") {
      paras.push({ bullet: "none", align, runs: runsFor(block.text) });
    } else if (block.type === "bullets") {
      for (const item of block.items) paras.push({ bullet: block.ordered ? "number" : "char", align, runs: runsFor(item) });
    } else if (block.type === "heading") {
      // Only the slide title uses the heading font - a ##/### inside the body is bold body text.
      const hs = (size || BODY_PT) + (block.level === 2 ? 5 : 2);
      paras.push({ bullet: "none", align, runs: [{ text: block.text, font: "body", b: true, size: hs, color: S.color.accentDark }] });
    }
    // Block-to-block gap (the old renderer's 0.16") as space before the block's first paragraph.
    if (bi > 0 && paras[first]) paras[first].spaceBefore = 12;
  });
  return paras;
}

function eyebrowShape(text, box, align, S) {
  return {
    kind: "text", ...box, geometry: "rect", anchor: "ctr",
    paras: [{ align, runs: [{ text: text.toUpperCase(), font: "heading", size: 9, b: true, color: S.color.accent, spacing: 1.5 }] }],
  };
}

function placeholderBox(box, label, textSize, S) {
  return {
    kind: "text", ...box, geometry: "rect", anchor: "ctr", fill: S.color.panel,
    line: { color: S.color.border, width: 1, dash: "dash" },
    paras: [{ align: "ctr", runs: [{ text: label, font: "body", size: textSize, i: true, color: S.color.muted }] }],
  };
}

// ---------------------------------------------------------------------------
// Images
// ---------------------------------------------------------------------------

// Slidev decks reference images either markdown-relative ("./foo.png") or public-dir-rooted
// ("/assets/foo.png", resolved against the deck's public/ folder like a Vite asset).
function resolveImageSrc(src, config) {
  if (/^https?:\/\//.test(src)) return null; // no network fetches during export
  const candidates = [];
  if (src.startsWith("/")) {
    if (config.publicDir) candidates.push(join(config.publicDir, src.slice(1)));
  } else {
    if (config.inputDir) candidates.push(join(config.inputDir, src));
    if (config.publicDir) candidates.push(join(config.publicDir, src));
  }
  return candidates.find((p) => existsSync(p)) || null;
}

// "contain" inside the box, anchored top-left like an inline Slidev image.
function fitContain(box, width, height) {
  if (!(width > 0 && height > 0)) return { ...box };
  const s = Math.min(box.w / width, box.h / height);
  return { x: box.x, y: box.y, w: width * s, h: height * s };
}

function imageShape(block, box, textSize, S) {
  const src = resolveImageSrc(block.src, S.config);
  if (!src) return placeholderBox(box, `[ image not found: ${block.src} ]`, textSize, S);
  if (!IMAGE_EXTS.has(extname(src).toLowerCase())) return placeholderBox(box, `[ unsupported image format: ${block.src} ]`, textSize, S);
  let dims;
  try {
    dims = imageSize(readFileSync(src));
  } catch {
    return placeholderBox(box, `[ unreadable image: ${block.src} ]`, textSize, S);
  }
  return { kind: "image", ...fitContain(box, dims.width, dims.height), path: src, alt: block.alt };
}

// ---------------------------------------------------------------------------
// Charts (resolver output -> native shapes/charts)
// ---------------------------------------------------------------------------

function label(text, box, size, { align, anchor = "ctr", ...run } = {}) {
  return { kind: "text", ...box, geometry: "rect", anchor, paras: [{ align, runs: [{ text: String(text), font: "body", size, ...run }] }] };
}

// statcards: a row of panel cards, each either a small grid of "big number + label" stat
// pairs, or one large centered percentage (+ subtitle).
function statCardShapes(cards, { x, y, w, h }, baseTextSize, S) {
  const { color } = S;
  const shapes = [];
  const gap = 0.2;
  const cardW = (w - gap * (cards.length - 1)) / cards.length;
  cards.forEach((card, i) => {
    const cx = x + i * (cardW + gap);
    shapes.push({ kind: "text", x: cx, y, w: cardW, h, geometry: "roundRect", radius: 0.1, fill: color.panel, shadow: true, paras: [] });
    let cy = y + 0.18;
    if (card.title) {
      const dot = 0.1;
      shapes.push({ kind: "text", x: cx + 0.18, y: cy + 0.03, w: dot, h: dot, geometry: "ellipse", fill: card.accent ? S.hex(card.accent) : color.accent, paras: [] });
      shapes.push(label(card.title, { x: cx + 0.18 + dot + 0.09, y: cy - 0.04, w: cardW - 0.36 - dot - 0.09, h: 0.3 }, Math.max(9, baseTextSize - 1), { b: true, color: color.text }));
      cy += 0.42;
    }
    if (card.big) {
      shapes.push(label(card.big.value, { x: cx, y: cy, w: cardW, h: 0.85 }, 38, { align: "ctr", anchor: "t", b: true, color: card.big.color ? S.hex(card.big.color) : color.text }));
      cy += 0.85;
      if (card.big.label) {
        shapes.push(label(card.big.label, { x: cx + 0.18, y: cy, w: cardW - 0.36, h: 0.4 }, Math.max(9, baseTextSize - 1), { align: "ctr", color: color.muted }));
        cy += 0.42;
      }
    }
    if (card.subtitle) {
      shapes.push(label(card.subtitle, { x: cx + 0.18, y: cy, w: cardW - 0.36, h: 0.32 }, Math.max(9, baseTextSize), { align: "ctr", b: true, color: color.text }));
    }
    if (card.stats && card.stats.length) {
      const cols = card.stats.length >= 4 ? 2 : card.stats.length;
      const statW = (cardW - 0.36) / cols;
      card.stats.forEach((s, si) => {
        const sx = cx + 0.18 + (si % cols) * statW;
        const sy = cy + Math.floor(si / cols) * 0.66;
        shapes.push(label(s.value, { x: sx, y: sy, w: statW, h: 0.34 }, 19, { b: true, color: s.color ? S.hex(s.color) : color.text }));
        shapes.push(label(s.label, { x: sx, y: sy + 0.32, w: statW, h: 0.3 }, Math.max(8, baseTextSize - 3), { color: color.muted }));
      });
    }
    if (card.footnote) {
      shapes.push(label(card.footnote, { x: cx + 0.18, y: y + h - 0.34, w: cardW - 0.36, h: 0.3 }, Math.max(8, baseTextSize - 4), { color: color.muted }));
    }
  });
  return shapes;
}

// barchart-panels: one native horizontal ranked bar chart per panel, side by side.
function barPanelShapes(resolved, { x, y, w, h }, S) {
  const { color } = S;
  const shapes = [];
  const gap = 0.2;
  const panelW = (w - gap * (resolved.panels.length - 1)) / resolved.panels.length;
  resolved.panels.forEach((panel, i) => {
    const px = x + i * (panelW + gap);
    const pc = panel.color ? S.hex(panel.color) : PALETTE[i % PALETTE.length];
    const dot = 0.1;
    shapes.push({ kind: "text", x: px, y, w: panelW, h, geometry: "roundRect", radius: 0.1, fill: color.panel, shadow: true, paras: [] });
    shapes.push({ kind: "text", x: px + 0.18, y: y + 0.16, w: dot, h: dot, geometry: "ellipse", fill: pc, paras: [] });
    shapes.push(label(panel.title, { x: px + 0.18 + dot + 0.09, y: y + 0.09, w: panelW - 0.36 - dot - 0.09, h: 0.3 }, 11, { b: true, color: color.text }));
    if (!panel.rows.length) return;
    // PowerPoint bar charts plot category 0 at the bottom; a Chart.js indexAxis:'y' chart (the
    // usual live-preview equivalent) plots it at the top - so reverse the already-sorted rows
    // to keep "highest first" reading top-down.
    const rows = [...panel.rows].reverse();
    shapes.push({
      kind: "chart", x: px + 0.12, y: y + 0.52, w: panelW - 0.24, h: h - 0.68,
      chart: {
        type: "bar", categories: rows.map((r) => String(r.model)),
        series: [{ name: "Value", color: pc, values: rows.map((r) => Number(r.success)) }],
        // The data labels already say the number, so the value axis is dropped entirely.
        valueAxis: false, legend: false, gapWidth: 35,
        dataLabels: { color: pc, size: 11 }, background: color.panel, catLabelColor: color.text, catLabelSize: 10,
      },
    });
  });
  return shapes;
}

// barchart: one native clustered column chart.
function groupedBarShape(resolved, box, S) {
  const { color } = S;
  return {
    kind: "chart", ...box,
    chart: {
      type: "column", categories: resolved.labels.map(String),
      series: resolved.datasets.map((d, j) => ({ name: String(d.name), color: d.color ? S.hex(d.color) : PALETTE[j % PALETTE.length], values: d.data.map(Number) })),
      min: resolved.min, max: resolved.max, valueAxis: true, legend: true, gapWidth: 45, overlap: -6,
      dataLabels: { color: color.text, size: 10 }, background: "lt1",
      catLabelColor: color.text, catLabelSize: 11, valLabelColor: color.muted, valLabelSize: 9, legendColor: color.text, legendSize: 11,
    },
  };
}

function chartShapes(block, box, textSize, S) {
  const r = block.resolved;
  if (!r) {
    const hint = S.config.resolvedDataDir ? "no matching data found" : "no data directory configured";
    return [placeholderBox(box, `[ ${block.name} — interactive chart, ${hint} ]`, textSize, S)];
  }
  if (r.kind === "statcards") return statCardShapes(r.cards, box, textSize, S);
  if (r.kind === "barchart-panels") return barPanelShapes(r, box, S);
  if (r.kind === "barchart") return [groupedBarShape(r, box, S)];
  return [placeholderBox(box, `[ ${block.name} — unknown chart kind "${r.kind}" ]`, textSize, S)];
}

// ---------------------------------------------------------------------------
// Positioned (free) shapes for rich bodies
// ---------------------------------------------------------------------------

function blockShapes(block, x, y, w, textSize, codeStore, scale, S) {
  const { color } = S;
  const h = Math.max(0.25, estimateBlockHeight(block, w, textSize, codeStore) * scale);
  const box = { x, y, w, h };
  const text = (paras, extra = {}) => ({ kind: "text", ...box, geometry: "rect", anchor: "t", paras, ...extra });
  let shapes;
  switch (block.type) {
    case "heading":
    case "paragraph":
    case "bullets":
      shapes = [text(blockParas([block], textSize, S))];
      break;
    case "quote": {
      const indent = 0.18;
      shapes = [
        { kind: "text", x, y, w: 0.04, h, geometry: "rect", fill: color.border, paras: [] },
        { ...text([{ runs: inlineRuns(block.text, { font: "body", size: textSize, i: true, color: color.muted }, S) }]), x: x + indent, w: w - indent },
      ];
      break;
    }
    case "table":
      // Header bold/color comes from the deck table style, so cells only carry size + inline markdown.
      shapes = [{ kind: "table", ...box, header: true, variant: "default", rows: block.rows.map((row) => row.map((c) => [{ runs: inlineRuns(c, { font: "body", size: textSize }, S) }])) }];
      break;
    case "deflist":
      // A compact 2-column reference table (term | description) instead of stacked paragraphs.
      shapes = [{
        kind: "table", ...box, colW: [w * 0.4, w * 0.6], header: false, variant: "deflist",
        rows: block.rows.map((r) => [
          [{ runs: [{ text: r.term.replace(/`/g, ""), font: "mono", size: Math.max(8, textSize - 1), color: color.accentDark }] }],
          [{ runs: inlineRuns(r.desc, { font: "body", size: Math.max(8, textSize - 2), color: color.muted }, S) }],
        ]),
      }];
      break;
    case "code": {
      const size = Math.max(8, textSize - 2);
      const lines = (codeStore[block.idx] || "").split("\n");
      shapes = [text(lines.map((line) => ({ runs: [{ text: line, font: "mono", size, color: color.accentDark }] })), { fill: color.panel, inset: 6 / 72 })];
      break;
    }
    case "image":
      shapes = [imageShape(block, box, textSize, S)];
      break;
    case "styledtext": {
      const st = block.style;
      const size = Math.max(8, Math.round((st.size || textSize) * scale));
      shapes = [text([{ align: st.align === "center" ? "ctr" : undefined, runs: [{ text: block.text, font: "body", size, b: st.bold, color: st.color ? S.hex(st.color) : color.text }] }])];
      break;
    }
    case "chart":
      shapes = chartShapes(block, box, textSize, S);
      break;
    case "cardrow": {
      const gap = 0.18;
      const pad = 0.14;
      const cardW = (w - gap * (block.cards.length - 1)) / block.cards.length;
      shapes = [];
      block.cards.forEach((card, i) => {
        const cx = x + i * (cardW + gap);
        shapes.push({ kind: "text", x: cx, y, w: cardW, h, geometry: "roundRect", radius: 0.06, fill: color.panel, line: { color: color.border, width: 1 }, paras: [] });
        // Each card's own blocks get the real per-type rendering - a table inside a card is a
        // real table, not joined "a | b" text.
        shapes.push(...sequenceShapes(card.blocks, cx + pad, y + pad, cardW - pad * 2, Math.max(8, textSize - 2), codeStore, scale, S, 0.1).shapes);
      });
      break;
    }
    default:
      shapes = [];
  }
  return { h, shapes };
}

function sequenceShapes(blocks, x, y, w, textSize, codeStore, scale, S, gap) {
  let cy = y;
  const shapes = [];
  for (const block of blocks) {
    const r = blockShapes(block, x, cy, w, textSize, codeStore, scale, S);
    shapes.push(...r.shapes);
    cy += r.h + gap;
  }
  return { shapes, h: cy - y };
}

// The old renderer's whole-body flow: pick a text size, scale everything down (to 55% at most)
// if the estimate doesn't fit, then stack the blocks.
function freeBody(blocks, top, available, codeStore, S) {
  const g = S.geometry;
  const textSize = blocks.length > 5 ? 11 : 13;
  const total = blocksHeight(blocks, g.contentW, textSize, codeStore);
  const scale = total > available && total > 0 ? Math.max(0.55, available / total) : 1;
  const drawTextSize = Math.max(9, Math.round(textSize * Math.min(1, scale + 0.3)));
  return sequenceShapes(blocks, g.marginX, top, g.contentW, drawTextSize, codeStore, scale, S, 0.16 * scale);
}

// ---------------------------------------------------------------------------
// Slides
// ---------------------------------------------------------------------------

function layoutCentered(slide, title, textOnly, align, spec, codeStore, S) {
  const g = S.geometry;
  const blocks = slide.blocks;
  const gap = 0.2;
  const eyebrowH = slide.eyebrow ? 0.35 : 0;
  spec.layout = "centered";
  // Room for the body if the stack were pinned to the top margin.
  const maxBodyH = g.height - g.bottomMargin - (g.eyebrow.y + eyebrowH + g.title.h + gap);
  let bodyH = 0;
  let free = null;
  if (textOnly) bodyH = blocksHeight(blocks, g.contentW, BODY_PT, codeStore);
  else if (blocks.length) {
    free = freeBody(blocks, 0, maxBodyH, codeStore, S);
    bodyH = free.h;
  }
  const stackH = eyebrowH + g.title.h + (blocks.length ? gap + Math.min(bodyH, maxBodyH) : 0);
  const top = Math.max(g.eyebrow.y, (g.height - stackH) / 2);

  if (slide.eyebrow) spec.shapes.push(eyebrowShape(slide.eyebrow, { ...g.eyebrow, y: top }, align, S));
  spec.placeholders.title = title;
  spec.placeholders.titleBox = { ...g.title, y: top + eyebrowH };
  const bodyTop = spec.placeholders.titleBox.y + g.title.h + gap;
  if (textOnly) {
    const room = g.height - g.bottomMargin - bodyTop;
    spec.placeholders.body = blockParas(blocks, null, S, align);
    spec.placeholders.bodyBox = { ...g.body, y: bodyTop, h: Math.max(0.5, Math.min(bodyH, room)) };
    if (bodyH > room) spec.placeholders.bodyFontScale = shrinkScale(room, bodyH);
  } else if (free) {
    spec.shapes.push(...free.shapes.map((s) => ({ ...s, y: s.y + bodyTop })));
  }
  return spec;
}

function layoutSlide(slide, index, codeStore, S) {
  const g = S.geometry;
  const fm = slide.frontmatter || {};
  const cls = String(fm.class || "");
  const align = /\btext-center\b/.test(cls) ? "ctr" : undefined;
  const blocks = slide.blocks;
  const spec = { layout: "blank", placeholders: {}, shapes: [] };
  if (slide.notes) spec.notes = slide.notes;
  const title = slide.title ? [{ align, runs: [{ text: slide.title, font: "heading" }] }] : null;

  if (index === 0 || /\bcover\b/.test(cls) || fm.layout === "cover" || fm.layout === "intro") {
    spec.layout = "title";
    if (slide.eyebrow) spec.shapes.push(eyebrowShape(slide.eyebrow, g.eyebrow, undefined, S));
    if (title) spec.placeholders.title = title;
    const first = blocks.find((b) => b.type === "paragraph");
    if (first) spec.placeholders.subtitle = [{ runs: inlineRuns(first.text, { font: "body" }, S) }];
    return spec;
  }

  if (!title) {
    // Nothing for a title placeholder: everything is drawn as shapes on the Blank layout.
    if (slide.eyebrow) spec.shapes.push(eyebrowShape(slide.eyebrow, g.eyebrow, align, S));
    const top = g.title.y;
    spec.shapes.push(...freeBody(blocks, top, g.height - g.bottomMargin - top, codeStore, S).shapes);
    return spec;
  }

  const textOnly = blocks.length > 0 && blocks.every((b) => TEXT_ONLY.has(b.type));
  if (fm.layout === "center" || fm.layout === "end") return layoutCentered(slide, title, textOnly, align, spec, codeStore, S);

  if (fm.layout === "section" && (textOnly || blocks.length === 0)) {
    spec.layout = "section";
    spec.placeholders.title = title;
    if (slide.eyebrow) spec.shapes.push(eyebrowShape(slide.eyebrow, { ...g.eyebrow, y: g.sectionTitle.y - 0.35 }, align, S));
    if (textOnly) spec.placeholders.body = blockParas(blocks, null, S, align);
    return spec;
  }

  if (slide.eyebrow) spec.shapes.push(eyebrowShape(slide.eyebrow, g.eyebrow, align, S));
  spec.placeholders.title = title;
  if (textOnly) {
    spec.layout = "content";
    spec.placeholders.body = blockParas(blocks, null, S, align);
    const needed = blocksHeight(blocks, g.body.w, BODY_PT, codeStore);
    if (needed > g.body.h) spec.placeholders.bodyFontScale = shrinkScale(g.body.h, needed);
    return spec;
  }
  spec.layout = "titleOnly";
  if (blocks.length) spec.shapes.push(...freeBody(blocks, g.body.y, g.body.h, codeStore, S).shapes);
  return spec;
}

export function layoutDeck(parsedSlides, codeStore, config) {
  const S = deckStyle(config);
  return parsedSlides.map((slide, i) => layoutSlide(slide, i, codeStore, S));
}
