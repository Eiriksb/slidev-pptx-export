// Renders parsed slide data (see parse.mjs) into pptxgenjs shapes. Every color/font/size
// decision here reads from `config` (see config.mjs) - nothing in this file is specific to
// one deck.

import { existsSync } from "node:fs";
import { join } from "node:path";
import { mdInlineToRuns } from "./parse.mjs";

const CARD_SHADOW = { type: "outer", color: "14532D", opacity: 0.16, blur: 9, offset: 3, angle: 90 };

// Applies a CSS-opacity-style fade (composited against a white background) uniformly to
// every run in a text-run array, overriding whatever color mdInlineToRuns already picked
// (plain text, **bold** accent, `code` accent, *italic*) - matching how real CSS opacity
// fades a whole element and its descendants together, not just the plain-text color.
function blendWithWhite(hex, alpha) {
  const n = parseInt(hex, 16);
  const mix = (c) => Math.round(c * alpha + 255 * (1 - alpha));
  const r = mix((n >> 16) & 0xff);
  const g = mix((n >> 8) & 0xff);
  const b = mix(n & 0xff);
  return [r, g, b].map((v) => v.toString(16).padStart(2, "0")).join("").toUpperCase();
}

function fadeRuns(runs, alpha, config) {
  for (const r of runs) r.options.color = blendWithWhite(r.options.color || config.colors.text, alpha);
  return runs;
}

function estimateBlockHeight(block, width, textSize, codeStore, config) {
  const charsPerLine = Math.max(20, Math.floor((width * 72) / (textSize * 0.52)));
  const lineHeight = (textSize + 5) / 72;
  switch (block.type) {
    case "heading":
      return lineHeight * 1.5 + 0.08;
    case "paragraph": {
      const lines = Math.max(1, Math.ceil(block.text.length / charsPerLine));
      return lines * lineHeight + 0.1;
    }
    case "quote":
      return Math.max(1, Math.ceil(block.text.length / charsPerLine)) * lineHeight + 0.16;
    case "bullets": {
      let lines = 0;
      for (const item of block.items) lines += Math.max(1, Math.ceil(item.length / (charsPerLine * 0.9)));
      return lines * lineHeight + 0.1;
    }
    case "code": {
      const codeText = codeStore[block.idx] || "";
      return codeText.split("\n").length * lineHeight * 0.95 + 0.2;
    }
    case "image":
      return 2.4;
    case "table":
      return block.rows.length * 0.32 + 0.15;
    case "deflist": {
      const descColW = width * 0.58;
      const perLine = Math.max(20, Math.floor((descColW * 72) / (textSize * 0.52)));
      let total = 0;
      for (const r of block.rows) total += Math.max(1, Math.ceil(r.desc.length / perLine)) * lineHeight + 0.12;
      return total + 0.05;
    }
    case "styledtext": {
      const size = block.style.size || textSize;
      const perLine = Math.max(20, Math.floor((width * 72) / (size * 0.55)));
      const lines = Math.max(1, Math.ceil(block.text.length / perLine));
      return lines * ((size + 6) / 72) + 0.06;
    }
    case "chart": {
      const kind = block.resolved && block.resolved.kind;
      if (kind === "statcards") return 2.1;
      if (kind === "barchart-panels" || kind === "barchart") return 3.1;
      return 2.1; // unresolved placeholder
    }
    case "cardrow": {
      const n = block.cards.length;
      const gap = 0.18;
      const cardW = (width - gap * (n - 1)) / n;
      const cardTextSize = Math.max(8, textSize - 2);
      const pad = 0.14;
      let maxH = 0.6; // floor so a one-line card isn't razor thin
      for (const card of block.cards) {
        maxH = Math.max(maxH, estimateBlockSequenceHeight(card.blocks, cardW - pad * 2, cardTextSize, codeStore, config) + pad * 2);
      }
      return maxH;
    }
    default:
      return 0.3;
  }
}

// Sum of a run of blocks stacked vertically - used to size a card to its actual content
// instead of a fixed guess (see "cardrow" above and renderBlockSequence below, its rendering
// counterpart).
function estimateBlockSequenceHeight(blocks, width, textSize, codeStore, config) {
  let total = 0;
  for (const b of blocks) total += estimateBlockHeight(b, width, textSize, codeStore, config) + 0.1;
  return total;
}

// statcards: a row of panel cards, each either a small grid of "big number + label" stat
// pairs, or one large centered percentage (+ subtitle).
function renderStatCards(pptx, slide, cards, x, y, w, h, baseTextSize, config) {
  const n = cards.length;
  const gap = 0.2;
  const cardW = (w - gap * (n - 1)) / n;
  cards.forEach((card, i) => {
    const cx = x + i * (cardW + gap);
    const accent = card.accent || config.colors.accent;
    slide.addShape(pptx.ShapeType.roundRect, {
      x: cx, y, w: cardW, h, fill: { color: config.colors.panel }, rectRadius: 0.1, shadow: CARD_SHADOW,
    });
    let cy = y + 0.18;

    if (card.title) {
      const dot = 0.1;
      slide.addShape(pptx.ShapeType.ellipse, {
        x: cx + 0.18, y: cy + 0.03, w: dot, h: dot, fill: { color: accent }, line: { type: "none" },
      });
      slide.addText(card.title, {
        x: cx + 0.18 + dot + 0.09, y: cy - 0.04, w: cardW - 0.36 - dot - 0.09, h: 0.3, fontFace: config.fonts.body,
        fontSize: Math.max(9, baseTextSize - 1), bold: true, color: config.colors.text,
      });
      cy += 0.42;
    }

    if (card.big) {
      slide.addText(String(card.big.value), {
        x: cx, y: cy, w: cardW, h: 0.85, align: "center", valign: "top", fontFace: config.fonts.body,
        fontSize: 38, bold: true, color: card.big.color || config.colors.text,
      });
      cy += 0.85;
      if (card.big.label) {
        slide.addText(card.big.label, {
          x: cx + 0.18, y: cy, w: cardW - 0.36, h: 0.4, align: "center", fontFace: config.fonts.body,
          fontSize: Math.max(9, baseTextSize - 1), color: config.colors.muted,
        });
        cy += 0.42;
      }
    }
    if (card.subtitle) {
      slide.addText(card.subtitle, {
        x: cx + 0.18, y: cy, w: cardW - 0.36, h: 0.32, align: "center", fontFace: config.fonts.body,
        fontSize: Math.max(9, baseTextSize), bold: true, color: config.colors.text,
      });
    }

    if (card.stats && card.stats.length) {
      const cols = card.stats.length >= 4 ? 2 : card.stats.length;
      const rows = Math.ceil(card.stats.length / cols);
      const statW = (cardW - 0.36) / cols;
      const statH = 0.66;
      card.stats.forEach((s, si) => {
        const col = si % cols;
        const row = Math.floor(si / cols);
        const sx = cx + 0.18 + col * statW;
        const sy = cy + row * statH;
        slide.addText(String(s.value), { x: sx, y: sy, w: statW, h: 0.34, fontFace: config.fonts.body, fontSize: 19, bold: true, color: s.color || config.colors.text });
        slide.addText(s.label, { x: sx, y: sy + 0.32, w: statW, h: 0.3, fontFace: config.fonts.body, fontSize: Math.max(8, baseTextSize - 3), color: config.colors.muted });
      });
    }

    if (card.footnote) {
      slide.addText(card.footnote, {
        x: cx + 0.18, y: y + h - 0.34, w: cardW - 0.36, h: 0.3, fontFace: config.fonts.body,
        fontSize: Math.max(8, baseTextSize - 4), color: config.colors.muted,
      });
    }
  });
}

// barchart-panels: one native (editable) horizontal ranked bar chart per panel, side by side.
function renderBarChartPanels(pptx, slide, resolved, x, y, w, h, config) {
  const n = resolved.panels.length;
  const gap = 0.2;
  const panelW = (w - gap * (n - 1)) / n;
  resolved.panels.forEach((panel, i) => {
    const px = x + i * (panelW + gap);
    slide.addShape(pptx.ShapeType.roundRect, {
      x: px, y, w: panelW, h, fill: { color: config.colors.panel }, rectRadius: 0.1, shadow: CARD_SHADOW,
    });
    const dot = 0.1;
    slide.addShape(pptx.ShapeType.ellipse, {
      x: px + 0.18, y: y + 0.16, w: dot, h: dot, fill: { color: panel.color }, line: { type: "none" },
    });
    slide.addText(panel.title, {
      x: px + 0.18 + dot + 0.09, y: y + 0.09, w: panelW - 0.36 - dot - 0.09, h: 0.3, fontFace: config.fonts.body, fontSize: 11, bold: true, color: config.colors.text,
    });
    if (!panel.rows.length) return;
    // PowerPoint bar charts plot category 0 at the bottom; a Chart.js indexAxis:'y' chart
    // (the usual live-preview equivalent) plots it at the top - so reverse the
    // already-sorted rows to keep "highest first" reading top-down.
    const rows = [...panel.rows].reverse();
    slide.addChart(
      pptx.ChartType.bar,
      [{ name: "Value", labels: rows.map((r) => r.model), values: rows.map((r) => r.success) }],
      {
        x: px + 0.12, y: y + 0.52, w: panelW - 0.24, h: h - 0.68,
        barDir: "bar", chartColors: [panel.color],
        // Chunkier "flat design" bars instead of the thin default Office proportions, and
        // the whole value axis dropped since the data labels already say the number - no
        // ruled gridlines, no axis line, no redundant 0/10/20.../100 scale competing with them.
        barGapWidthPct: 35,
        showValue: true, dataLabelColor: panel.color, dataLabelFontSize: 11, dataLabelFontFace: config.fonts.body, dataLabelPosition: "outEnd",
        valAxisHidden: true,
        catAxisLineShow: false, valAxisLineShow: false,
        catGridLine: { style: "none" }, valGridLine: { style: "none" },
        catAxisLabelColor: config.colors.text, catAxisLabelFontSize: 10, catAxisLabelFontFace: config.fonts.body,
        showLegend: false, showTitle: false,
        chartArea: { fill: { color: config.colors.panel } }, plotArea: { fill: { color: config.colors.panel } },
      }
    );
  });
}

// barchart: one native clustered/grouped column chart.
function renderGroupedBarChart(pptx, slide, resolved, x, y, w, h, config) {
  const chartData = resolved.datasets.map((ds) => ({ name: ds.name, labels: resolved.labels, values: ds.data }));
  slide.addChart(pptx.ChartType.bar, chartData, {
    x, y, w, h,
    barDir: "col", barGrouping: "clustered",
    // Thicker, tighter-grouped bars and no gridlines/axis line read as a deliberate
    // flat-design chart instead of an untouched Office default - the value-axis numbers stay
    // (this chart kind's whole point is usually the size of the jump), just without the
    // ruled lines cluttering the plot.
    barGapWidthPct: 45, barOverlapPct: -6,
    chartColors: resolved.datasets.map((d) => d.color),
    showLegend: true, legendPos: "b", legendColor: config.colors.text, legendFontFace: config.fonts.body, legendFontSize: 11,
    showValue: true, dataLabelColor: config.colors.text, dataLabelFontSize: 10, dataLabelFontFace: config.fonts.body, dataLabelPosition: "outEnd",
    valAxisMinVal: resolved.min, valAxisMaxVal: resolved.max,
    catAxisLineShow: false, valAxisLineShow: false,
    catGridLine: { style: "none" }, valGridLine: { style: "none" },
    catAxisLabelColor: config.colors.text, catAxisLabelFontFace: config.fonts.body, catAxisLabelFontSize: 11,
    valAxisLabelColor: config.colors.muted, valAxisLabelFontFace: config.fonts.body, valAxisLabelFontSize: 9,
    chartArea: { fill: { color: "FFFFFF" } }, plotArea: { fill: { color: "FFFFFF" } },
  });
}

// Renders a run of blocks stacked vertically (the main slide body, or the inside of one
// card) using renderBlock for each - the rendering counterpart of
// estimateBlockSequenceHeight above, so what gets measured is what gets drawn.
function renderBlockSequence(pptx, slide, blocks, x, y, w, textSize, codeStore, scale, config) {
  let cy = y;
  for (const block of blocks) {
    const used = renderBlock(pptx, slide, block, x, cy, w, textSize, codeStore, scale, config);
    cy += used + 0.1;
  }
  return cy - y;
}

function renderBlock(pptx, slide, block, x, y, w, textSize, codeStore, scale, config) {
  const h = Math.max(0.25, estimateBlockHeight(block, w, textSize, codeStore, config) * scale);
  switch (block.type) {
    case "heading":
      // Only a slide's actual title (and eyebrow) get the heading font - a markdown ###/##
      // inside body content is never re-styled off the base body font, just bolded.
      slide.addText(block.text, {
        x, y, w, h,
        fontSize: block.level === 2 ? textSize + 5 : textSize + 2,
        bold: true, color: config.colors.accentDark, fontFace: config.fonts.body, valign: "top",
      });
      break;
    case "paragraph": {
      const runs = mdInlineToRuns(block.text, { fontSize: textSize, color: config.colors.text }, config);
      if (block.opacity != null) fadeRuns(runs, block.opacity, config);
      slide.addText(runs, { x, y, w, h, valign: "top" });
      break;
    }
    case "quote": {
      const indent = 0.18;
      slide.addShape(pptx.ShapeType.rect, {
        x, y, w: 0.04, h, fill: { color: config.colors.border }, line: { type: "none" },
      });
      const runs = mdInlineToRuns(block.text, { fontSize: textSize, italic: true, color: config.colors.muted }, config);
      slide.addText(runs, { x: x + indent, y, w: w - indent, h, valign: "top" });
      break;
    }
    case "bullets": {
      const runs = [];
      block.items.forEach((item, i) => {
        // `bullet` is a paragraph-level marker in pptxgenjs: any run that carries it starts
        // a new bulleted/numbered paragraph. An item with inline **bold** splits into several
        // runs (see mdInlineToRuns) - putting `bullet` on all of them would turn one list
        // item into several, renumbering mid-sentence, so only the first run gets it. Each
        // ordered item also needs its own explicit startAt - pptxgenjs defaults it to 1, so
        // without this every item would render as "1." instead of counting up.
        const bulletOpt = block.ordered ? { type: "number", numberType: "arabicPeriod", startAt: i + 1 } : { code: "2022" };
        const r = mdInlineToRuns(item, { fontSize: textSize, color: config.colors.text }, config);
        if (block.opacity != null) fadeRuns(r, block.opacity, config);
        r[0].options.bullet = bulletOpt;
        r[r.length - 1].options.breakLine = true;
        runs.push(...r);
      });
      slide.addText(runs, { x, y, w, h, valign: "top" });
      break;
    }
    case "table": {
      const rows = block.rows.map((row, ri) =>
        row.map((cell) => ({
          text: mdInlineToRuns(cell, ri === 0
            ? { bold: true, color: config.colors.accentDark, fontSize: textSize }
            : { color: config.colors.text, fontSize: textSize }, config),
          options: ri === 0 ? { fill: { color: config.colors.panel } } : {},
        }))
      );
      slide.addTable(rows, { x, y, w, h, border: { type: "solid", color: config.colors.border, pt: 1 }, autoPage: false });
      break;
    }
    case "deflist": {
      // A compact, borderless-look 2-column reference table (term | description) instead of
      // stacking each row's term and description as separate full-width paragraphs - keeps
      // the term's monospace/code styling and takes a fraction of the vertical space.
      const rows = block.rows.map((r) => [
        { text: r.term.replace(/`/g, ""), options: { fontFace: config.fonts.mono, color: config.colors.accentDark, fontSize: Math.max(8, textSize - 1), valign: "middle" } },
        { text: mdInlineToRuns(r.desc, { color: config.colors.muted, fontSize: Math.max(8, textSize - 2) }, config), options: { valign: "middle" } },
      ]);
      slide.addTable(rows, {
        x, y, w, h, colW: [w * 0.4, w * 0.6],
        border: { type: "solid", color: config.colors.border, pt: 0.5 },
        fill: { color: config.colors.panel }, autoPage: false, margin: [4, 6, 4, 6],
      });
      break;
    }
    case "code": {
      const codeText = codeStore[block.idx] || "";
      slide.addText(codeText, {
        x, y, w, h,
        fontFace: config.fonts.mono, fontSize: Math.max(8, textSize - 2),
        color: config.colors.accentDark, fill: { color: config.colors.panel }, valign: "top", margin: 6,
      });
      break;
    }
    case "image": {
      const resolvedSrc = resolveImageSrc(block.src, config);
      if (resolvedSrc) {
        slide.addImage({ path: resolvedSrc, x, y, w, h, sizing: { type: "contain", w, h } });
      } else {
        slide.addText(`[ image not found: ${block.src} ]`, {
          x, y, w, h, align: "center", valign: "middle", italic: true, fontFace: config.fonts.body, fontSize: textSize,
          color: config.colors.muted, fill: { color: config.colors.panel }, line: { color: config.colors.border, dashType: "dash", width: 1 },
        });
      }
      break;
    }
    case "styledtext": {
      const st = block.style;
      const size = Math.max(8, Math.round((st.size || textSize) * scale));
      slide.addText(block.text, {
        x, y, w, h, valign: "top", fontFace: config.fonts.body,
        fontSize: size, bold: st.bold, color: st.color || config.colors.text, align: st.align || "left",
      });
      break;
    }
    case "chart": {
      const resolved = block.resolved;
      if (!resolved) {
        const hint = config.resolvedDataDir ? "no matching data found" : "no data directory configured";
        slide.addText(`[ ${block.name} — interactive chart, ${hint} ]`, {
          x, y, w, h, align: "center", valign: "middle", italic: true, fontFace: config.fonts.body, fontSize: textSize,
          color: config.colors.muted, fill: { color: config.colors.panel }, line: { color: config.colors.border, dashType: "dash", width: 1 },
        });
      } else if (resolved.kind === "statcards") {
        renderStatCards(pptx, slide, resolved.cards, x, y, w, h, textSize, config);
      } else if (resolved.kind === "barchart-panels") {
        renderBarChartPanels(pptx, slide, resolved, x, y, w, h, config);
      } else if (resolved.kind === "barchart") {
        renderGroupedBarChart(pptx, slide, resolved, x, y, w, h, config);
      }
      break;
    }
    case "cardrow": {
      const n = block.cards.length;
      const gap = 0.18;
      const cardW = (w - gap * (n - 1)) / n;
      const cardTextSize = Math.max(8, textSize - 2);
      const pad = 0.14;
      block.cards.forEach((card, i) => {
        const cx = x + i * (cardW + gap);
        slide.addShape(pptx.ShapeType.roundRect, {
          x: cx, y, w: cardW, h,
          fill: { color: config.colors.panel }, line: { color: config.colors.border, width: 1 }, rectRadius: 0.06,
        });
        // Render each card's own blocks (paragraphs, tables, headings, ...) with the real
        // per-type renderer below instead of flattening them to plain text runs - a table
        // inside a card becomes a real aligned pptx table, not joined "a | b" text.
        renderBlockSequence(pptx, slide, card.blocks, cx + pad, y + pad, cardW - pad * 2, cardTextSize, codeStore, scale, config);
      });
      break;
    }
  }
  return h;
}

// Slidev decks reference images either markdown-relative ("./foo.png") or public-dir-rooted
// ("/assets/foo.png", resolved against the deck's public/ folder like a Vite asset). Try both
// conventions since a generic deck may use either.
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

export function renderSlide(pptx, slideData, index, codeStore, config) {
  const slide = pptx.addSlide();
  slide.background = { color: "FFFFFF" };

  if (config.logo && existsSync(config.logo.path)) {
    const logoH = config.logo.height || 0.27;
    slide.addImage({
      path: config.logo.path,
      x: 0.5,
      y: config.slide.height - 0.29 - logoH,
      w: logoH * (config.logo.aspect || 1),
      h: logoH,
    });
  }

  const marginX = config.slide.marginX;
  const contentW = config.slide.width - marginX * 2;

  const isCover = index === 0 || /\bcover\b/.test(slideData.frontmatter.class || "");
  let y = 0.45;

  if (slideData.eyebrow) {
    slide.addText(slideData.eyebrow.toUpperCase(), {
      x: marginX, y, w: contentW, h: 0.3,
      fontSize: 9, bold: true, fontFace: config.fonts.heading, color: config.colors.accent, charSpacing: 1.5,
    });
    y += 0.35;
  }

  if (isCover) {
    slide.addText(slideData.title || "", {
      x: marginX, y: 2.7, w: contentW, h: 1.3,
      fontSize: 40, bold: true, color: config.colors.accent2, align: "center", fontFace: config.fonts.heading, valign: "middle",
    });
    const firstPara = slideData.blocks.find((b) => b.type === "paragraph");
    if (firstPara) {
      slide.addText(firstPara.text, {
        x: marginX, y: 4.1, w: contentW, h: 0.8,
        fontSize: 18, fontFace: config.fonts.body, color: config.colors.text, align: "center",
      });
    }
    return;
  }

  slide.addText(slideData.title || "", {
    x: marginX, y, w: contentW, h: 0.6,
    fontSize: 26, bold: true, color: config.colors.text, fontFace: config.fonts.heading,
  });
  y += 0.8;

  const textSize = slideData.blocks.length > 5 ? 11 : 13;
  const available = config.slide.height - y - 0.35;
  let totalEst = 0;
  for (const b of slideData.blocks) totalEst += estimateBlockHeight(b, contentW, textSize, codeStore, config) + 0.14;
  const scale = totalEst > available && totalEst > 0 ? Math.max(0.55, available / totalEst) : 1;
  const gap = 0.16 * scale;
  const drawTextSize = Math.max(9, Math.round(textSize * Math.min(1, scale + 0.3)));

  for (const block of slideData.blocks) {
    const used = renderBlock(pptx, slide, block, marginX, y, contentW, drawTextSize, codeStore, scale, config);
    y += used + gap;
  }
}
