// Markdown -> { globalFrontmatter, slides: [{ eyebrow, title, blocks, notes }] } parsing. Generic
// across any Slidev deck: nothing here is specific to one deck's colors/fonts/components -
// those come in through `config` (see config.mjs) and are only used for cosmetic decisions
// (which CSS var maps to which color, which component tag resolves to which chart data).

import yaml from "js-yaml";

const CODE_TOKEN_RE = /^@@CODEBLOCK_(\d+)@@$/;

// Tailwind/UnoCSS text-size utility classes big enough to mean "this div IS a headline
// number" (the "text-2xl font-bold" stat pattern common in card grids, e.g. "40%" / "Tid")
// rather than ordinary body copy - text-sm/base/lg stay as normal paragraphs so their inline
// **bold**/`code` markdown formatting still goes through mdInlineTokens.
const BIG_TEXT_SIZE_PT = { "text-6xl": 40, "text-5xl": 34, "text-4xl": 28, "text-3xl": 24, "text-2xl": 20 };

// Matches both digit ("grid-cols-3") and arbitrary-value ("grid-cols-[minmax(0,1fr)_2fr]")
// column classes.
const GRID_COLS_RE = /grid-cols-(?:\d+|\[)/;

const LIST_MARKER_RE = /^([-*]|\d+\.)\s+/;
const IMAGE_RE = /^!\[([^\]]*)\]\(([^)]+)\)$/;
const BLOCKQUOTE_RE = /^>\s?(.*)$/;

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// ---------------------------------------------------------------------------
// 1. Split the raw markdown into { globalFrontmatter, slides: [{frontmatter, content}] }
// ---------------------------------------------------------------------------

function looksLikeFrontmatter(chunk) {
  const lines = chunk.split("\n").map((l) => l.trimEnd());
  const nonBlank = lines.filter((l) => l.trim() !== "");
  if (nonBlank.length === 0) return false;
  return nonBlank.every((l) => /^[A-Za-z_][\w-]*:\s*.*$/.test(l) || /^\s+\S/.test(l));
}

export function splitSlides(raw) {
  const text = raw.replace(/\r\n/g, "\n");
  const lines = text.split("\n");
  const rawChunks = [];
  let buf = [];
  for (const line of lines) {
    if (line.trim() === "---") {
      rawChunks.push(buf.join("\n"));
      buf = [];
    } else {
      buf.push(line);
    }
  }
  rawChunks.push(buf.join("\n"));

  let globalFrontmatter = {};
  let pendingFrontmatter = null;
  let sawFirstContent = false;
  const slides = [];

  for (const chunk of rawChunks) {
    if (chunk.trim() === "") continue;
    if (looksLikeFrontmatter(chunk)) {
      const fm = yaml.load(chunk) || {};
      if (!sawFirstContent) globalFrontmatter = { ...globalFrontmatter, ...fm };
      pendingFrontmatter = fm;
      continue;
    }
    sawFirstContent = true;
    slides.push({ frontmatter: pendingFrontmatter || {}, content: chunk });
    pendingFrontmatter = null;
  }

  return { globalFrontmatter, slides };
}

// ---------------------------------------------------------------------------
// 2. Parse one slide's content into { eyebrow, title, blocks[] }
// ---------------------------------------------------------------------------

function findBalancedDiv(str, openTagStart) {
  const openTagEnd = str.indexOf(">", openTagStart) + 1;
  let depth = 1;
  const divTagRe = /<\/?div\b[^>]*>/g;
  divTagRe.lastIndex = openTagEnd;
  let m;
  while ((m = divTagRe.exec(str))) {
    if (m[0].startsWith("</")) depth--;
    else if (!m[0].endsWith("/>")) depth++;
    if (depth === 0) {
      return { inner: str.slice(openTagEnd, m.index), end: m.index + m[0].length };
    }
  }
  return { inner: str.slice(openTagEnd), end: str.length };
}

function parseDivAttrs(str, tagStart) {
  const tagEnd = str.indexOf(">", tagStart) + 1;
  const tagSrc = str.slice(tagStart, tagEnd);
  const classMatch = /class="([^"]*)"/.exec(tagSrc);
  const styleMatch = /style="([^"]*)"/.exec(tagSrc);
  return { cls: classMatch ? classMatch[1] : "", style: styleMatch ? styleMatch[1] : "" };
}

function topLevelDivs(str) {
  const out = [];
  let pos = 0;
  while (true) {
    const nextOpen = str.indexOf("<div", pos);
    if (nextOpen === -1) break;
    const { cls, style } = parseDivAttrs(str, nextOpen);
    const { inner, end } = findBalancedDiv(str, nextOpen);
    out.push({ cls, style, inner });
    pos = end;
  }
  return out;
}

function classToStyle(cls, styleAttr, config) {
  let size = null;
  for (const [token, pt] of Object.entries(BIG_TEXT_SIZE_PT)) {
    if (cls.includes(token)) {
      size = pt;
      break;
    }
  }
  let color = null;
  for (const [cssVar, colorKey] of Object.entries(config.cssVarColors)) {
    if (new RegExp(`\\b${escapeRegExp(cssVar)}\\b`).test(styleAttr || "")) {
      color = config.colors[colorKey];
      break;
    }
  }
  const align = /\btext-center\b/.test(cls) ? "center" : null;
  return { size, bold: true, color, align };
}

// Inline markdown -> tokens. Pure markdown only: what a token *looks* like (colors, fonts) is
// decided by layout.mjs.
export function mdInlineTokens(text) {
  const tokens = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)/g;
  let last = 0;
  let m;
  while ((m = re.exec(text))) {
    if (m.index > last) tokens.push({ text: text.slice(last, m.index) });
    const token = m[0];
    if (token.startsWith("**")) tokens.push({ text: token.slice(2, -2), bold: true });
    else if (token.startsWith("`")) tokens.push({ text: token.slice(1, -1), code: true });
    else tokens.push({ text: token.slice(1, -1), italic: true });
    last = re.lastIndex;
  }
  if (last < text.length) tokens.push({ text: text.slice(last) });
  if (tokens.length === 0) tokens.push({ text });
  return tokens;
}

function parseTable(text) {
  const lines = text.split("\n").map((l) => l.trim()).filter(Boolean);
  const parseRow = (l) => l.replace(/^\|/, "").replace(/\|$/, "").split("|").map((c) => c.trim());
  const rows = [];
  lines.forEach((line, i) => {
    if (i === 1 && /^[\s|:-]+$/.test(line)) return; // header separator row
    rows.push(parseRow(line));
  });
  return rows;
}

// Within one blank-line-separated paragraph, walk line by line so a bold label directly
// followed by a table (no blank line between them) still splits into a paragraph block + a
// proper table block instead of one garbled text blob.

// A markdown list item can wrap onto an unmarked continuation line. Splitting by line
// naively would turn that continuation into its own paragraph; folding it back onto the item
// it continues keeps the list intact. Plain multi-line prose (no markers at all) is
// unaffected - every line already merges into one paragraph either way.
function mergeContinuationLines(lines) {
  const merged = [];
  let inListItem = false;
  for (const line of lines) {
    const isTableRow = /^\|.*\|$/.test(line);
    if (inListItem && !LIST_MARKER_RE.test(line) && !isTableRow) {
      merged[merged.length - 1] += " " + line;
    } else {
      merged.push(line);
      inListItem = LIST_MARKER_RE.test(line);
    }
  }
  return merged;
}

function processPlainText(rawText, blocks) {
  const cleaned = rawText.replace(/<[^>]+>/g, ""); // strip any stray inline tags
  const paras = cleaned.split(/\n\s*\n/);
  for (const paraRaw of paras) {
    const lines = mergeContinuationLines(paraRaw.split("\n").map((l) => l.trim()).filter(Boolean));
    if (lines.length === 0) continue;

    if (lines.length === 1) {
      const codeToken = CODE_TOKEN_RE.exec(lines[0]);
      if (codeToken) {
        blocks.push({ type: "code", idx: Number(codeToken[1]) });
        continue;
      }
      const imageMatch = IMAGE_RE.exec(lines[0]);
      if (imageMatch) {
        blocks.push({ type: "image", alt: imageMatch[1].trim(), src: imageMatch[2].trim() });
        continue;
      }
      // Levels 1/2/3 are handled above this call (title extraction, then ## / ###); a bare
      // "#" heading left in body content by this point is a duplicate H1 further down the
      // slide, treated the same as H2 so it still renders as *something* rather than a
      // literal "# " line.
      const headingMatch = /^(#{1,6})\s+(.+)$/.exec(lines[0]);
      if (headingMatch) {
        blocks.push({ type: "heading", level: Math.max(2, headingMatch[1].length), text: headingMatch[2].trim() });
        continue;
      }
    }

    if (lines.every((l) => BLOCKQUOTE_RE.test(l))) {
      blocks.push({ type: "quote", text: lines.map((l) => BLOCKQUOTE_RE.exec(l)[1]).join(" ").trim() });
      continue;
    }

    let buffer = [];
    let bufferType = null; // 'text' | 'list'
    let bufferOrdered = false;
    const flushBuffer = () => {
      if (buffer.length === 0) return;
      if (bufferType === "list") {
        blocks.push({ type: "bullets", ordered: bufferOrdered, items: buffer.map((l) => l.replace(LIST_MARKER_RE, "")) });
      } else {
        blocks.push({ type: "paragraph", text: buffer.join(" ") });
      }
      buffer = [];
      bufferType = null;
    };

    let tableLines = [];
    const flushTable = () => {
      if (tableLines.length === 0) return;
      blocks.push({ type: "table", rows: parseTable(tableLines.join("\n")) });
      tableLines = [];
    };

    for (const line of lines) {
      if (/^\|.*\|$/.test(line)) {
        flushBuffer();
        tableLines.push(line);
        continue;
      }
      flushTable();
      const isList = LIST_MARKER_RE.test(line);
      const type = isList ? "list" : "text";
      if (bufferType && bufferType !== type) flushBuffer();
      if (bufferType !== "list" && type === "list") bufferOrdered = /^\d+\./.test(line);
      bufferType = type;
      buffer.push(line);
    }
    flushBuffer();
    flushTable();
  }
}

function pushTextBlocks(rawText, blocks, config) {
  let text = rawText
    .replace(/<code\b[^>]*>([\s\S]*?)<\/code>/g, "`$1`")
    .replace(/<strong\b[^>]*>([\s\S]*?)<\/strong>/g, "**$1**")
    .replace(/<br\s*\/?>/g, "\n")
    .replace(/<\/?span[^>]*>/g, "");

  const compRe = /<([A-Z][A-Za-z0-9]*)\b[^>]*\/>|<([A-Z][A-Za-z0-9]*)\b[^>]*>[\s\S]*?<\/\2>/g;
  let lastIdx = 0;
  let cm;
  const segments = [];
  while ((cm = compRe.exec(text))) {
    if (cm.index > lastIdx) segments.push({ type: "text", value: text.slice(lastIdx, cm.index) });
    segments.push({ type: "chart", name: cm[1] || cm[2] });
    lastIdx = compRe.lastIndex;
  }
  if (lastIdx < text.length) segments.push({ type: "text", value: text.slice(lastIdx) });

  for (const seg of segments) {
    if (seg.type === "chart") blocks.push({ type: "chart", name: seg.name, resolved: resolveChart(seg.name, config) });
    else processPlainText(seg.value, blocks);
  }
}

function resolveChart(name, config) {
  const fn = config.chartResolvers[name];
  if (!fn) return null;
  try {
    return fn(config.resolvedDataDir);
  } catch {
    return null;
  }
}

// A "definition list" row - one <div class="grid grid-cols-..."> holding exactly a term
// (often <code>) and a description, e.g. a glossary/API reference row:
//   <div class="grid grid-cols-[...] gap-4 items-center">
//     <code>search_datasets(question, limit=10)</code>
//     <div class="opacity-70">Natural-language search against the catalog...</div>
//   </div>
// Flattened generically this would unwrap into two full-width stacked paragraphs per row (too
// tall, and the <code> term loses its monospace styling once stripped down to plain text) -
// recognized specially here so a whole run of these rows renders as one compact, code-styled,
// native 2-column table instead.
function tryBuildDefListRow(rowInner, config) {
  const rowBlocks = flattenContent(rowInner, config);
  const textOf = (b) => (b.type === "paragraph" || b.type === "styledtext" ? b.text : null);
  if (rowBlocks.length !== 2) return null;
  const term = textOf(rowBlocks[0]);
  const desc = textOf(rowBlocks[1]);
  if (term == null || desc == null) return null;
  return { term, desc };
}

function flattenContent(str, config) {
  const blocks = [];
  let pos = 0;
  while (pos < str.length) {
    const nextOpen = str.indexOf("<div", pos);
    if (nextOpen === -1) {
      pushTextBlocks(str.slice(pos), blocks, config);
      break;
    }
    if (nextOpen > pos) pushTextBlocks(str.slice(pos, nextOpen), blocks, config);

    const { cls, style } = parseDivAttrs(str, nextOpen);
    const { inner, end } = findBalancedDiv(str, nextOpen);

    const isGrid = GRID_COLS_RE.test(cls);
    const isBigStat = !isGrid && !/<div/.test(inner) && Object.keys(BIG_TEXT_SIZE_PT).some((t) => cls.includes(t));
    const children = !isGrid && !isBigStat ? topLevelDivs(inner) : null;
    const isDefList = children && children.length >= 2 && children.every((c) => GRID_COLS_RE.test(c.cls));

    if (isGrid) {
      const cols = topLevelDivs(inner);
      const cards = (cols.length > 0 ? cols : [{ cls: "", style: "", inner }]).map((c) => ({
        blocks: flattenContent(c.inner, config),
      }));
      blocks.push({ type: "cardrow", cards });
    } else if (isBigStat) {
      const plainText = inner
        .replace(/<code\b[^>]*>([\s\S]*?)<\/code>/g, "$1")
        .replace(/<strong\b[^>]*>([\s\S]*?)<\/strong>/g, "$1")
        .replace(/<[^>]+>/g, "")
        .trim();
      if (plainText) blocks.push({ type: "styledtext", text: plainText, style: classToStyle(cls, style, config) });
    } else if (isDefList) {
      const rows = children.map((c) => tryBuildDefListRow(c.inner, config)).filter(Boolean);
      if (rows.length === children.length) blocks.push({ type: "deflist", rows });
      else blocks.push(...flattenContent(inner, config)); // pattern didn't fully match - fall back safely
    } else {
      const innerBlocks = flattenContent(inner, config);
      // Tailwind/UnoCSS opacity-NN fades the whole element toward the page background -
      // stamp that alpha onto the plain-text blocks it wraps so rendering can reproduce the
      // actual faded shade instead of rendering every paragraph at full-strength ink. A block
      // that already has an opacity from a more deeply nested wrapper keeps that value.
      const opacityMatch = /\bopacity-(\d+)\b/.exec(cls);
      if (opacityMatch) {
        const alpha = Number(opacityMatch[1]) / 100;
        for (const b of innerBlocks) {
          if ((b.type === "paragraph" || b.type === "bullets") && b.opacity == null) b.opacity = alpha;
        }
      }
      blocks.push(...innerBlocks);
    }
    pos = end;
  }
  return blocks;
}

export function parseSlide(content, codeStore, config) {
  let text = content;

  // eyebrow label
  let eyebrow = null;
  const eyebrowRe = new RegExp(`<div class="${escapeRegExp(config.eyebrowClass)}">([\\s\\S]*?)</div>`);
  text = text.replace(eyebrowRe, (_, inner) => {
    eyebrow = inner.replace(/<[^>]+>/g, "").trim();
    return "";
  });

  // title (first top-level "# Heading")
  let title = null;
  text = text.replace(/^#[ \t]+(.+)$/m, (_, t) => {
    title = t.trim();
    return "";
  });

  // pull out fenced code blocks so blank-line splitting doesn't break them
  text = text.replace(/```[^\n]*\n([\s\S]*?)```/g, (_, code) => {
    const idx = codeStore.length;
    codeStore.push(code.replace(/\n$/, ""));
    return `\n\n@@CODEBLOCK_${idx}@@\n\n`;
  });

  // Speaker notes: Slidev treats the slide's last HTML comment as its notes when nothing but
  // whitespace follows it. The capture can't contain "<!--", so an earlier comment never
  // swallows the real one. Runs after code fences are tokenized so comments in code are kept.
  let notes = null;
  const notesMatch = /<!--((?:(?!<!--)[\s\S])*?)-->\s*$/.exec(text);
  if (notesMatch) {
    notes = notesMatch[1].trim() || null;
    text = text.slice(0, notesMatch.index);
  }

  const blocks = flattenContent(text, config);
  return { eyebrow, title, blocks, notes };
}
