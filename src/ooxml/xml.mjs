// Shared XML primitives for every OOXML part writer: escaping, unit conversion, deterministic
// GUIDs and DrawingML color elements. Nothing here knows about slides or Slidev.

export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export const NS = {
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
  c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
  rels: "http://schemas.openxmlformats.org/package/2006/relationships",
  ct: "http://schemas.openxmlformats.org/package/2006/content-types",
  ssml: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
};

export const NS_PML = `xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"`;

// XML 1.0 allows TAB, LF, CR, U+0020-U+D7FF, U+E000-U+FFFD and U+10000-U+10FFFF. With the `u`
// flag a lone surrogate is its own code point, so it falls outside every allowed range and is
// stripped too - JS strings can carry those (Rust strings can't), and one is enough to make
// PowerPoint reject the whole part.
const INVALID_XML_CHARS = /[^\t\n\r\u0020-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/gu;

export function esc(value) {
  return String(value ?? "")
    .replace(INVALID_XML_CHARS, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

export const EMU_PER_INCH = 914400;

export function emu(inches) {
  const n = Math.round(Number(inches) * EMU_PER_INCH);
  return Number.isFinite(n) ? n : 0;
}

// Extents (a:ext cx/cy) must not be negative - PowerPoint repairs the file if they are.
export function emuSize(inches) {
  return Math.max(0, emu(inches));
}

export function pt100(pt) {
  return Math.round(Number(pt) * 100);
}

function fnv1a(str) {
  let h = 0x811c9dc5;
  for (const ch of str) {
    h ^= ch.codePointAt(0);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h >>> 0;
}

// Content-derived GUID (four salted 32-bit FNV-1a hashes) so table style ids etc. are stable
// across exports - same input, byte-identical file.
export function guid(seed) {
  const hex = [0, 1, 2, 3]
    .map((i) => fnv1a(`${i}:${seed}`).toString(16).padStart(8, "0"))
    .join("")
    .toUpperCase();
  return `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20, 32)}}`;
}

// Theme slots dk1/lt1/dk2/lt2 are referenced from slide content through the master's color
// map (tx1/bg1/tx2/bg2) - that's what PowerPoint itself writes, and what Google Slides and
// Keynote expect. Accents are referenced directly.
const SCHEME_REF = { dk1: "tx1", lt1: "bg1", dk2: "tx2", lt2: "bg2" };
const ACCENTS = new Set(["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]);

export function colorXml(color, alpha) {
  const mod = alpha != null && alpha < 1 ? `<a:alpha val="${Math.round(alpha * 100000)}"/>` : "";
  let open;
  if (typeof color === "string" && /^#[0-9a-fA-F]{6}$/.test(color)) {
    open = `<a:srgbClr val="${color.slice(1).toUpperCase()}"`;
    return mod ? `${open}>${mod}</a:srgbClr>` : `${open}/>`;
  }
  const val = SCHEME_REF[color] || (ACCENTS.has(color) ? color : null);
  if (!val) throw new Error(`Unknown color ${JSON.stringify(color)} (expected a theme slot or #RRGGBB)`);
  open = `<a:schemeClr val="${val}"`;
  return mod ? `${open}>${mod}</a:schemeClr>` : `${open}/>`;
}

export function solidFill(color, alpha) {
  return `<a:solidFill>${colorXml(color, alpha)}</a:solidFill>`;
}

// Linear or radial gradient fill. Stops: [{ pos 0..1, color, alpha? }], at least two. Linear
// angles are DrawingML's (0 = left to right, 90 = top to bottom); a radial gradient runs from
// its focus point (cx, cy as fractions of the shape) at pos 0 out to the shape's edge at pos 1.
export function gradFillXml(g) {
  const pct = (v) => Math.round(Math.min(1, Math.max(0, v)) * 100000);
  const stops = g.stops.map((s) => `<a:gs pos="${pct(s.pos)}">${colorXml(s.color, s.alpha)}</a:gs>`).join("");
  const shade =
    g.type === "radial"
      ? `<a:path path="circle"><a:fillToRect l="${pct(g.cx)}" t="${pct(g.cy)}" r="${pct(1 - g.cx)}" b="${pct(1 - g.cy)}"/></a:path>`
      : `<a:lin ang="${Math.round(((((g.angle ?? 90) % 360) + 360) % 360) * 60000) % 21600000}" scaled="0"/>`;
  return `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst>${shade}</a:gradFill>`;
}

// A shape fill: a color (with optional alpha) or a gradient object.
export function fillXml(fill, alpha) {
  return fill && typeof fill === "object" ? gradFillXml(fill) : solidFill(fill, alpha);
}
