// Native tables. Styling comes from deck table styles defined in tableStyles.xml with theme
// colors only, and cells carry no direct fills or borders - so PowerPoint's Table Design tab
// and theme recoloring both work. (If Keynote/Google Slides turn out to ignore custom table
// styles, the fallback is to add direct cell borders/fills here.) Plain tables (render mode) are
// the exception: they use the built-in No Style table and write each cell's own fill, borders,
// margins and anchor.

import { XML_DECL, NS, guid, emu, emuSize, solidFill } from "./xml.mjs";
import { txBodyXml } from "./text.mjs";

export const TABLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/table";

export const TABLE_STYLES = {
  default: guid("deck-table-default"),
  deflist: guid("deck-table-deflist"),
};

function tcBdr(width) {
  const ln = `<a:ln w="${width}" cmpd="sng"><a:solidFill><a:schemeClr val="accent4"/></a:solidFill></a:ln>`;
  return `<a:tcBdr>${["left", "right", "top", "bottom", "insideH", "insideV"].map((s) => `<a:${s}>${ln}</a:${s}>`).join("")}</a:tcBdr>`;
}

function tcTxStyle(color, bold = false) {
  return `<a:tcTxStyle${bold ? ' b="on"' : ""}><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="${color}"/></a:tcTxStyle>`;
}

export function tableStylesXml() {
  // Header row: panel (bg2) fill with bold accentDark (tx2) text; borders: border color (accent4).
  const deck =
    `<a:tblStyle styleId="${TABLE_STYLES.default}" styleName="Deck Table">` +
    `<a:wholeTbl>${tcTxStyle("tx1")}<a:tcStyle>${tcBdr(12700)}<a:fill><a:noFill/></a:fill></a:tcStyle></a:wholeTbl>` +
    `<a:firstRow>${tcTxStyle("tx2", true)}<a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="bg2"/></a:solidFill></a:fill></a:tcStyle></a:firstRow>` +
    "</a:tblStyle>";
  const deflist =
    `<a:tblStyle styleId="${TABLE_STYLES.deflist}" styleName="Deck Definition List">` +
    `<a:wholeTbl>${tcTxStyle("tx1")}<a:tcStyle>${tcBdr(6350)}<a:fill><a:solidFill><a:schemeClr val="bg2"/></a:solidFill></a:fill></a:tcStyle></a:wholeTbl>` +
    "</a:tblStyle>";
  // def = the deck style, so tables users insert later in PowerPoint match the deck too.
  return `${XML_DECL}<a:tblStyleLst xmlns:a="${NS.a}" def="${TABLE_STYLES.default}">${deck}${deflist}</a:tblStyleLst>`;
}

// Built-in "No Style, No Grid" table style: render mode writes every cell's fill and borders
// itself, so nothing may come from a style.
export const NO_STYLE = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}";

// A side without a border writes no line element at all: an explicit noFill would erase a
// neighbouring cell's rule on a shared edge under border-collapse. A width that rounds to zero
// EMU (a hairline too thin to store) counts as no border.
function lnXml(tag, side) {
  const w = side ? Math.round(side.width * 12700) : 0;
  return w > 0 ? `<a:${tag} w="${w}">${solidFill(side.color, side.alpha)}</a:${tag}>` : "";
}

// A cell is a list of paragraphs (deck-styled tables, formatting from the table style) or, in
// render mode, an object carrying its own fill, borders, margins, anchor and merge flags.
// Cell bodies are written like placeholder bodies: a bare <a:bodyPr/>, no wrap/inset attrs.
function cellXml(cell, ctx, tcPr) {
  const body = (paras) => txBodyXml(paras, ctx, { placeholder: true }, "a:txBody");
  if (Array.isArray(cell)) return `<a:tc>${body(cell)}${tcPr}</a:tc>`;
  const attrs =
    (cell.gridSpan > 1 ? ` gridSpan="${cell.gridSpan}"` : "") +
    (cell.rowSpan > 1 ? ` rowSpan="${cell.rowSpan}"` : "") +
    (cell.hMerge ? ' hMerge="1"' : "") +
    (cell.vMerge ? ' vMerge="1"' : "");
  const m = cell.margin || {};
  const mar = ["l", "r", "t", "b"].map((k) => (m[k] != null ? ` mar${k.toUpperCase()}="${emu(m[k])}"` : "")).join("");
  const anchor = cell.anchor ? ` anchor="${cell.anchor}"` : "";
  const b = cell.borders || {};
  // CT_TableCellProperties child order: lnL, lnR, lnT, lnB, ..., fill.
  const own =
    `<a:tcPr${mar}${anchor}>${lnXml("lnL", b.l)}${lnXml("lnR", b.r)}${lnXml("lnT", b.t)}${lnXml("lnB", b.b)}` +
    `${cell.fill ? solidFill(cell.fill, cell.fillAlpha) : "<a:noFill/>"}</a:tcPr>`;
  return `<a:tc${attrs}>${body(cell.paras || [])}${own}</a:tc>`;
}

export function tableXml(shape, ctx) {
  const cols = Math.max(1, ...shape.rows.map((r) => r.length));
  const widths = shape.colW && shape.colW.length === cols ? shape.colW : Array(cols).fill(shape.w / cols);
  // A row's h is its minimum height (rows still grow to fit their text). `rowH` (inches per
  // row) lets the layout give wrapped rows more room; otherwise the frame height is shared.
  const evenH = shape.h / Math.max(1, shape.rows.length);
  const rowH = (i) => emuSize(shape.rowH && shape.rowH.length === shape.rows.length ? shape.rowH[i] : evenH);
  const deflist = shape.variant === "deflist";
  const plain = shape.variant === "plain";
  const mar = deflist ? [76200, 76200, 50800, 50800] : [91440, 91440, 45720, 45720];
  const tcPr = `<a:tcPr marL="${mar[0]}" marR="${mar[1]}" marT="${mar[2]}" marB="${mar[3]}"${deflist ? ' anchor="ctr"' : ""}/>`;
  const rows = shape.rows
    .map((row, ri) => {
      // Short rows are padded to the grid width - a row with fewer a:tc than gridCols is a
      // repair prompt.
      const cells = Array.from({ length: cols }, (_, i) => row[i] || []);
      return `<a:tr h="${rowH(ri)}">${cells.map((cell) => cellXml(cell, ctx, tcPr)).join("")}</a:tr>`;
    })
    .join("");
  const styleId = plain ? NO_STYLE : deflist ? TABLE_STYLES.deflist : TABLE_STYLES.default;
  const firstRow = shape.header && !deflist && !plain ? ' firstRow="1"' : "";
  const grid = widths.map((w) => `<a:gridCol w="${emuSize(w)}"/>`).join("");
  return `<a:tbl><a:tblPr${firstRow} bandRow="0"><a:tableStyleId>${styleId}</a:tableStyleId></a:tblPr><a:tblGrid>${grid}</a:tblGrid>${rows}</a:tbl>`;
}
