// Native tables. Styling comes from deck table styles defined in tableStyles.xml with theme
// colors only, and cells carry no direct fills or borders - so PowerPoint's Table Design tab
// and theme recoloring both work. (If Keynote/Google Slides turn out to ignore custom table
// styles, the fallback is to add direct cell borders/fills here.)

import { XML_DECL, NS, guid, emuSize } from "./xml.mjs";
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

export function tableXml(shape, ctx) {
  const cols = Math.max(1, ...shape.rows.map((r) => r.length));
  const widths = shape.colW && shape.colW.length === cols ? shape.colW : Array(cols).fill(shape.w / cols);
  const rowH = emuSize(shape.h / Math.max(1, shape.rows.length));
  const deflist = shape.variant === "deflist";
  const mar = deflist ? [76200, 76200, 50800, 50800] : [91440, 91440, 45720, 45720];
  const tcPr = `<a:tcPr marL="${mar[0]}" marR="${mar[1]}" marT="${mar[2]}" marB="${mar[3]}"${deflist ? ' anchor="ctr"' : ""}/>`;
  const rows = shape.rows
    .map((row) => {
      // Short rows are padded to the grid width - a row with fewer a:tc than gridCols is a
      // repair prompt.
      const cells = Array.from({ length: cols }, (_, i) => row[i] || []);
      // Cell bodies are written like placeholder bodies: a bare <a:bodyPr/>, no wrap/inset attrs.
      return `<a:tr h="${rowH}">${cells.map((paras) => `<a:tc>${txBodyXml(paras, ctx, { placeholder: true }, "a:txBody")}${tcPr}</a:tc>`).join("")}</a:tr>`;
    })
    .join("");
  const styleId = deflist ? TABLE_STYLES.deflist : TABLE_STYLES.default;
  const firstRow = shape.header && !deflist ? ' firstRow="1"' : "";
  const grid = widths.map((w) => `<a:gridCol w="${emuSize(w)}"/>`).join("");
  return `<a:tbl><a:tblPr${firstRow} bandRow="0"><a:tableStyleId>${styleId}</a:tableStyleId></a:tblPr><a:tblGrid>${grid}</a:tblGrid>${rows}</a:tbl>`;
}
