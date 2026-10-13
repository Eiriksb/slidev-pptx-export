// theme1.xml from the deck's theme slots and fonts. The format scheme is the minimal valid one
// (each list needs exactly three entries); shapes here set their own fills/lines, so it only
// matters for shapes users insert later in PowerPoint.

import { XML_DECL, NS, esc } from "./xml.mjs";

const SLOTS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

const LINE = (w) =>
  `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;

const FMT_SCHEME =
  '<a:fmtScheme name="Deck">' +
  "<a:fillStyleLst>" +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:solidFill><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:solidFill>' +
  '<a:solidFill><a:schemeClr val="phClr"><a:shade val="80000"/></a:schemeClr></a:solidFill>' +
  "</a:fillStyleLst>" +
  `<a:lnStyleLst>${LINE(6350)}${LINE(12700)}${LINE(19050)}</a:lnStyleLst>` +
  "<a:effectStyleLst>" +
  "<a:effectStyle><a:effectLst/></a:effectStyle>".repeat(3) +
  "</a:effectStyleLst>" +
  "<a:bgFillStyleLst>" +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/></a:schemeClr></a:solidFill>' +
  '<a:solidFill><a:schemeClr val="phClr"><a:shade val="90000"/></a:schemeClr></a:solidFill>' +
  "</a:bgFillStyleLst>" +
  "</a:fmtScheme>";

export function themeXml({ name = "Deck", colors, fonts }) {
  const clr = SLOTS.map((s) => `<a:${s}><a:srgbClr val="${colors[s]}"/></a:${s}>`).join("");
  const font = (typeface) => `<a:latin typeface="${esc(typeface)}"/><a:ea typeface=""/><a:cs typeface=""/>`;
  return (
    `${XML_DECL}<a:theme xmlns:a="${NS.a}" name="${esc(name)}"><a:themeElements>` +
    `<a:clrScheme name="${esc(name)}">${clr}</a:clrScheme>` +
    `<a:fontScheme name="${esc(name)}"><a:majorFont>${font(fonts.heading)}</a:majorFont><a:minorFont>${font(fonts.body)}</a:minorFont></a:fontScheme>` +
    FMT_SCHEME +
    "</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>"
  );
}
