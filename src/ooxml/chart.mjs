// Clustered bar/column charts (c:chartSpace). Every series carries strCache/numCache - that's
// what Google Slides and Keynote render from - plus c:f formulas into the embedded workbook
// (xlsx.mjs) so "Edit Data" works in PowerPoint. Exactly two axes per bar chart (pptxgenjs
// writes three, which is schema-invalid).

import { XML_DECL, NS, esc, pt100, solidFill } from "./xml.mjs";
import { colName } from "./xlsx.mjs";

export const CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";

const CAT_AX = 50001;
const VAL_AX = 50002;
const NO_LINE = "<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>";

export function chartRefXml(rId) {
  return `<c:chart xmlns:c="${NS.c}" r:id="${rId}"/>`;
}

function txPr(size, color, lang) {
  return (
    `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${pt100(size)}" b="0">${solidFill(color)}` +
    `<a:latin typeface="+mn-lt"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="${esc(lang)}"/></a:p></c:txPr>`
  );
}

export function chartXml(spec, { lang, workbookRId }) {
  const horizontal = spec.type === "bar";
  const n = spec.categories.length;
  const catCache =
    `<c:strCache><c:ptCount val="${n}"/>` +
    `${spec.categories.map((c, i) => `<c:pt idx="${i}"><c:v>${esc(c)}</c:v></c:pt>`).join("")}</c:strCache>`;
  const dLbls = spec.dataLabels
    ? `<c:dLbls>${NO_LINE}${txPr(spec.dataLabels.size, spec.dataLabels.color, lang)}<c:dLblPos val="outEnd"/>` +
      '<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/>' +
      '<c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'
    : "";

  // CT_BarSer order: idx, order, tx, spPr, invertIfNegative, dLbls, cat, val.
  const series = spec.series
    .map((s, j) => {
      const col = colName(j + 1);
      // A series longer than the category list would write c:pt idx values >= ptCount (n) -
      // schema-invalid and fatal to LibreOffice. A shorter series is fine; missing points
      // just mean a gap.
      const pts = s.values
        .slice(0, n)
        .map((v, i) => (Number.isFinite(v) ? `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>` : ""))
        .join("");
      return (
        `<c:ser><c:idx val="${j}"/><c:order val="${j}"/>` +
        `<c:tx><c:strRef><c:f>Sheet1!$${col}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${esc(s.name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
        `<c:spPr>${solidFill(s.color)}<a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/>${dLbls}` +
        `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$${n + 1}</c:f>${catCache}</c:strRef></c:cat>` +
        `<c:val><c:numRef><c:f>Sheet1!$${col}$2:$${col}$${n + 1}</c:f><c:numCache><c:formatCode>${esc(spec.numFmt || "General")}</c:formatCode><c:ptCount val="${n}"/>${pts}</c:numCache></c:numRef></c:val>` +
        "</c:ser>"
      );
    })
    .join("");

  // Number.isFinite (not != null) so Math.max()/Math.min() of an empty array (+/-Infinity) or
  // a NaN never gets written as an invalid xs:double; 0 is finite and still written.
  const scaling =
    '<c:scaling><c:orientation val="minMax"/>' +
    `${Number.isFinite(spec.max) ? `<c:max val="${spec.max}"/>` : ""}${Number.isFinite(spec.min) ? `<c:min val="${spec.min}"/>` : ""}</c:scaling>`;
  const tickMarks = '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>';
  const ticks = `<c:numFmt formatCode="General" sourceLinked="1"/>${tickMarks}`;
  const valTicks = spec.numFmt ? `<c:numFmt formatCode="${esc(spec.numFmt)}" sourceLinked="0"/>${tickMarks}` : ticks;
  // CT_ValAx: majorGridlines sits between axPos and numFmt.
  const grid = spec.gridColor ? `<c:majorGridlines><c:spPr><a:ln w="9525">${solidFill(spec.gridColor, spec.gridAlpha)}</a:ln></c:spPr></c:majorGridlines>` : "";
  const catAx =
    `<c:catAx><c:axId val="${CAT_AX}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>` +
    `<c:axPos val="${horizontal ? "l" : "b"}"/>${ticks}${NO_LINE}${txPr(spec.catLabelSize, spec.catLabelColor, lang)}` +
    `<c:crossAx val="${VAL_AX}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
  // No gridColor = no c:majorGridlines element = no gridlines.
  const valAx =
    `<c:valAx><c:axId val="${VAL_AX}"/>${scaling}<c:delete val="${spec.valueAxis ? 0 : 1}"/>` +
    `<c:axPos val="${horizontal ? "b" : "l"}"/>${grid}${valTicks}${NO_LINE}${txPr(spec.valLabelSize ?? 9, spec.valLabelColor ?? "accent3", lang)}` +
    `<c:crossAx val="${CAT_AX}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  const legend = spec.legend
    ? `<c:legend><c:legendPos val="${spec.legendPos || "b"}"/><c:overlay val="0"/>${txPr(spec.legendSize ?? 11, spec.legendColor ?? "dk1", lang)}</c:legend>`
    : "";
  const background = `<c:spPr>${spec.background ? solidFill(spec.background) : "<a:noFill/>"}<a:ln><a:noFill/></a:ln></c:spPr>`;
  const overlap = spec.stacked ? '<c:overlap val="100"/>' : spec.overlap != null ? `<c:overlap val="${spec.overlap}"/>` : "";

  return (
    `${XML_DECL}<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">` +
    '<c:date1904 val="0"/><c:roundedCorners val="0"/>' +
    '<c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>' +
    `<c:barChart><c:barDir val="${horizontal ? "bar" : "col"}"/><c:grouping val="${spec.stacked ? "stacked" : "clustered"}"/><c:varyColors val="0"/>${series}` +
    `<c:gapWidth val="${spec.gapWidth ?? 100}"/>${overlap}<c:axId val="${CAT_AX}"/><c:axId val="${VAL_AX}"/></c:barChart>` +
    `${catAx}${valAx}${background}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `${background}<c:externalData r:id="${workbookRId}"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`
  );
}
