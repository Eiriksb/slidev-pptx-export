import { test } from "node:test";
import assert from "node:assert/strict";
import { colName, chartWorkbook } from "../src/ooxml/xlsx.mjs";
import { chartXml } from "../src/ooxml/chart.mjs";
import { writePptx } from "../src/ooxml/index.mjs";
import { DECK } from "./deck.mjs";
import { lintPackage } from "./lint.mjs";
import { parseXml, unzip, byLocal } from "./helpers.mjs";

const SPEC = {
  type: "column", categories: ["Q1", "Q2", "Q3"],
  series: [{ name: "Before", color: "accent5", values: [22, 31, 28] }, { name: "After & now", color: "#123456", values: [41, NaN, 63] }],
  min: 0, max: 100, valueAxis: true, legend: true, gapWidth: 45, overlap: -6,
  dataLabels: { color: "dk1", size: 10 }, background: "lt1", catLabelColor: "dk1", catLabelSize: 11, valLabelColor: "accent3", valLabelSize: 9,
};

test("colName maps indexes to spreadsheet columns", () => {
  assert.deepEqual([0, 1, 25, 26, 27].map(colName), ["A", "B", "Z", "AA", "AB"]);
});

test("workbook holds categories, series names and finite values", async () => {
  const buffer = await chartWorkbook(SPEC);
  assert.deepEqual(await lintPackage(buffer), []);
  const z = await unzip(buffer);
  const sheet = z.text("xl/worksheets/sheet1.xml");
  assert.match(sheet, /<dimension ref="A1:C4"\/>/);
  assert.match(sheet, /<c r="B2"><v>22<\/v><\/c>/);
  assert.doesNotMatch(sheet, /r="C3"/); // NaN is left empty
  assert.match(z.text("xl/sharedStrings.xml"), /<t xml:space="preserve">After &amp; now<\/t>/);
  assert.match(z.text("xl/workbook.xml"), /<sheet name="Sheet1" sheetId="1" r:id="rId1"\/>/);
  assert.match(z.text("xl/_rels/workbook.xml.rels"), /Id="rId1" Type="[^"]+\/worksheet" Target="worksheets\/sheet1.xml"/);
});

test("chart XML: two axes, caches, formulas, colors, hidden value axis option", () => {
  const xml = chartXml(SPEC, { lang: "en-US", workbookRId: "rId1" });
  const doc = parseXml(xml);
  assert.equal(byLocal(byLocal(doc, "barChart")[0], "axId").length, 2);
  assert.match(xml, /<c:barDir val="col"\/><c:grouping val="clustered"\/><c:varyColors val="0"\/>/);
  assert.match(xml, /<c:f>Sheet1!\$B\$1<\/c:f><c:strCache><c:ptCount val="1"\/><c:pt idx="0"><c:v>Before<\/c:v>/);
  assert.match(xml, /<c:f>Sheet1!\$A\$2:\$A\$4<\/c:f>/);
  assert.match(xml, /<c:f>Sheet1!\$C\$2:\$C\$4<\/c:f><c:numCache><c:formatCode>General<\/c:formatCode><c:ptCount val="3"\/><c:pt idx="0"><c:v>41<\/c:v><\/c:pt><c:pt idx="2">/);
  assert.match(xml, /<c:spPr><a:solidFill><a:schemeClr val="accent5"\/><\/a:solidFill>/);
  assert.match(xml, /<a:srgbClr val="123456"\/>/);
  assert.match(xml, /<c:scaling><c:orientation val="minMax"\/><c:max val="100"\/><c:min val="0"\/><\/c:scaling><c:delete val="0"\/>/);
  assert.match(xml, /<c:gapWidth val="45"\/><c:overlap val="-6"\/><c:axId val="50001"\/><c:axId val="50002"\/><\/c:barChart>/);
  assert.match(xml, /<c:legend><c:legendPos val="b"\/>/);
  assert.match(xml, /<c:externalData r:id="rId1"><c:autoUpdate val="0"\/><\/c:externalData><\/c:chartSpace>$/);

  const bar = chartXml({ ...SPEC, type: "bar", valueAxis: false, legend: false }, { lang: "en-US", workbookRId: "rId1" });
  assert.match(bar, /<c:barDir val="bar"\/>/);
  assert.match(bar, /<c:valAx><c:axId val="50002"\/><c:scaling>.*<\/c:scaling><c:delete val="1"\/><c:axPos val="b"\/>/);
  assert.doesNotMatch(bar, /<c:legend>/);
});

test("F1: series values beyond the category count are clamped to ptCount, not written past it", async () => {
  const spec = { ...SPEC, categories: ["Q1", "Q2"], series: [{ name: "Before", color: "accent5", values: [22, 31, 28, 99] }] };
  const xml = chartXml(spec, { lang: "en-US", workbookRId: "rId1" });
  parseXml(xml); // well-formed
  // ptCount matches the category count, and no c:pt idx reaches or exceeds it.
  assert.match(xml, /<c:numCache><c:formatCode>General<\/c:formatCode><c:ptCount val="2"\/><c:pt idx="0"><c:v>22<\/c:v><\/c:pt><c:pt idx="1"><c:v>31<\/c:v><\/c:pt><\/c:numCache>/);
  assert.doesNotMatch(xml, /idx="2"><c:v>28/);
  assert.doesNotMatch(xml, /idx="3"/);

  // The embedded workbook is built from the same spec, indexed by category - it never reads
  // (and so never writes) the out-of-range values either, keeping the sheet and the caches in
  // agreement.
  const book = await unzip(await chartWorkbook(spec));
  assert.doesNotMatch(book.text("xl/worksheets/sheet1.xml"), /99/);
  assert.match(book.text("xl/worksheets/sheet1.xml"), /<dimension ref="A1:B3"\/>/);

  const buffer = await writePptx([{ layout: "blank", shapes: [{ kind: "chart", x: 1, y: 1, w: 4, h: 3, chart: spec }] }], DECK);
  assert.deepEqual(await lintPackage(buffer), []);
});

test("F1: non-finite min/max (e.g. Math.max() of an empty array) are omitted, not written as xs:double garbage", () => {
  const spec = { ...SPEC, min: Math.min(), max: Math.max() }; // Infinity / -Infinity
  const xml = chartXml(spec, { lang: "en-US", workbookRId: "rId1" });
  assert.doesNotMatch(xml, /<c:min /);
  assert.doesNotMatch(xml, /<c:max /);
  assert.doesNotMatch(xml, /Infinity/);

  const nanXml = chartXml({ ...SPEC, min: NaN, max: NaN }, { lang: "en-US", workbookRId: "rId1" });
  assert.doesNotMatch(nanXml, /<c:min /);
  assert.doesNotMatch(nanXml, /<c:max /);
  assert.doesNotMatch(nanXml, /NaN/);

  // 0 is finite and must still be written (not treated as falsy/missing).
  const zeroXml = chartXml({ ...SPEC, min: 0, max: 10 }, { lang: "en-US", workbookRId: "rId1" });
  assert.match(zeroXml, /<c:min val="0"\/>/);
});

test("writer adds chart part + embedded workbook and passes lint", async () => {
  const shape = { kind: "chart", x: 1, y: 2, w: 6, h: 3, chart: SPEC };
  const buffer = await writePptx([{ layout: "blank", shapes: [shape, { ...shape, x: 7 }] }], DECK);
  assert.deepEqual(await lintPackage(buffer), []);
  const z = await unzip(buffer);
  assert.ok(z.names.includes("ppt/charts/chart1.xml") && z.names.includes("ppt/charts/chart2.xml"));
  assert.match(z.text("ppt/charts/_rels/chart1.xml.rels"), /\/package" Target="..\/embeddings\/Microsoft_Excel_Worksheet1.xlsx"/);
  assert.match(z.text("[Content_Types].xml"), /<Default Extension="xlsx" ContentType="application\/vnd.openxmlformats-officedocument.spreadsheetml.sheet"\/>/);
  assert.match(z.text("ppt/slides/slide1.xml"), /<a:graphicData uri="http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/chart"><c:chart xmlns:c="[^"]+" r:id="rId2"\/>/);
});

test("render-mode chart options: number format, legend position, stacking, gridlines and a transparent background", async () => {
  const spec = { ...SPEC, numFmt: '0"%"', legendPos: "t", stacked: true, gridColor: "#14181A", gridAlpha: 0.1, background: null };
  const xml = chartXml(spec, { lang: "en-US", workbookRId: "rId1" });
  assert.match(xml, /<c:grouping val="stacked"\/>/);
  assert.match(xml, /<c:gapWidth val="45"\/><c:overlap val="100"\/>/);
  assert.match(xml, /<c:axPos val="l"\/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="14181A"><a:alpha val="10000"\/><\/a:srgbClr><\/a:solidFill><\/a:ln><\/c:spPr><\/c:majorGridlines><c:numFmt formatCode="0&quot;%&quot;" sourceLinked="0"\/>/);
  assert.match(xml, /<c:formatCode>0&quot;%&quot;<\/c:formatCode>/);
  assert.match(xml, /<c:legendPos val="t"\/>/);
  assert.match(xml, /<c:spPr><a:noFill\/><a:ln><a:noFill\/><\/a:ln><\/c:spPr><\/c:plotArea>/);
  const buffer = await writePptx([{ layout: "blank", shapes: [{ kind: "chart", x: 1, y: 1, w: 6, h: 3, chart: spec }] }], DECK);
  assert.deepEqual(await lintPackage(buffer), []);
});
