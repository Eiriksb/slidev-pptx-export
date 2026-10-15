import { test } from "node:test";
import assert from "node:assert/strict";
import { tableStylesXml, tableXml, TABLE_STYLES } from "../src/ooxml/table.mjs";
import { writePptx } from "../src/ooxml/index.mjs";
import { DECK } from "./deck.mjs";
import { lintPackage } from "./lint.mjs";
import { parseXml, unzip, byLocal } from "./helpers.mjs";

const ctx = { lang: "en-US", fonts: { mono: "M" } };
const cell = (text) => [{ runs: [{ text, size: 13 }] }];
const TABLE = { kind: "table", x: 1, y: 2, w: 6, h: 1.2, header: true, rows: [[cell("A"), cell("B"), cell("C")], [cell("1"), cell("2")], [cell("x"), cell("y"), cell("z")]] };

test("table styles part defines both deck styles with theme colors only", () => {
  const xml = tableStylesXml();
  const doc = parseXml(xml);
  assert.equal(doc.documentElement.getAttribute("def"), TABLE_STYLES.default);
  const ids = byLocal(doc, "tblStyle").map((s) => s.getAttribute("styleId"));
  assert.deepEqual(ids, [TABLE_STYLES.default, TABLE_STYLES.deflist]);
  assert.doesNotMatch(xml, /srgbClr/);
  assert.match(xml, /<a:firstRow><a:tcTxStyle b="on">/);
});

test("table: style id + header flag, padded rows, grid widths sum to frame width", () => {
  const xml = tableXml(TABLE, ctx);
  assert.match(xml, new RegExp(`<a:tblPr firstRow="1" bandRow="0"><a:tableStyleId>${TABLE_STYLES.default.replace(/[{}]/g, "\\$&")}</a:tableStyleId></a:tblPr>`));
  const doc = parseXml(`<x xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${xml}</x>`);
  const rows = byLocal(doc, "tr");
  assert.deepEqual(rows.map((r) => byLocal(r, "tc").length), [3, 3, 3]);
  const widths = byLocal(doc, "gridCol").map((c) => Number(c.getAttribute("w")));
  assert.ok(Math.abs(widths.reduce((a, b) => a + b, 0) - 6 * 914400) <= 3);
  assert.match(xml, /<a:tcPr marL="91440" marR="91440" marT="45720" marB="45720"\/>/);
  assert.doesNotMatch(xml, /<a:tcPr[^>]*>.*<a:solidFill/);
});

test("rows share the frame height evenly unless the shape gives per-row heights", () => {
  const heights = (xml) => byLocal(parseXml(`<x xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${xml}</x>`), "tr").map((r) => Number(r.getAttribute("h")));
  assert.deepEqual(heights(tableXml(TABLE, ctx)), [365760, 365760, 365760]);
  assert.deepEqual(heights(tableXml({ ...TABLE, rowH: [0.3, 0.5, 0.4] }, ctx)), [274320, 457200, 365760]);
});

test("deflist variant: own style, no header, explicit column widths, centered cells", () => {
  const xml = tableXml({ ...TABLE, header: false, variant: "deflist", colW: [2.4, 3.6], rows: [[cell("t"), cell("d")]] }, ctx);
  assert.match(xml, /<a:tblPr bandRow="0"><a:tableStyleId>/);
  assert.ok(xml.includes(TABLE_STYLES.deflist));
  assert.match(xml, /<a:gridCol w="2194560"\/><a:gridCol w="3291840"\/>/);
  assert.match(xml, /anchor="ctr"/);
});

test("writer embeds tables in graphic frames and passes lint", async () => {
  const buffer = await writePptx([{ layout: "blank", shapes: [TABLE] }], DECK);
  assert.deepEqual(await lintPackage(buffer), []);
  const z = await unzip(buffer);
  assert.match(z.text("ppt/slides/slide1.xml"), /<a:graphicData uri="http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/table"><a:tbl>/);
  assert.ok(z.text("ppt/tableStyles.xml").includes(TABLE_STYLES.default));
});

test("plain tables write each cell's fill, borders, margins, anchor and merges, with the No Style table style", async () => {
  const PLAIN = {
    kind: "table", x: 0, y: 0, w: 4, h: 1, variant: "plain", header: true, colW: [2, 2], rowH: [0.5, 0.5],
    rows: [
      [{ paras: cell("Head"), fill: "#EEF4F0", borders: { b: { width: 1, color: "accent4" } }, margin: { l: 0.1, r: 0.1, t: 0.05, b: 0.05 }, anchor: "ctr", gridSpan: 2 }, { paras: [], hMerge: true }],
      [{ paras: cell("a") }, { paras: cell("b"), fill: "lt2", fillAlpha: 0.5 }],
    ],
  };
  const xml = tableXml(PLAIN, ctx);
  assert.match(xml, /^<a:tbl><a:tblPr bandRow="0"><a:tableStyleId>\{2D5ABB26-0587-4C30-8999-92F81FD0307C\}<\/a:tableStyleId><\/a:tblPr>/);
  assert.match(xml, /<a:tc gridSpan="2"><a:txBody>[\s\S]*?<\/a:txBody><a:tcPr marL="91440" marR="91440" marT="45720" marB="45720" anchor="ctr"><a:lnB w="12700"><a:solidFill><a:schemeClr val="accent4"\/><\/a:solidFill><\/a:lnB><a:solidFill><a:srgbClr val="EEF4F0"\/><\/a:solidFill><\/a:tcPr><\/a:tc><a:tc hMerge="1">/);
  assert.match(xml, /<a:solidFill><a:schemeClr val="bg2"><a:alpha val="50000"\/><\/a:schemeClr><\/a:solidFill><\/a:tcPr><\/a:tc><\/a:tr><\/a:tbl>$/);
  const buffer = await writePptx([{ layout: "blank", shapes: [PLAIN] }], DECK);
  assert.deepEqual(await lintPackage(buffer), []);
});

test("plain tables: row spans, bare object and array cells, alpha and sub-pixel borders", async () => {
  const EDGE = {
    kind: "table", x: 0, y: 0, w: 4, h: 1.5, variant: "plain", colW: [2, 2], rowH: [0.5, 0.5, 0.5],
    rows: [
      [{ paras: cell("span"), rowSpan: 2 }, { paras: cell("sub"), borders: { r: { width: 0.00001, color: "accent4" } } }],
      [{ paras: [], vMerge: true }, { paras: cell("alpha"), borders: { t: { width: 1, color: "accent4", alpha: 0.5 } } }],
      [cell("bare"), { paras: cell("obj") }],
    ],
  };
  const xml = tableXml(EDGE, ctx);
  const doc = parseXml(`<x xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${xml}</x>`);
  const [span, sub, covered, alpha, bare, obj] = byLocal(doc, "tc");
  const tcPr = (tc) => byLocal(tc, "tcPr")[0];
  const children = (el) => [...el.childNodes].filter((n) => n.nodeType === 1).map((n) => n.localName);

  assert.equal(span.getAttribute("rowSpan"), "2");
  assert.equal(covered.getAttribute("vMerge"), "1");
  // An object cell with no margin, anchor or borders writes a bare tcPr holding only the fill.
  assert.equal(tcPr(obj).attributes.length, 0);
  assert.deepEqual(children(tcPr(obj)), ["noFill"]);
  assert.equal(tcPr(span).attributes.length, 0);
  assert.deepEqual(children(tcPr(span)), ["noFill"]);
  // A bare paragraph-array cell takes the table's shared margins and nothing else.
  assert.equal(tcPr(bare).getAttribute("marL"), "91440");
  assert.deepEqual(children(tcPr(bare)), []);
  // A border of 0.00001pt rounds to zero and writes no line element.
  assert.equal(byLocal(sub, "lnR").length, 0);
  assert.deepEqual(children(tcPr(sub)), ["noFill"]);
  // Only the sides that have a border get a line element, in lnL, lnR, lnT, lnB order, before the fill.
  assert.deepEqual(children(tcPr(alpha)), ["lnT", "noFill"]);
  // Alpha sits inside the line's color.
  const lnT = byLocal(alpha, "lnT")[0];
  assert.equal(lnT.getAttribute("w"), "12700");
  const a = byLocal(lnT, "alpha")[0];
  assert.equal(a.getAttribute("val"), "50000");
  assert.equal(a.parentNode.localName, "schemeClr");
  assert.equal(a.parentNode.getAttribute("val"), "accent4");
  const buffer = await writePptx([{ layout: "blank", shapes: [EDGE] }], DECK);
  assert.deepEqual(await lintPackage(buffer), []);
});
