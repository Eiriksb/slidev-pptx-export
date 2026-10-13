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
