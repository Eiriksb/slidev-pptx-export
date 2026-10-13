// Minimal embedded workbook for one chart, so PowerPoint's "Edit Data" opens the numbers in
// Excel. Column A holds categories, row 1 series names, B2.. the values - matching the c:f
// formulas chart.mjs writes. Built with the same Package class, so it's lint-checked too.

import { Package, REL, CT } from "./package.mjs";
import { XML_DECL, NS, esc } from "./xml.mjs";

export function colName(i) {
  let n = i + 1;
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

const STYLES =
  `${XML_DECL}<styleSheet xmlns="${NS.ssml}">` +
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

export async function chartWorkbook({ categories, series }) {
  const strings = [];
  const index = new Map();
  const str = (text) => {
    const key = String(text);
    if (!index.has(key)) {
      index.set(key, strings.length);
      strings.push(key);
    }
    return index.get(key);
  };

  const rows = [`<row r="1">${series.map((s, j) => `<c r="${colName(j + 1)}1" t="s"><v>${str(s.name)}</v></c>`).join("")}</row>`];
  categories.forEach((cat, i) => {
    const r = i + 2;
    const cells = [`<c r="A${r}" t="s"><v>${str(cat)}</v></c>`];
    series.forEach((s, j) => {
      const v = s.values[i];
      if (Number.isFinite(v)) cells.push(`<c r="${colName(j + 1)}${r}"><v>${v}</v></c>`);
    });
    rows.push(`<row r="${r}">${cells.join("")}</row>`);
  });

  const sheet =
    `${XML_DECL}<worksheet xmlns="${NS.ssml}" xmlns:r="${NS.r}"><dimension ref="A1:${colName(series.length)}${categories.length + 1}"/>` +
    `<sheetData>${rows.join("")}</sheetData></worksheet>`;
  const sst =
    `${XML_DECL}<sst xmlns="${NS.ssml}" count="${strings.length}" uniqueCount="${strings.length}">` +
    `${strings.map((t) => `<si><t xml:space="preserve">${esc(t)}</t></si>`).join("")}</sst>`;
  const workbook =
    `${XML_DECL}<workbook xmlns="${NS.ssml}" xmlns:r="${NS.r}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`;

  const pkg = new Package();
  const WB = "/xl/workbook.xml";
  pkg.rel("", WB, REL.officeDocument);
  pkg.rel(WB, "/xl/worksheets/sheet1.xml", REL.worksheet); // rId1, referenced by <sheet r:id>
  pkg.rel(WB, "/xl/styles.xml", REL.styles);
  pkg.rel(WB, "/xl/sharedStrings.xml", REL.sharedStrings);
  pkg.addXml(WB, CT.workbook, workbook);
  pkg.addXml("/xl/worksheets/sheet1.xml", CT.worksheet, sheet);
  pkg.addXml("/xl/styles.xml", CT.styles, STYLES);
  pkg.addXml("/xl/sharedStrings.xml", CT.sharedStrings, sst);
  return pkg.toBuffer();
}
