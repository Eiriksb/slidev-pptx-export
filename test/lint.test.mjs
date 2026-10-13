import { test } from "node:test";
import assert from "node:assert/strict";
import { Package, REL, CT } from "../src/ooxml/package.mjs";
import { lintPackage } from "./lint.mjs";

const P = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

function slide(body) {
  return `<p:sld ${P}><p:cSld><p:spTree>${body}</p:spTree></p:cSld></p:sld>`;
}

async function lintSlide(body, extra = (pkg) => {}) {
  const pkg = new Package();
  pkg.addXml("/ppt/slides/slide1.xml", CT.slide, slide(body));
  extra(pkg);
  return lintPackage(await pkg.toBuffer());
}

test("a clean package has no problems", async () => {
  assert.deepEqual(await lintSlide('<p:sp><p:nvSpPr><p:cNvPr id="2" name="a"/></p:nvSpPr><p:txBody><a:p/></p:txBody></p:sp>'), []);
});

test("flags two pPr in one paragraph (the pptxgenjs bullet bug)", async () => {
  const problems = await lintSlide('<p:sp><p:txBody><a:p><a:pPr/><a:r><a:t>x</a:t></a:r><a:pPr/></a:p></p:txBody></p:sp>');
  assert.ok(problems.some((p) => /more than one a:pPr/.test(p)), problems.join("\n"));
});

test("flags pPr that isn't the first child", async () => {
  const problems = await lintSlide('<p:sp><p:txBody><a:p><a:r><a:t>x</a:t></a:r><a:pPr/></a:p></p:txBody></p:sp>');
  assert.ok(problems.some((p) => /a:pPr is not the first child/.test(p)), problems.join("\n"));
});

test("flags duplicate shape ids and empty text bodies", async () => {
  const problems = await lintSlide('<p:cNvPr id="2" name="a"/><p:cNvPr id="2" name="b"/><p:txBody></p:txBody>');
  assert.ok(problems.some((p) => /duplicate cNvPr id 2/.test(p)), problems.join("\n"));
  assert.ok(problems.some((p) => /txBody without a:p/.test(p)), problems.join("\n"));
});

test("flags an r:id that is not in the part's rels", async () => {
  const problems = await lintSlide('<p:pic><a:blip r:embed="rId9"/></p:pic>');
  assert.ok(problems.some((p) => /r:embed="rId9" not defined/.test(p)), problems.join("\n"));
});

test("flags a bar chart with three axis ids", async () => {
  const problems = await lintSlide("", (pkg) =>
    pkg.addXml("/ppt/charts/chart1.xml", CT.chart, '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:axId val="1"/><c:axId val="2"/><c:axId val="3"/></c:barChart></c:plotArea></c:chart></c:chartSpace>')
  );
  assert.ok(problems.some((p) => /barChart has 3 c:axId/.test(p)), problems.join("\n"));
});

test("flags malformed XML", async () => {
  const problems = await lintSlide("<p:sp>");
  assert.ok(problems.some((p) => /not well-formed/.test(p)), problems.join("\n"));
});

test("flags an Override for a part that is missing (pptxgenjs#1449)", async () => {
  const JSZip = (await import("jszip")).default;
  const zip = new JSZip();
  zip.file("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/slideMasters/slideMaster2.xml" ContentType="x"/></Types>');
  const problems = await lintPackage(await zip.generateAsync({ type: "nodebuffer" }));
  assert.ok(problems.some((p) => /Override for missing part \/ppt\/slideMasters\/slideMaster2.xml/.test(p)), problems.join("\n"));
});
