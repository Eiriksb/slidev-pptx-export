import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync, writeFileSync, mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { convert } from "../src/index.mjs";
import fixtureConfig from "./fixtures/kitchen-sink/pptx.config.mjs";
import { lintPackage } from "./lint.mjs";
import { unzip, byLocal } from "./helpers.mjs";

const DIR = fileURLToPath(new URL("./fixtures/kitchen-sink/", import.meta.url));
const OUT = mkdtempSync(join(tmpdir(), "slidev-pptx-"));

async function exportFixture(name = "deck.pptx") {
  const result = await convert({ inputPath: join(DIR, "slides.md"), outputPath: join(OUT, name), config: fixtureConfig, configDir: DIR });
  const buffer = readFileSync(result.outputPath);
  return { result, buffer, z: await unzip(buffer) };
}

function layoutName(z, n) {
  const rels = z.text(`ppt/slides/_rels/slide${n}.xml.rels`);
  const target = /\/slideLayout" Target="..\/slideLayouts\/(slideLayout\d+\.xml)"/.exec(rels)[1];
  return byLocal(z.xml(`ppt/slideLayouts/${target}`), "cSld")[0].getAttribute("name");
}

test("fixture deck exports, lints clean and reports charts", async () => {
  const { result, buffer } = await exportFixture();
  assert.deepEqual(await lintPackage(buffer), []);
  assert.equal(result.slideCount, 9);
  assert.equal(result.translatedCharts, 3);
  assert.deepEqual(result.unresolvedCharts, ["UnknownWidget"]);
});

test("slides map to the expected layouts", async () => {
  const { z } = await exportFixture();
  assert.deepEqual([1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => layoutName(z, n)), [
    "Title Slide", "Title and Content", "Title Only", "Centered", "Section Header", "Blank", "Title Only", "Title Only", "Title Only",
  ]);
});

test("titles and text-only bodies live in placeholders", async () => {
  const { z } = await exportFixture();
  assert.match(z.text("ppt/slides/slide1.xml"), /<p:ph type="ctrTitle"\/>.*<a:t>Kitchen Sink Deck<\/a:t>/);
  const s2 = z.text("ppt/slides/slide2.xml");
  assert.match(s2, /<p:ph type="title"\/>.*<a:t>Plain content slide<\/a:t>/);
  assert.match(s2, /<p:ph idx="1"\/>.*<a:t>First bullet<\/a:t>/);
  assert.match(s2, /<a:buAutoNum type="arabicPeriod"\/>/);
});

test("speaker notes are exported", async () => {
  const { z } = await exportFixture();
  assert.match(z.text("ppt/notesSlides/notesSlide1.xml"), /<a:t>Cover notes line one.<\/a:t>.*<a:t>Cover notes after a blank line.<\/a:t>/);
  assert.match(z.text("ppt/notesSlides/notesSlide2.xml"), /<a:t>Notes for the text slide<\/a:t>/);
});

test("charts carry an embedded workbook with the data", async () => {
  const { z } = await exportFixture();
  assert.deepEqual(z.names.filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n)).length, 3);
  const book = await unzip(z.buffer("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"));
  assert.match(book.text("xl/sharedStrings.xml"), /Before/);
});

test("theme, fonts, logo and document title come from the deck", async () => {
  const { z } = await exportFixture();
  const theme = z.text("ppt/theme/theme1.xml");
  assert.match(theme, /<a:accent1><a:srgbClr val="1A833B"\/><\/a:accent1>/);
  assert.match(theme, /<a:majorFont><a:latin typeface="Plus Jakarta Sans"\/>/);
  assert.match(z.text("ppt/slideMasters/slideMaster1.xml"), /name="Logo"/);
  assert.match(z.text("docProps/core.xml"), /<dc:title>Kitchen Sink<\/dc:title><dc:creator>Fixture<\/dc:creator>/);
});

const EMU = 914400;
// Top and bottom (inches) of a named shape's xfrm in a part.
function shapeSpan(doc, shapeName) {
  const nv = byLocal(doc, "cNvPr").find((e) => e.getAttribute("name") === shapeName);
  const xfrm = byLocal(nv.parentNode.parentNode, "xfrm")[0];
  const y = Number(byLocal(xfrm, "off")[0].getAttribute("y")) / EMU;
  return { top: y, bottom: y + Number(byLocal(xfrm, "ext")[0].getAttribute("cy")) / EMU };
}
function contentLayout(z) {
  const name = z.names.find((n) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(n)
    && byLocal(z.xml(n), "cSld")[0].getAttribute("name") === "Title and Content");
  return z.xml(name);
}

test("the content layout's body placeholder ends 0.1in above the footer logo", async () => {
  const { z } = await exportFixture("logo.pptx");
  const logo = shapeSpan(z.xml("ppt/slideMasters/slideMaster1.xml"), "Logo");
  const body = shapeSpan(contentLayout(z), "Content Placeholder 2");
  assert.ok(Math.abs(body.bottom - (logo.top - 0.1)) < 0.001, `body ends ${body.bottom}, logo starts ${logo.top}`);
});

test("a logo that can't be embedded reserves no footer space", async () => {
  const result = await convert({
    inputPath: join(DIR, "slides.md"), outputPath: join(OUT, "nologo.pptx"),
    config: { ...fixtureConfig, logo: { path: "./public/missing.png", aspect: 2 } }, configDir: DIR,
  });
  const z = await unzip(readFileSync(result.outputPath));
  assert.doesNotMatch(z.text("ppt/slideMasters/slideMaster1.xml"), /name="Logo"/);
  const body = shapeSpan(contentLayout(z), "Content Placeholder 2");
  assert.ok(Math.abs(body.bottom - (7.5 - 0.35)) < 0.001, `body ends ${body.bottom}`);
});

test("a long table continues on extra slides; slideCount counts exported slides", async () => {
  const rows = Array.from({ length: 40 }, (_, i) => `| row ${i} | ${i} |`).join("\n");
  const md = join(OUT, "long-table.md");
  writeFileSync(md, `# Cover\n\n---\n\n# Big table\n\n| Name | Value |\n| --- | --- |\n${rows}\n\n<!-- notes -->\n`);
  const result = await convert({ inputPath: md, outputPath: join(OUT, "long-table.pptx") });
  const buffer = readFileSync(result.outputPath);
  assert.deepEqual(await lintPackage(buffer), []);
  assert.equal(result.slideCount, 4);
  const z = await unzip(buffer);
  assert.ok(z.names.includes("ppt/slides/slide4.xml"));
  assert.match(z.text("ppt/slides/slide4.xml"), /Big table \(cont\.\)/);
  assert.equal(z.names.filter((n) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(n)).length, 1);
});

test("two exports of the same deck are byte-identical", async () => {
  const a = await exportFixture("a.pptx");
  const b = await exportFixture("b.pptx");
  assert.ok(a.buffer.equals(b.buffer));
});
