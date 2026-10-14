import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync, mkdtempSync } from "node:fs";
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

test("two exports of the same deck are byte-identical", async () => {
  const a = await exportFixture("a.pptx");
  const b = await exportFixture("b.pptx");
  assert.ok(a.buffer.equals(b.buffer));
});
