import { test } from "node:test";
import assert from "node:assert/strict";
import { existsSync, readFileSync, mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, dirname } from "node:path";
import { pathToFileURL } from "node:url";
import { ROOT, usableChromium } from "./helpers.mjs";
import { startDeckServer } from "../../src/render/server.mjs";
import { captureDeck } from "../../src/render/capture.mjs";
import { convert } from "../../src/index.mjs";
import { lintPackage } from "../lint.mjs";
import { unzip } from "../helpers.mjs";

export const INNOTEK = join(ROOT, "../innotek-SlideDevBase/slides.md");
const chromium = existsSync(INNOTEK) ? await usableChromium() : null;
export const deckOpts = {
  skip: !existsSync(INNOTEK) ? "../innotek-SlideDevBase is not checked out" : !chromium ? "Playwright/Chromium not available" : false,
  timeout: 300000,
};
const PNG_MAGIC = "89504e47";

test("captures every slide of the innotek deck: scenes, notes, image bytes and screenshots", deckOpts, async () => {
  const server = await startDeckServer({ inputPath: INNOTEK });
  let captured;
  try {
    captured = await captureDeck({ baseUrl: server.url, chromium, screenshots: true, slidevVersion: server.version });
  } finally {
    await server.stop();
  }
  assert.deepEqual(captured.canvas, { w: 980, h: 552 });
  assert.equal(captured.slides.length, 7);
  assert.deepEqual(captured.slides.filter((s) => s.error), []);
  assert.deepEqual(captured.slides.map((s) => s.sourceIndex), [0, 1, 2, 3, 4, 5, 6]);
  assert.equal(captured.slides[0].layout, "cover");
  const cardRow = captured.slides[2];
  assert.equal(cardRow.nodes.find((n) => n.role === "title").runs[0].text, "Card row");
  assert.equal(cardRow.nodes.filter((n) => n.kind === "box" && n.radius === 8).length, 3);
  for (const s of captured.slides) {
    assert.equal(s.screenshot.subarray(0, 4).toString("hex"), PNG_MAGIC);
    for (const n of s.nodes.filter((x) => x.kind === "raster")) assert.equal(n.png.subarray(0, 4).toString("hex"), PNG_MAGIC, n.label);
    for (const n of s.nodes.filter((x) => x.kind === "image")) assert.ok(n.data.length > 0 && (!n.natural || n.natural.w > 0), n.src);
  }
});

export async function exportInnotek(name = "innotek-render.pptx", overrides = {}) {
  const config = { ...(await import(pathToFileURL(join(dirname(INNOTEK), "pptx.config.mjs")).href)).default, ...overrides };
  const outputPath = join(mkdtempSync(join(tmpdir(), "pptx-render-")), name);
  const result = await convert({ inputPath: INNOTEK, outputPath, config, configDir: dirname(INNOTEK), render: true });
  const buffer = readFileSync(outputPath);
  return { result, buffer, z: await unzip(buffer) };
}

test("exports the innotek deck with --render: lint-clean, title placeholders, native cards and eyebrow", deckOpts, async () => {
  const { result, buffer, z } = await exportInnotek();
  assert.equal(result.slideCount, 7);
  assert.deepEqual(await lintPackage(buffer), []);
  const cards = z.text("ppt/slides/slide3.xml");
  assert.match(cards, /<p:ph type="title"\/><\/p:nvPr><\/p:nvSpPr><p:spPr><a:xfrm>[\s\S]*?<a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:noAutofit\/><\/a:bodyPr>[\s\S]*?Card row/);
  assert.equal((cards.match(/prst="roundRect"/g) || []).length, 3);
  assert.match(cards, /<a:solidFill><a:schemeClr val="bg2"\/><\/a:solidFill><a:ln w="12446"><a:solidFill><a:schemeClr val="accent4"\/>/);
  assert.match(cards, /<a:solidFill><a:schemeClr val="accent1"\/><\/a:solidFill><a:latin typeface="\+mj-lt"\/><a:ea typeface="\+mj-ea"\/><a:cs typeface="\+mj-cs"\/><\/a:rPr><a:t>OVERVIEW<\/a:t>/);
  assert.match(z.text("ppt/slides/slide1.xml"), /<p:bg><p:bgPr><a:blipFill/);
});
