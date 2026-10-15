import { test } from "node:test";
import assert from "node:assert/strict";
import { existsSync } from "node:fs";
import { join } from "node:path";
import { ROOT, usableChromium } from "./helpers.mjs";
import { startDeckServer } from "../../src/render/server.mjs";
import { captureDeck } from "../../src/render/capture.mjs";

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
