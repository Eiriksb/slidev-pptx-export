import { test, after } from "node:test";
import assert from "node:assert/strict";
import { usableChromium } from "./helpers.mjs";
import { similarity, inkOverlap } from "../../src/render/compare.mjs";

const chromium = await usableChromium();
const browser = chromium ? await chromium.launch() : null;
after(() => browser?.close());
const opts = { skip: browser ? false : "Playwright/Chromium not available" };

// A 320x180 page of `body` markup (white unless it says otherwise) as a PNG; `omitBackground`
// leaves the unpainted parts transparent.
async function shot(page, body, { omitBackground = false } = {}) {
  await page.setContent(`<body style="margin:0">${body}</body>`);
  return page.screenshot({ omitBackground });
}
const box = (x, y) => `<div style="position:absolute;left:${x}px;top:${y}px;width:40px;height:30px;background:#000"></div>`;
const withPage = (fn) => async () => {
  const page = await browser.newPage({ viewport: { width: 320, height: 180 } });
  try {
    await fn(page);
  } finally {
    await page.close();
  }
};

test("identical images score 1; black against white scores 0", opts, withPage(async (page) => {
  const white = await shot(page, "");
  const black = await shot(page, '<div style="position:absolute;inset:0;background:#000"></div>');
  assert.equal(await similarity(page, white, white), 1);
  assert.ok(Math.abs(await similarity(page, white, black)) < 0.001);
}));

test("a half-black image scores about 0.5 against white", opts, withPage(async (page) => {
  const white = await shot(page, "");
  const half = await shot(page, '<div style="position:absolute;inset:0;background:linear-gradient(to right,#000 50%,#fff 50%)"></div>');
  assert.ok(Math.abs((await similarity(page, white, half)) - 0.5) < 0.02);
}));

test("a transparent PNG scores as if it were on white", opts, withPage(async (page) => {
  const white = await shot(page, "");
  const clear = await shot(page, "", { omitBackground: true });
  assert.notEqual(clear.toString("base64"), white.toString("base64"));
  assert.equal(await similarity(page, white, clear), 1);
  assert.equal(await inkOverlap(page, white, clear), 1);
}));

test("ink overlap: identical images score 1, two blank pages score 1, blank against a page with ink scores 0", opts, withPage(async (page) => {
  const blank = await shot(page, "");
  const marked = await shot(page, box(20, 20));
  assert.equal(await inkOverlap(page, marked, marked), 1);
  assert.equal(await inkOverlap(page, blank, blank), 1);
  assert.ok((await inkOverlap(page, blank, marked)) < 0.001);
  assert.ok((await inkOverlap(page, marked, blank)) < 0.001);
}));

test("ink overlap sees marks that moved, where the mean difference barely changes", opts, withPage(async (page) => {
  const here = await shot(page, box(20, 20));
  const there = await shot(page, box(200, 120));
  assert.ok((await similarity(page, here, there)) > 0.9);
  assert.ok((await inkOverlap(page, here, there)) < 0.001);
  // Marks half over each other share 20px of a 60px-wide union: a third of their ink.
  const nudged = await shot(page, box(40, 20));
  const overlap = await inkOverlap(page, here, nudged);
  assert.ok(overlap > 0.25 && overlap < 0.45, `overlap ${overlap}`);
}));
