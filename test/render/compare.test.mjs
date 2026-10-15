import { test, after } from "node:test";
import assert from "node:assert/strict";
import { usableChromium } from "./helpers.mjs";
import { similarity } from "../../src/render/compare.mjs";

const chromium = await usableChromium();
const browser = chromium ? await chromium.launch() : null;
after(() => browser?.close());
const opts = { skip: browser ? false : "Playwright/Chromium not available" };

test("identical images score 1; black against white scores 0", opts, async () => {
  const page = await browser.newPage({ viewport: { width: 320, height: 180 } });
  const shot = async (color) => {
    await page.setContent(`<body style="margin:0;background:${color}"></body>`);
    return page.screenshot();
  };
  const white = await shot("#ffffff");
  const black = await shot("#000000");
  assert.equal(await similarity(page, white, white), 1);
  assert.ok(Math.abs(await similarity(page, white, black)) < 0.001);
  await page.close();
});
