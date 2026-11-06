import { test } from "node:test";
import assert from "node:assert/strict";
import { loadChromium } from "../src/render/browser.mjs";

const fake = (present) => (name) => (present.includes(name) ? `/fake/${name}/index.js` : null);
// The "module" at each fake path is a chromium tagged with the package it came from.
const importer = async (p) => ({ chromium: { from: p.split("/")[2] } });

test("playwright-chromium still comes first", async () => {
  const chromium = await loadChromium("/d", { resolve: fake(["playwright-chromium", "playwright", "playwright-core"]), importer });
  assert.equal(chromium.from, "playwright-chromium");
});

test("playwright-core alone is used when a browser is named", async () => {
  const chromium = await loadChromium("/d", { resolve: fake(["playwright-core"]), executablePath: "/opt/shell", importer });
  assert.equal(chromium.from, "playwright-core");
});

test("playwright-core is passed over for playwright when no browser is named", async () => {
  const chromium = await loadChromium("/d", { resolve: fake(["playwright", "playwright-core"]), importer });
  assert.equal(chromium.from, "playwright");
});

test("playwright-core alone with no browser named is an error that names render.browser", async () => {
  await assert.rejects(loadChromium("/d", { resolve: fake(["playwright-core"]), importer }), /render\.browser|SLIDEV_PPTX_BROWSER/);
});
