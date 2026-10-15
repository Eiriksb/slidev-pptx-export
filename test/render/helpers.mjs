// Shared setup for the browser tests (npm run test:render). Tests skip - never fail - when
// Playwright or its Chromium isn't installed, so this suite is safe to run anywhere.
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { loadChromium } from "../../src/render/browser.mjs";

export const ROOT = fileURLToPath(new URL("../../", import.meta.url));
export const EXTRACT = readFileSync(new URL("../../src/render/extract.js", import.meta.url), "utf8");

// Playwright's chromium if it loads and can launch here, else null.
export async function usableChromium() {
  try {
    const chromium = await loadChromium(ROOT);
    const browser = await chromium.launch();
    await browser.close();
    return chromium;
  } catch {
    return null;
  }
}

// One slide the way Slidev's /<n>?print draws it: #slide-content (the 980x552 canvas, styled by
// `contentStyle`) holding the global-bottom layer, #slideshow with the `.slidev-page` (styled by
// `pageStyle`; its layout root by `layoutClass` / `layoutStyle`), then the global-top layer.
export function slideHtml(html, { layoutClass = "slidev-layout default", layoutStyle = "", pageStyle = "", contentStyle = "", globalBottom = "", globalTop = "", css = "", head = "" } = {}) {
  return (
    `<!doctype html><html><head><style>body{margin:0;font-family:Arial,sans-serif}${css}</style>${head}</head><body>` +
    `<div id="slide-content" class="slidev-slide-content" style="position:relative;width:980px;height:552px;overflow:hidden;${contentStyle}">${globalBottom}` +
    '<div id="slideshow" style="height:100%">' +
    `<div class="slidev-page slidev-page-1" style="position:absolute;inset:0;overflow:hidden;${pageStyle}">` +
    `<div class="${layoutClass}" style="width:100%;height:100%;box-sizing:border-box;padding:40px;${layoutStyle}">${html}</div></div></div>${globalTop}</div></body></html>`
  );
}

// Renders `html` as one Slidev slide (see slideHtml for the options) and returns its Scene.
// `setup` (optional) runs in the page before extraction.
export async function extractHtml(browser, html, { setup = null, ...options } = {}) {
  const page = await browser.newPage({ viewport: { width: 980, height: 552 } });
  try {
    await page.setContent(slideHtml(html, options));
    if (setup) await page.evaluate(setup);
    await page.addScriptTag({ content: EXTRACT });
    return await page.evaluate(() => window.__pptxExtract(document.querySelector(".slidev-page"), { w: 980, h: 552 }));
  } finally {
    await page.close();
  }
}
