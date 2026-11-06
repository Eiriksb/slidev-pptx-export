// Render mode drives Chromium through Playwright. Playwright is an optional peer dependency -
// installing slidev-pptx-export never downloads a browser - so it's resolved at runtime: the
// deck's own node_modules first (where `slidev export` users already have playwright-chromium),
// then this package's. `playwright-core` ships no browser, so it is only used when an executable
// is named (`render.browser` / SLIDEV_PPTX_BROWSER).

import { createRequire } from "node:module";
import { join } from "node:path";
import { pathToFileURL, fileURLToPath } from "node:url";

const PACKAGE_ROOT = fileURLToPath(new URL("../../", import.meta.url));

function defaultResolve(name, base) {
  try {
    return createRequire(join(base, "package.json")).resolve(name);
  } catch {
    return null;
  }
}

export async function loadChromium(deckDir, { resolve = defaultResolve, executablePath, importer = (p) => import(pathToFileURL(p).href) } = {}) {
  for (const name of ["playwright-chromium", "playwright", "playwright-core"]) {
    for (const base of [deckDir, PACKAGE_ROOT]) {
      const path = resolve(name, base);
      if (!path) continue;
      if (name === "playwright-core" && !executablePath) {
        throw new Error("playwright-core drives a browser it doesn't ship: name one with `render.browser` in pptx.config.mjs or the SLIDEV_PPTX_BROWSER environment variable.");
      }
      const mod = await importer(path);
      const chromium = mod.chromium ?? mod.default?.chromium;
      if (chromium) return chromium;
    }
  }
  throw new Error(
    "render mode needs Playwright's Chromium: run `npm i -D playwright-chromium` (the package `slidev export` uses), then `npx playwright install chromium`.",
  );
}
