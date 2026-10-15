// Render mode drives Chromium through Playwright. Playwright is an optional peer dependency -
// installing slidev-pptx-export never downloads a browser - so it's resolved at runtime: the
// deck's own node_modules first (where `slidev export` users already have playwright-chromium),
// then this package's.

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

export async function loadChromium(deckDir, { resolve = defaultResolve } = {}) {
  for (const name of ["playwright-chromium", "playwright"]) {
    for (const base of [deckDir, PACKAGE_ROOT]) {
      const path = resolve(name, base);
      if (!path) continue;
      const mod = await import(pathToFileURL(path).href);
      const chromium = mod.chromium ?? mod.default?.chromium;
      if (chromium) return chromium;
    }
  }
  throw new Error(
    "render mode needs Playwright's Chromium: run `npm i -D playwright-chromium` (the package `slidev export` uses), then `npx playwright install chromium`.",
  );
}
