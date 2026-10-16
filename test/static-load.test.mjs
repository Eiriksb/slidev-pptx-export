// The static export never loads render mode's browser pipeline (capture.mjs reads extract.js as
// soon as it is imported). Only dynamic imports here, so the hook sees every module load.
import { test, after } from "node:test";
import assert from "node:assert/strict";
import module from "node:module";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const OUT = mkdtempSync(join(tmpdir(), "slidev-pptx-static-"));
after(() => rmSync(OUT, { recursive: true, force: true }));

test("a static export loads nothing from src/render", { skip: typeof module.registerHooks !== "function" && "needs module.registerHooks (Node 22.15+)" }, async () => {
  const loaded = [];
  module.registerHooks({
    load(url, context, next) {
      if (url.includes("/src/render/")) loaded.push(url.slice(url.indexOf("/src/") + 1));
      return next(url, context);
    },
  });
  const { convert } = await import("../src/index.mjs");
  const deck = fileURLToPath(new URL("./fixtures/kitchen-sink/slides.md", import.meta.url));
  const result = await convert({ inputPath: deck, outputPath: join(OUT, "deck.pptx") });
  assert.ok(result.slideCount > 0);
  assert.deepEqual(loaded, []);
});
