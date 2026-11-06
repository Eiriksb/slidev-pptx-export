import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, mkdirSync, writeFileSync, readFileSync, existsSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { startDeckServer } from "../src/render/server.mjs";
import { runStop } from "../src/stop.mjs";

// This file has its own process (node --test runs each file alone), so the stopped state
// runStop leaves behind affects no other test.
test("Slidev started after a stop has begun is killed, not left running", async () => {
  const root = mkdtempSync(join(tmpdir(), "pptx-stop-"));
  try {
    const pkgDir = join(root, "node_modules", "@slidev", "cli");
    mkdirSync(join(pkgDir, "bin"), { recursive: true });
    writeFileSync(join(root, "slides.md"), "# Hi\n");
    writeFileSync(join(pkgDir, "package.json"), JSON.stringify({ name: "@slidev/cli", version: "52.0.0-test", bin: { slidev: "./bin/slidev.mjs" } }));
    // Beats into a file every 50 ms for as long as it lives.
    const beats = join(root, "beats");
    writeFileSync(join(pkgDir, "bin", "slidev.mjs"), `import { appendFileSync } from "node:fs"; setInterval(() => appendFileSync(${JSON.stringify(beats)}, "."), 50);`);
    await runStop();
    const started = startDeckServer({ inputPath: join(root, "slides.md"), timeout: 1800 }).catch(() => {});
    await new Promise((r) => setTimeout(r, 1000));
    const size = () => (existsSync(beats) ? readFileSync(beats, "utf8").length : 0);
    const before = size();
    await new Promise((r) => setTimeout(r, 400));
    assert.equal(size(), before, "the Slidev child stopped beating: it was killed");
    await started;
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});
