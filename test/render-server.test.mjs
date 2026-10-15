import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, mkdirSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { slidevCli, startDeckServer } from "../src/render/server.mjs";
import { loadChromium } from "../src/render/browser.mjs";

// A throwaway deck whose node_modules/@slidev/cli is a tiny fake CLI script (null = none).
function fakeDeck(cliSource, { hoisted = false } = {}) {
  const root = mkdtempSync(join(tmpdir(), "pptx-render-"));
  const deckDir = hoisted ? join(root, "decks", "talk") : root;
  mkdirSync(deckDir, { recursive: true });
  writeFileSync(join(deckDir, "slides.md"), "# Hi\n");
  if (cliSource != null) {
    const pkgDir = join(root, "node_modules", "@slidev", "cli");
    mkdirSync(join(pkgDir, "bin"), { recursive: true });
    writeFileSync(join(pkgDir, "package.json"), JSON.stringify({ name: "@slidev/cli", version: "52.0.0-test", bin: { slidev: "./bin/slidev.mjs" } }));
    writeFileSync(join(pkgDir, "bin", "slidev.mjs"), cliSource);
  }
  return { root, deckDir, inputPath: join(deckDir, "slides.md") };
}

const SERVES = [
  'import { createServer } from "node:http";',
  'const port = Number(process.argv[process.argv.indexOf("--port") + 1]);',
  'createServer((req, res) => res.end("ok")).listen(port, "127.0.0.1");',
].join("\n");

test("finds @slidev/cli next to the deck or hoisted into a parent folder", () => {
  const local = fakeDeck(SERVES);
  assert.equal(slidevCli(local.deckDir).bin, join(local.root, "node_modules/@slidev/cli/bin/slidev.mjs"));
  const hoisted = fakeDeck(SERVES, { hoisted: true });
  assert.deepEqual(slidevCli(hoisted.deckDir), { bin: join(hoisted.root, "node_modules/@slidev/cli/bin/slidev.mjs"), version: "52.0.0-test" });
  assert.equal(slidevCli(fakeDeck(null).deckDir), null);
});

test("without @slidev/cli the export stops and says to run npm install in the deck folder", async () => {
  const { inputPath, deckDir } = fakeDeck(null);
  await assert.rejects(startDeckServer({ inputPath }), (err) => err.message.includes("@slidev/cli") && err.message.includes(`run npm install in ${deckDir}`));
});

test("a Slidev that exits during startup reports its own error output", async () => {
  const { inputPath } = fakeDeck('console.error(\'The theme "@slidev/theme-default" was not found\'); process.exit(1);');
  await assert.rejects(startDeckServer({ inputPath }), /Slidev exited before it was ready \(code 1\):[\s\S]*The theme "@slidev\/theme-default" was not found/);
});

test("a Slidev that never answers times out, showing its output", async () => {
  const { inputPath } = fakeDeck('console.log("still starting"); setInterval(() => {}, 1000);');
  await assert.rejects(startDeckServer({ inputPath, timeout: 1500 }), /Slidev didn't start within 1.5s:[\s\S]*still starting/);
});

test("starts the deck's Slidev on a free local port and stops it", async () => {
  const { inputPath } = fakeDeck(SERVES);
  const server = await startDeckServer({ inputPath });
  assert.match(server.url, /^http:\/\/127\.0\.0\.1:\d+$/);
  assert.equal(await (await fetch(server.url)).text(), "ok");
  assert.equal(server.version, "52.0.0-test");
  await server.stop();
  await assert.rejects(fetch(server.url));
});

test("loadChromium explains how to install Playwright when it's missing", async () => {
  await assert.rejects(loadChromium("/nowhere", { resolve: () => null }), /npm i -D playwright-chromium[\s\S]*npx playwright install chromium/);
});

test("loadChromium takes playwright-chromium from the deck folder first", async () => {
  const dir = mkdtempSync(join(tmpdir(), "pptx-pw-"));
  const file = join(dir, "fake-playwright.mjs");
  writeFileSync(file, "export const chromium = { name: 'fake' };");
  const seen = [];
  const chromium = await loadChromium("/deck", {
    resolve: (name, base) => {
      seen.push(`${name}@${base}`);
      return name === "playwright-chromium" && base === "/deck" ? file : null;
    },
  });
  assert.equal(chromium.name, "fake");
  assert.deepEqual(seen, ["playwright-chromium@/deck"]);
});
