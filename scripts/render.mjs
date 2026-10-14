#!/usr/bin/env node
// LibreOffice render check. Exports decks, converts them to PDF and per-slide PNGs for visual
// review, and checks page count and slide titles with poppler. Skips (exit 0) when soffice
// isn't installed. Usage: node scripts/render.mjs [path/to/slides.md ...]
// Default decks: the kitchen-sink fixture, plus ../innotek-SlideDevBase if it's checked out.

import { execFileSync, spawnSync } from "node:child_process";
import { existsSync, mkdirSync, rmSync, readFileSync } from "node:fs";
import { resolve, dirname, basename, join } from "node:path";
import { pathToFileURL, fileURLToPath } from "node:url";
import { convert, mergeConfig } from "../src/index.mjs";
import { splitSlides, parseSlide } from "../src/parse.mjs";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const DEFAULT_DECKS = [join(ROOT, "test/fixtures/kitchen-sink/slides.md"), resolve(ROOT, "../innotek-SlideDevBase/slides.md")];

const has = (cmd) => spawnSync("which", [cmd]).status === 0;
const squash = (s) => s.replace(/\s+/g, " ");

async function loadConfig(deckPath) {
  const path = join(dirname(deckPath), "pptx.config.mjs");
  if (!existsSync(path)) return { config: {}, configDir: dirname(deckPath) };
  return { config: (await import(pathToFileURL(path).href)).default || {}, configDir: dirname(path) };
}

function slideTitles(deckPath, config) {
  const merged = mergeConfig(config);
  const { slides } = splitSlides(readFileSync(deckPath, "utf8"));
  return slides.map((s) => parseSlide(s.content, [], merged).title).filter(Boolean);
}

async function main() {
  if (!has("soffice")) {
    console.log("soffice (LibreOffice) not found - skipping render check.");
    return;
  }
  const args = process.argv.slice(2).map((p) => resolve(p));
  const decks = (args.length ? args : DEFAULT_DECKS).filter((p) => existsSync(p));
  let failed = false;
  for (const deck of decks) {
    const name = basename(dirname(deck));
    const outDir = join(ROOT, "out/render", name);
    rmSync(outDir, { recursive: true, force: true });
    mkdirSync(outDir, { recursive: true });

    const { config, configDir } = await loadConfig(deck);
    const result = await convert({ inputPath: deck, outputPath: join(outDir, "deck.pptx"), config, configDir });
    execFileSync("soffice", ["--headless", "--convert-to", "pdf", "--outdir", outDir, result.outputPath], { stdio: "ignore", timeout: 180000 });
    const pdf = join(outDir, "deck.pdf");
    const pages = Number(/Pages:\s+(\d+)/.exec(execFileSync("pdfinfo", [pdf], { encoding: "utf8" }))[1]);
    const text = squash(execFileSync("pdftotext", [pdf, "-"], { encoding: "utf8" }));
    const missing = slideTitles(deck, config).filter((t) => !text.includes(squash(t)));
    execFileSync("pdftoppm", ["-r", "60", "-png", pdf, join(outDir, "slide")]);

    const ok = pages === result.slideCount && missing.length === 0;
    failed ||= !ok;
    const detail = missing.length ? `, titles missing from the render: ${missing.join(" | ")}` : "";
    console.log(`${ok ? "ok  " : "FAIL"} ${name}: ${pages}/${result.slideCount} pages${detail} -> ${outDir}`);
  }
  if (failed) process.exitCode = 1;
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
