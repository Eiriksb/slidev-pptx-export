#!/usr/bin/env node
// LibreOffice render check. Exports decks, converts them to PDF and per-slide PNGs for visual
// review, and checks page count and slide titles with poppler. Skips (exit 0) when soffice
// isn't installed. Usage: node scripts/render.mjs [--render] [path/to/slides.md ...]
// Default decks: the kitchen-sink fixture, plus ../innotek-SlideDevBase if it's checked out.
// --render exports in render mode and adds a visual check: each slide's LibreOffice render is
// compared with Slidev's own screenshot of it, and a slide scoring below 0.90 fails.

import { execFileSync, spawnSync } from "node:child_process";
import { existsSync, mkdirSync, rmSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import { resolve, dirname, basename, join } from "node:path";
import { pathToFileURL, fileURLToPath } from "node:url";
import { convert, mergeConfig } from "../src/index.mjs";
import { splitSlides, parseSlide } from "../src/parse.mjs";
import { loadChromium } from "../src/render/browser.mjs";
import { slidevCli } from "../src/render/server.mjs";
import { similarity } from "../src/render/compare.mjs";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const DEFAULT_DECKS = [join(ROOT, "test/fixtures/kitchen-sink/slides.md"), resolve(ROOT, "../innotek-SlideDevBase/slides.md")];

const VISUAL_THRESHOLD = 0.9;

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

// LibreOffice's slide-N.png pages against Slidev's screenshots (null = a slide that fell back
// to the markdown export, not compared).
async function visualScores(chromium, outDir, screenshots) {
  const pages = readdirSync(outDir).filter((f) => /^slide-\d+\.png$/.test(f)).sort((a, b) => parseInt(a.slice(6), 10) - parseInt(b.slice(6), 10));
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage();
    const scores = [];
    for (const [i, file] of pages.entries()) {
      if (!screenshots[i]) continue;
      writeFileSync(join(outDir, `slidev-${i + 1}.png`), screenshots[i]);
      scores.push({ slide: i + 1, score: await similarity(page, screenshots[i], readFileSync(join(outDir, file))) });
    }
    return scores;
  } finally {
    await browser.close();
  }
}

async function main() {
  if (!has("soffice")) {
    console.log("soffice (LibreOffice) not found - skipping render check.");
    return;
  }
  const argv = process.argv.slice(2);
  const render = argv.includes("--render");
  const args = argv.filter((a) => a !== "--render").map((p) => resolve(p));
  let chromium = null;
  if (render) {
    try {
      chromium = await loadChromium(ROOT);
    } catch (err) {
      console.log(`${err.message} - skipping the render-mode check.`);
      return;
    }
  }
  const decks = (args.length ? args : DEFAULT_DECKS).filter((p) => existsSync(p));
  let failed = false;
  for (const deck of decks) {
    if (render && !slidevCli(dirname(deck))) {
      console.log(`skip ${basename(dirname(deck))}: not a Slidev deck with @slidev/cli installed`);
      continue;
    }
    const name = basename(dirname(deck));
    const outDir = join(ROOT, "out/render", name);
    rmSync(outDir, { recursive: true, force: true });
    mkdirSync(outDir, { recursive: true });

    const { config, configDir } = await loadConfig(deck);
    const result = await convert({ inputPath: deck, outputPath: join(outDir, "deck.pptx"), config, configDir, render: render ? { screenshots: true } : undefined });
    execFileSync("soffice", ["--headless", "--convert-to", "pdf", "--outdir", outDir, result.outputPath], { stdio: "ignore", timeout: 180000 });
    const pdf = join(outDir, "deck.pdf");
    const pages = Number(/Pages:\s+(\d+)/.exec(execFileSync("pdfinfo", [pdf], { encoding: "utf8" }))[1]);
    const text = squash(execFileSync("pdftotext", [pdf, "-"], { encoding: "utf8" }));
    const missing = slideTitles(deck, config).filter((t) => !text.includes(squash(t)));
    execFileSync("pdftoppm", ["-r", "60", "-png", pdf, join(outDir, "slide")]);

    let ok = pages === result.slideCount && missing.length === 0;
    let detail = missing.length ? `, titles missing from the render: ${missing.join(" | ")}` : "";
    if (render) {
      const scores = await visualScores(chromium, outDir, result.screenshots);
      const low = scores.filter((s) => s.score < VISUAL_THRESHOLD);
      ok &&= low.length === 0;
      detail += `, visual ${scores.map((s) => s.score.toFixed(2)).join(" ")}`;
      if (low.length) detail += ` (below ${VISUAL_THRESHOLD}: slide ${low.map((s) => s.slide).join(", ")})`;
    }
    failed ||= !ok;
    console.log(`${ok ? "ok  " : "FAIL"} ${name}: ${pages}/${result.slideCount} pages${detail} -> ${outDir}`);
  }
  if (failed) process.exitCode = 1;
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
