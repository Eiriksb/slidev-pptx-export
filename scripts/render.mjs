#!/usr/bin/env node
// LibreOffice render check. Exports decks, converts them to PDF and per-slide PNGs for visual
// review, and checks page count and slide titles with poppler. Skips (exit 0) when soffice
// isn't installed. Usage: node scripts/render.mjs [--render] [path/to/slides.md ...]
// Default decks: the kitchen-sink fixture, plus ../innotek-SlideDevBase if it's checked out.
// --render exports in render mode and adds a visual check: each slide's LibreOffice render is
// compared with Slidev's own screenshot of it. A slide passes only when both scores clear their
// thresholds (see compare.mjs), and every slide of the deck has to be compared.

import { execFileSync, spawnSync } from "node:child_process";
import { existsSync, mkdirSync, rmSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import { resolve, dirname, basename, join } from "node:path";
import { pathToFileURL, fileURLToPath } from "node:url";
import { convert, mergeConfig } from "../src/index.mjs";
import { splitSlides, parseSlide } from "../src/parse.mjs";
import { loadChromium } from "../src/render/browser.mjs";
import { slidevCli } from "../src/render/server.mjs";
import { compareImages } from "../src/render/compare.mjs";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const DEFAULT_DECKS = [join(ROOT, "test/fixtures/kitchen-sink/slides.md"), resolve(ROOT, "../innotek-SlideDevBase/slides.md")];

// Gate 1, the mean-difference score. The spec's metric; it catches a wrong background, but not a
// blank or shifted slide (a blank page scores 0.95-1.00 against every slide of the innotek deck).
const VISUAL_THRESHOLD = 0.9;
// Gate 2, the ink overlap. Tuned once, on ../innotek-SlideDevBase (LibreOffice's 60 dpi pages
// against Slidev's screenshots): the real slides score 0.46-0.70; a blank page 0.00; another
// slide's page at most 0.18; a 20px vertical or diagonal shift of the page mostly 0.02-0.40 (but
// 0.54 once, on the chart slide shifted up); a 20px horizontal shift barely registers (0.30-0.59).
// 0.40 sits in the gap, with some room below the real scores for LibreOffice's font substitution.
const INK_THRESHOLD = 0.4;

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

// LibreOffice's slide-N.png pages against Slidev's screenshots, one result per exported slide:
// { slide, similarity, ink }, or { slide, fallback } for a slide that fell back to the markdown
// export (no screenshot), or { slide, missing } when LibreOffice drew no page for it.
async function visualScores(browser, outDir, screenshots) {
  const pages = readdirSync(outDir).filter((f) => /^slide-\d+\.png$/.test(f)).sort((a, b) => parseInt(a.slice(6), 10) - parseInt(b.slice(6), 10));
  const page = await browser.newPage();
  try {
    const results = [];
    for (const [i, shot] of screenshots.entries()) {
      const slide = i + 1;
      if (!shot) results.push({ slide, fallback: true });
      else if (!pages[i]) results.push({ slide, missing: true });
      else {
        writeFileSync(join(outDir, `slidev-${slide}.png`), shot);
        results.push({ slide, ...(await compareImages(page, shot, readFileSync(join(outDir, pages[i])))) });
      }
    }
    return results;
  } finally {
    await page.close();
  }
}

const scoreText = (r) => (r.fallback ? `${r.slide}:fallback` : r.missing ? `${r.slide}:missing` : `${r.slide}:${r.similarity.toFixed(2)}/${r.ink.toFixed(2)}`);

// First line of an error message, without its closing period, for "<why> - skipping ..." lines.
const reason = (err) => String(err.message || err).split("\n")[0].replace(/[.:]\s*$/, "");

async function main() {
  if (!has("soffice")) {
    console.log("soffice (LibreOffice) not found - skipping render check.");
    return;
  }
  const noPoppler = ["pdfinfo", "pdftotext", "pdftoppm"].filter((cmd) => !has(cmd));
  if (noPoppler.length) {
    console.log(`${noPoppler.join(", ")} (poppler) not found - skipping render check.`);
    return;
  }
  const argv = process.argv.slice(2);
  const render = argv.includes("--render");
  const args = argv.filter((a) => a !== "--render").map((p) => resolve(p));
  // Launched, not just resolved: a missing browser binary only shows up when Chromium starts.
  let browser = null;
  if (render) {
    let chromium;
    try {
      chromium = await loadChromium(ROOT);
    } catch (err) {
      console.log(`${reason(err)} - skipping the render-mode check.`);
      return;
    }
    try {
      browser = await chromium.launch();
    } catch (err) {
      console.log(`Chromium could not be launched (${reason(err)}) - skipping the render-mode check.`);
      return;
    }
  }
  try {
    await checkDecks(browser, render, args);
  } finally {
    await browser?.close();
  }
}

async function checkDecks(browser, render, args) {
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
      const results = await visualScores(browser, outDir, result.screenshots);
      const compared = results.filter((r) => r.ink !== undefined);
      const low = compared.filter((r) => r.similarity < VISUAL_THRESHOLD || r.ink < INK_THRESHOLD);
      ok &&= compared.length > 0 && compared.length === result.slideCount && low.length === 0;
      detail += `, visual ${results.map(scoreText).join(" ")}`;
      if (low.length) detail += ` (below ${VISUAL_THRESHOLD}/${INK_THRESHOLD}: slide ${low.map((r) => r.slide).join(", ")})`;
      if (compared.length < result.slideCount) detail += ` (compared ${compared.length} of ${result.slideCount} slides)`;
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
