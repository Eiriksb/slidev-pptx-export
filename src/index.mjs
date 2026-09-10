// Public API: convert(options) turns a Slidev markdown deck into a fully editable native
// .pptx (real PowerPoint text boxes, tables, shapes and charts - not the screenshot-per-slide
// export `slidev export` produces). See README.md for the config shape and bin/cli.mjs for
// the command-line wrapper.

import { readFileSync, existsSync } from "node:fs";
import { resolve, dirname, basename, extname } from "node:path";
import PptxGenJS from "pptxgenjs";
import { mergeConfig } from "./config.mjs";
import { splitSlides, parseSlide } from "./parse.mjs";
import { renderSlide } from "./render.mjs";
import { stripPhantomSlideMasterEntries } from "./fix-pptx.mjs";

/**
 * @param {object} options
 * @param {string} options.inputPath - path to the deck's .md file
 * @param {string} [options.outputPath] - defaults to "<input-basename>-editable.pptx" next to inputPath
 * @param {object} [options.config] - see default-config.mjs; typically loaded from a deck's pptx.config.mjs
 * @param {string} [options.configDir] - directory config-relative paths (logo.path) resolve against; defaults to inputPath's directory
 * @returns {Promise<{ outputPath: string, slideCount: number, translatedCharts: number, unresolvedCharts: string[] }>}
 */
export async function convert(options) {
  const inputPath = resolve(options.inputPath);
  const raw = readFileSync(inputPath, "utf8");
  const inputDir = dirname(inputPath);
  const configDir = options.configDir ? resolve(options.configDir) : inputDir;

  const config = mergeConfig(options.config || {});
  config.inputDir = inputDir;
  config.publicDir = resolve(inputDir, "public");
  if (!existsSync(config.publicDir)) config.publicDir = null;

  const dataDir = resolve(inputDir, config.dataDir);
  config.resolvedDataDir = existsSync(dataDir) ? dataDir : null;
  if (!config.resolvedDataDir && Object.keys(config.chartResolvers).length) {
    console.warn(`Warning: data directory "${config.dataDir}" not found next to ${basename(inputPath)} - charts will fall back to placeholders.`);
  }

  if (config.logo && config.logo.path) {
    config.logo = { ...config.logo, path: resolve(configDir, config.logo.path) };
  }

  const { slides } = splitSlides(raw);

  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "DECK", width: config.slide.width, height: config.slide.height });
  pptx.layout = "DECK";
  pptx.title = basename(inputPath, extname(inputPath));
  if (config.author) pptx.author = config.author;
  // Document-wide default font for any text run that doesn't set its own fontFace (body
  // copy, bullets, table/stat-card text). Headings/eyebrow/code explicitly override this
  // per-call (see config.fonts.heading/mono).
  pptx.theme = { headFontFace: config.fonts.heading, bodyFontFace: config.fonts.body };

  const codeStore = [];
  const parsed = slides.map((s) => ({ ...parseSlide(s.content, codeStore, config), frontmatter: s.frontmatter }));
  parsed.forEach((slideData, i) => renderSlide(pptx, slideData, i, codeStore, config));

  const outputPath = resolve(options.outputPath || `${basename(inputPath, extname(inputPath))}-editable.pptx`);
  await pptx.writeFile({ fileName: outputPath });
  await stripPhantomSlideMasterEntries(outputPath);

  const chartBlocks = parsed.flatMap((s) => s.blocks).filter((b) => b.type === "chart");
  const translatedCharts = chartBlocks.filter((b) => b.resolved).length;
  const unresolvedCharts = [...new Set(chartBlocks.filter((b) => !b.resolved).map((b) => b.name))];

  return { outputPath, slideCount: parsed.length, translatedCharts, unresolvedCharts };
}

export { mergeConfig } from "./config.mjs";
export { defaultConfig } from "./default-config.mjs";
