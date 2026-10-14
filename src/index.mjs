// Public API: convert(options) turns a Slidev markdown deck into a fully editable native
// .pptx - real PowerPoint placeholders, text, tables, shapes and charts on a themed slide
// master, not the screenshot-per-slide export `slidev export` produces. See README.md for the
// config shape and bin/cli.mjs for the command-line wrapper.

import { readFileSync, existsSync } from "node:fs";
import { writeFile } from "node:fs/promises";
import { resolve, dirname, basename, extname } from "node:path";
import { mergeConfig } from "./config.mjs";
import { splitSlides, parseSlide } from "./parse.mjs";
import { layoutDeck, deckStyle } from "./layout.mjs";
import { writePptx } from "./ooxml/index.mjs";

const LOGO_EXTS = new Set([".png", ".jpg", ".jpeg", ".gif"]);

// Footer logo drawn bottom-left on every slide (it lives on the slide master), or null when
// none is configured or the file isn't a format PowerPoint embeds natively.
function logoFor(config, geometry) {
  const logo = config.logo;
  if (!logo || !logo.path || !existsSync(logo.path)) return null;
  if (!LOGO_EXTS.has(extname(logo.path).toLowerCase())) {
    console.warn(`Warning: logo ${logo.path} is not PNG/JPEG/GIF - skipped.`);
    return null;
  }
  const h = logo.height || 0.27;
  return { path: logo.path, box: { x: 0.5, y: geometry.height - 0.29 - h, w: h * (logo.aspect || 1), h } };
}

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

  const { globalFrontmatter, slides } = splitSlides(raw);
  const codeStore = [];
  const parsed = slides.map((s) => ({ ...parseSlide(s.content, codeStore, config), frontmatter: s.frontmatter }));
  const specs = layoutDeck(parsed, codeStore, config);
  const style = deckStyle(config);

  const buffer = await writePptx(specs, {
    geometry: style.geometry,
    theme: style.theme,
    fonts: config.fonts,
    lang: config.lang,
    title: globalFrontmatter.title || basename(inputPath, extname(inputPath)),
    author: config.author,
    timestamps: config.timestamps === true,
    logo: logoFor(config, style.geometry),
  });

  const outputPath = resolve(options.outputPath || `${basename(inputPath, extname(inputPath))}-editable.pptx`);
  await writeFile(outputPath, buffer);

  const chartBlocks = parsed.flatMap((s) => s.blocks).filter((b) => b.type === "chart");
  const translatedCharts = chartBlocks.filter((b) => b.resolved).length;
  const unresolvedCharts = [...new Set(chartBlocks.filter((b) => !b.resolved).map((b) => b.name))];

  return { outputPath, slideCount: parsed.length, translatedCharts, unresolvedCharts };
}

export { mergeConfig } from "./config.mjs";
export { defaultConfig } from "./default-config.mjs";
