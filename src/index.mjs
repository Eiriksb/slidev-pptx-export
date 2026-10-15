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
import { renderSpecs } from "./render/index.mjs";

const LOGO_EXTS = new Set([".png", ".jpg", ".jpeg", ".gif"]);

// The footer logo config with its path resolved, or null when none is configured or the file
// is missing or isn't a format PowerPoint embeds natively. Resolved before layout, because the
// slide geometry reserves footer space only for a logo that will actually be drawn.
function usableLogo(logo, configDir) {
  if (!logo || !logo.path) return null;
  const path = resolve(configDir, logo.path);
  if (!existsSync(path)) return null;
  if (!LOGO_EXTS.has(extname(path).toLowerCase())) {
    console.warn(`Warning: logo ${path} is not PNG/JPEG/GIF - skipped.`);
    return null;
  }
  return { ...logo, path };
}

/**
 * @param {object} options
 * @param {string} options.inputPath - path to the deck's .md file
 * @param {string} [options.outputPath] - defaults to "<input-basename>-editable.pptx" next to inputPath
 * @param {object} [options.config] - see default-config.mjs; typically loaded from a deck's pptx.config.mjs
 * @param {string} [options.configDir] - directory config-relative paths (logo.path) resolve against; defaults to inputPath's directory
 * @param {boolean|{screenshots?: boolean}} [options.render] - render the deck with its own Slidev in Chromium and convert what is drawn (see README "Render mode")
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

  // Render mode draws whatever logo the page shows, so the config's logo isn't used there.
  config.logo = options.render ? null : usableLogo(config.logo, configDir);

  const { globalFrontmatter, slides } = splitSlides(raw);
  const codeStore = [];
  const parsed = slides.map((s) => ({ ...parseSlide(s.content, codeStore, config), frontmatter: s.frontmatter }));
  const chartBlocks = parsed.flatMap((s) => s.blocks).filter((b) => b.type === "chart");

  let specs;
  let style;
  let rendered = null;
  if (options.render) {
    rendered = await renderSpecs({ inputPath, userConfig: options.config || {}, config, parsed, codeStore, screenshots: options.render.screenshots === true });
    ({ specs, style } = rendered);
  } else {
    specs = layoutDeck(parsed, codeStore, config);
    style = deckStyle(config);
  }

  const buffer = await writePptx(specs, {
    geometry: style.geometry,
    theme: style.theme,
    fonts: config.fonts,
    lang: config.lang,
    title: globalFrontmatter.title || basename(inputPath, extname(inputPath)),
    author: config.author,
    timestamps: config.timestamps === true,
    // Drawn bottom-left on every slide via the slide master.
    logo: config.logo ? { path: config.logo.path, box: style.geometry.logo } : null,
    masterShapes: rendered ? rendered.masterShapes : [],
  });

  const outputPath = resolve(options.outputPath || `${basename(inputPath, extname(inputPath))}-editable.pptx`);
  await writeFile(outputPath, buffer);

  // specs, not parsed slides: a long table can continue onto extra slides.
  const result = {
    outputPath,
    slideCount: specs.length,
    translatedCharts: rendered ? rendered.translatedCharts : chartBlocks.filter((b) => b.resolved).length,
    unresolvedCharts: rendered ? [] : [...new Set(chartBlocks.filter((b) => !b.resolved).map((b) => b.name))],
  };
  if (rendered && options.render.screenshots === true) result.screenshots = rendered.screenshots;
  return result;
}

export { mergeConfig } from "./config.mjs";
export { defaultConfig } from "./default-config.mjs";
