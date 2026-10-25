#!/usr/bin/env node
// Usage: slidev-pptx-export [input.md] [output.pptx] [--config path/to/pptx.config.mjs] [--render] [--embed-fonts]
//                           [--videos animated|native]
//
// Without --config, looks for pptx.config.mjs next to input.md, falling back to the built-in
// generic defaults (plain colors/fonts, no logo, no chart resolvers) if none is found.
//
// --render runs the deck with its own Slidev in headless Chromium and converts what is actually
// drawn (custom CSS, HTML and Vue components included). Needs `playwright-chromium` or
// `playwright` installed (in the deck folder or next to this package), with its Chromium.
//
// --videos (render mode): videos LibreOffice Impress can't show (slide content over them, a slide's
// second video) become looping animated pictures (animated, the default; needs ffmpeg) or stay
// videos (native: for PowerPoint only), see README.

import { existsSync } from "node:fs";
import { resolve, dirname, basename, extname, isAbsolute } from "node:path";
import { pathToFileURL } from "node:url";
import { convert } from "../src/index.mjs";

function parseArgs(argv) {
  const positional = [];
  let configPath = null;
  let render = false;
  let embedFonts = false;
  let videos = null;
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === "--config") {
      configPath = argv[++i];
    } else if (argv[i].startsWith("--config=")) {
      configPath = argv[i].slice("--config=".length);
    } else if (argv[i] === "--render") {
      render = true;
    } else if (argv[i] === "--embed-fonts") {
      embedFonts = true;
    } else if (argv[i] === "--videos" || argv[i].startsWith("--videos=")) {
      videos = argv[i].includes("=") ? argv[i].split("=")[1] : argv[++i];
      if (videos !== "animated" && videos !== "native") throw new Error(`--videos takes animated or native, not ${JSON.stringify(videos)}`);
    } else {
      positional.push(argv[i]);
    }
  }
  return { positional, configPath, render, embedFonts, videos };
}

async function loadConfig(configPath, inputPath) {
  let resolvedPath = configPath ? resolve(configPath) : resolve(dirname(inputPath), "pptx.config.mjs");
  if (!isAbsolute(resolvedPath)) resolvedPath = resolve(resolvedPath);
  if (!existsSync(resolvedPath)) {
    if (configPath) throw new Error(`Config file not found: ${resolvedPath}`);
    return { config: {}, configDir: dirname(inputPath) };
  }
  const mod = await import(pathToFileURL(resolvedPath).href);
  return { config: mod.default || {}, configDir: dirname(resolvedPath) };
}

async function main() {
  const { positional, configPath, render, embedFonts, videos } = parseArgs(process.argv.slice(2));
  const inputArg = positional[0] || "slides.md";
  const inputPath = resolve(process.cwd(), inputArg);
  if (!existsSync(inputPath)) {
    console.error(`Input file not found: ${inputPath}`);
    process.exit(1);
  }

  const { config, configDir } = await loadConfig(configPath, inputPath);
  if (videos) config.render = { ...config.render, videos };
  const outputArg = positional[1] || `${basename(inputPath, extname(inputPath))}-editable.pptx`;
  const outputPath = resolve(process.cwd(), outputArg);

  const result = await convert({ inputPath, outputPath, config, configDir, render, embedFonts });
  console.log(`Wrote ${result.slideCount} slides to ${result.outputPath}`);
  if (result.translatedCharts) console.log(`Translated ${result.translatedCharts} chart(s) into native editable PowerPoint content.`);
  if (result.unresolvedCharts.length) {
    console.log(`No resolver for: ${result.unresolvedCharts.join(", ")} - exported as placeholder box(es). Add a resolver to chartResolvers in your pptx.config.mjs to translate them too.`);
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
