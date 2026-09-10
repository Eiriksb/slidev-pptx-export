#!/usr/bin/env node
// Usage: slidev-pptx-export [input.md] [output.pptx] [--config path/to/pptx.config.mjs]
//
// Without --config, looks for pptx.config.mjs next to input.md, falling back to the built-in
// generic defaults (plain colors/fonts, no logo, no chart resolvers) if none is found.

import { existsSync } from "node:fs";
import { resolve, dirname, basename, extname, isAbsolute } from "node:path";
import { pathToFileURL } from "node:url";
import { convert } from "../src/index.mjs";

function parseArgs(argv) {
  const positional = [];
  let configPath = null;
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === "--config") {
      configPath = argv[++i];
    } else if (argv[i].startsWith("--config=")) {
      configPath = argv[i].slice("--config=".length);
    } else {
      positional.push(argv[i]);
    }
  }
  return { positional, configPath };
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
  const { positional, configPath } = parseArgs(process.argv.slice(2));
  const inputArg = positional[0] || "slides.md";
  const inputPath = resolve(process.cwd(), inputArg);
  if (!existsSync(inputPath)) {
    console.error(`Input file not found: ${inputPath}`);
    process.exit(1);
  }

  const { config, configDir } = await loadConfig(configPath, inputPath);
  const outputArg = positional[1] || `${basename(inputPath, extname(inputPath))}-editable.pptx`;
  const outputPath = resolve(process.cwd(), outputArg);

  const result = await convert({ inputPath, outputPath, config, configDir });
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
