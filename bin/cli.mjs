#!/usr/bin/env node
// Usage: slidev-pptx-export [input.md] [output.pptx] [--config path/to/pptx.config.mjs] [--render] [--embed-fonts]
//                           [--videos animated|native] [--animations] [--progress json]
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
// videos (native: for PowerPoint only), see README. With --animations, videos always stay videos.
//
// --animations implies --render, and plays the deck's clicks and slide transitions in the slide show on
// PowerPoint for the web. A click whose page moves by script needs ffmpeg or GStreamer, see README.
//
// --progress json makes stdout carry only the export's progress, one JSON object per line (start, slide,
// state, record, clip, notice, write, done; or error as the last line when the export fails). What the
// export says for people (Notice:/Warning: lines, errors) stays on stderr, and each notice is also sent
// as a notice event.

import { existsSync } from "node:fs";
import { resolve, dirname, basename, extname, isAbsolute } from "node:path";
import { pathToFileURL } from "node:url";
import { format } from "node:util";
import { convert } from "../src/index.mjs";
import { runStop } from "../src/stop.mjs";

function parseArgs(argv) {
  const positional = [];
  let configPath = null;
  let render = false;
  let embedFonts = false;
  let videos = null;
  let animations = false;
  let progress = null;
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === "--config") {
      configPath = argv[++i];
    } else if (argv[i].startsWith("--config=")) {
      configPath = argv[i].slice("--config=".length);
    } else if (argv[i] === "--render") {
      render = true;
    } else if (argv[i] === "--animations") {
      animations = true;
    } else if (argv[i] === "--embed-fonts") {
      embedFonts = true;
    } else if (argv[i] === "--progress" || argv[i].startsWith("--progress=")) {
      progress = argv[i].includes("=") ? argv[i].split("=")[1] : argv[++i];
      if (progress !== "json") throw new Error(`--progress takes json, not ${JSON.stringify(progress)}`);
    } else if (argv[i] === "--videos" || argv[i].startsWith("--videos=")) {
      videos = argv[i].includes("=") ? argv[i].split("=")[1] : argv[++i];
      if (videos !== "animated" && videos !== "native") throw new Error(`--videos takes animated or native, not ${JSON.stringify(videos)}`);
    } else {
      positional.push(argv[i]);
    }
  }
  return { positional, configPath, render, embedFonts, videos, animations, progress };
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

// Whether --progress json was given, known even when parsing the arguments is what fails.
const wantsJson = (argv) => argv.some((a, i) => a === "--progress=json" || (a === "--progress" && argv[i + 1] === "json"));

// In json mode stdout is for events alone: console.log goes to stderr, and each console.warn line also
// becomes a notice event.
function jsonProgress() {
  const emit = (event, done) => process.stdout.write(`${JSON.stringify(event)}\n`, done);
  const warn = console.warn;
  console.log = console.error;
  console.warn = (...args) => {
    warn(...args);
    emit({ event: "notice", text: format(...args) });
  };
  return emit;
}

async function main(argv, emit) {
  const { positional, configPath, render, embedFonts, videos, animations, progress } = parseArgs(argv);
  const inputArg = positional[0] || "slides.md";
  const inputPath = resolve(process.cwd(), inputArg);
  if (!existsSync(inputPath)) throw new Error(`Input file not found: ${inputPath}`);

  const { config, configDir } = await loadConfig(configPath, inputPath);
  if (videos) config.render = { ...config.render, videos };
  if (animations) config.render = { ...config.render, animations: true };
  const outputArg = positional[1] || `${basename(inputPath, extname(inputPath))}-editable.pptx`;
  const outputPath = resolve(process.cwd(), outputArg);

  const result = await convert({ inputPath, outputPath, config, configDir, render, embedFonts, handleSignals: false, onProgress: progress ? (event) => emit(event) : undefined });
  console.log(`Wrote ${result.slideCount} slides to ${result.outputPath}`);
  if (result.translatedCharts) console.log(`Translated ${result.translatedCharts} chart(s) into native editable PowerPoint content.`);
  if (result.unresolvedCharts.length) {
    console.log(`No resolver for: ${result.unresolvedCharts.join(", ")} - exported as placeholder box(es). Export with --render to convert what these components draw, or add a resolver to chartResolvers in your pptx.config.mjs.`);
  }
}

const argv = process.argv.slice(2);
const emit = wantsJson(argv) ? jsonProgress() : null;

// SIGTERM or SIGINT stops the export: the Slidev child, the browser and any .partial file are cleaned
// up, then the process exits 143 / 130 (with a final "stopped" error event in json mode).
let stopping = false;
for (const [signal, code] of [["SIGTERM", 143], ["SIGINT", 130]]) {
  process.on(signal, async () => {
    const finish = () => (emit ? emit({ event: "error", message: "stopped" }, () => process.exit(code)) : process.exit(code));
    // A second signal ends it at once.
    if (stopping) return process.exit(code);
    stopping = true;
    await runStop();
    console.error("Stopped.");
    finish();
  });
}

main(argv, emit).catch((err) => {
  if (stopping) return; // the stop is already ending the process
  // A missing input file is a message, not a stack trace.
  console.error(err.message.startsWith("Input file not found") ? err.message : err);
  // Exit once the error line is flushed: stdout may be a pipe that is written asynchronously.
  if (emit) emit({ event: "error", message: err.message }, () => process.exit(1));
  else process.exit(1);
});
