// What bin/cli.mjs makes of its command line.

export function parseArgs(argv) {
  const positional = [];
  let configPath = null;
  let render = false;
  let embedFonts = false;
  let videos = null;
  let animations = null; // null: not given, the config decides
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
    } else if (argv[i] === "--no-animations") {
      animations = false;
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

// The flags that set render options win over the config: --videos, and --animations / --no-animations
// (--no-animations turns animations off even if the config turns them on).
export function applyFlags(config, { videos, animations }) {
  if (videos) config.render = { ...config.render, videos };
  if (animations !== null && animations !== undefined) config.render = { ...config.render, animations };
}
