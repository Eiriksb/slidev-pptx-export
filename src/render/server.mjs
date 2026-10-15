// Render mode runs the deck's own Slidev (its version, theme and addons) as a dev server on a
// free local port, and stops it again afterwards.

import { spawn } from "node:child_process";
import { existsSync, readFileSync } from "node:fs";
import { createServer } from "node:net";
import { join, dirname, basename, resolve as resolvePath } from "node:path";

// @slidev/cli as the deck sees it: in its own node_modules or any parent's (workspaces hoist it).
export function slidevCli(deckDir) {
  for (let dir = resolvePath(deckDir); ; dir = dirname(dir)) {
    const pkgPath = join(dir, "node_modules", "@slidev", "cli", "package.json");
    if (existsSync(pkgPath)) {
      const pkg = JSON.parse(readFileSync(pkgPath, "utf8"));
      const rel = typeof pkg.bin === "string" ? pkg.bin : pkg.bin?.slidev;
      if (rel) return { bin: join(dirname(pkgPath), rel), version: pkg.version };
    }
    if (dirname(dir) === dir) return null;
  }
}

function freePort() {
  return new Promise((resolve, reject) => {
    const s = createServer();
    s.unref();
    s.on("error", reject);
    s.listen(0, "127.0.0.1", () => {
      const { port } = s.address();
      s.close(() => resolve(port));
    });
  });
}

export async function startDeckServer({ inputPath, timeout = 60000 }) {
  const deckDir = dirname(inputPath);
  const cli = slidevCli(deckDir);
  if (!cli) throw new Error(`render mode runs the deck's own Slidev, but @slidev/cli isn't installed for this deck - run npm install in ${deckDir}`);
  const port = await freePort();
  const child = spawn(process.execPath, [cli.bin, basename(inputPath), "--port", String(port), "--bind", "127.0.0.1"], {
    cwd: deckDir,
    stdio: ["ignore", "pipe", "pipe"],
    env: { ...process.env, NO_COLOR: "1" },
  });
  const lines = [];
  const keep = (chunk) => {
    lines.push(...String(chunk).split(/\r?\n/));
    if (lines.length > 50) lines.splice(0, lines.length - 50);
  };
  child.stdout.on("data", keep);
  child.stderr.on("data", keep);
  let spawnError = null;
  child.on("error", (err) => {
    spawnError = err;
  });
  let exited = null;
  // "close" (not "exit") fires after the output streams are drained.
  child.on("close", (code, signal) => {
    exited = code ?? signal ?? "unknown";
  });
  const log = () => lines.filter((l) => l.trim()).join("\n");
  const stop = () =>
    new Promise((resolve) => {
      if (exited !== null) return resolve();
      child.once("close", () => resolve());
      child.kill();
      setTimeout(() => child.kill("SIGKILL"), 3000).unref();
    });

  const url = `http://127.0.0.1:${port}`;
  const deadline = Date.now() + timeout;
  while (Date.now() < deadline) {
    if (spawnError) throw new Error(`Slidev could not be started: ${spawnError.message}`);
    if (exited !== null) throw new Error(`Slidev exited before it was ready (code ${exited}):\n${log()}`);
    try {
      if ((await fetch(`${url}/`)).ok) return { url, stop, log, version: cli.version };
    } catch {
      // not listening yet
    }
    await new Promise((r) => setTimeout(r, 250));
  }
  await stop();
  throw new Error(`Slidev didn't start within ${timeout / 1000}s:\n${log()}`);
}
