// Atomic file write: the bytes go to `<path>.partial` and are renamed to `path` only once
// complete, so a stopped or failed export never leaves a half-written file under its name.
import { writeFile, rename, rm } from "node:fs/promises";
import { rmSync } from "node:fs";
import { onStop, isStopping } from "./stop.mjs";

const defaultFs = { writeFile, rename, rm };

export async function writeAtomic(path, data, { fs = defaultFs } = {}) {
  if (isStopping()) throw new Error("stopped");
  const partial = `${path}.partial`;
  // A stop removes the partial file at once; the real fs is the only one a stop concerns.
  const forget = fs === defaultFs ? onStop(() => rmSync(partial, { force: true })) : () => {};
  try {
    await fs.writeFile(partial, data);
    await fs.rename(partial, path);
  } catch (err) {
    await fs.rm(partial, { force: true }).catch(() => {});
    throw err;
  } finally {
    forget();
  }
}
