// Atomic file write: the bytes go to `<path>.partial` and are renamed to `path` only once
// complete, so a stopped or failed export never leaves a half-written file under its name.
import { writeFile, rename, rm } from "node:fs/promises";

const defaultFs = { writeFile, rename, rm };

export async function writeAtomic(path, data, { fs = defaultFs } = {}) {
  const partial = `${path}.partial`;
  try {
    await fs.writeFile(partial, data);
    await fs.rename(partial, path);
  } catch (err) {
    await fs.rm(partial, { force: true }).catch(() => {});
    throw err;
  }
}
