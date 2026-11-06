// What a stopped export must clean up (SIGTERM/SIGINT, see bin/cli.mjs): whoever starts something
// that outlives the export's own code - the Slidev child, the browser, a `.partial` file - registers
// how to undo it here, and unregisters once it is done with it.
const cleanups = new Set();
let stopping = false;

// Whether a stop has begun. From then on nothing new may be written (see writeAtomic).
export const isStopping = () => stopping;

// Registers `fn` (may return a promise); the returned function unregisters it. Once a stop has
// begun the registry is closed: `fn` runs at once, since runStop will not see it.
export function onStop(fn) {
  if (stopping) {
    try {
      Promise.resolve(fn()).catch(() => {});
    } catch {
      // stopping anyway
    }
    return () => {};
  }
  cleanups.add(fn);
  return () => cleanups.delete(fn);
}

// Runs every registered cleanup, all started together, each one's failure ignored, and returns once
// they have finished or after `deadline` ms, whichever comes first.
export async function runStop(deadline = 5000) {
  stopping = true;
  const fns = [...cleanups];
  cleanups.clear();
  const all = Promise.all(fns.map(async (fn) => {
    try {
      await fn();
    } catch {
      // stopping anyway
    }
  }));
  let timer;
  await Promise.race([all, new Promise((r) => (timer = setTimeout(r, deadline)))]);
  clearTimeout(timer);
}
