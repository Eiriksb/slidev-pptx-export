// What a stopped export must clean up (SIGTERM/SIGINT, see bin/cli.mjs): whoever starts something
// that outlives the export's own code - the Slidev child, the browser, a `.partial` file - registers
// how to undo it here, and unregisters once it is done with it.
const cleanups = new Set();

// Registers `fn` (may return a promise); the returned function unregisters it.
export function onStop(fn) {
  cleanups.add(fn);
  return () => cleanups.delete(fn);
}

// Runs every registered cleanup, newest first, each one's failure ignored.
export async function runStop() {
  const fns = [...cleanups].reverse();
  cleanups.clear();
  await Promise.all(fns.map(async (fn) => {
    try {
      await fn();
    } catch {
      // stopping anyway
    }
  }));
}
