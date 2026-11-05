// A clock the recorder steps by hand, so a page's motion is drawn frame by frame at an exact rate,
// however slowly the machine paints it. Loaded before the page's own scripts; time runs as normal until
// __startVirtualClock.
(() => {
  const real = {
    now: performance.now.bind(performance),
    dateNow: Date.now.bind(Date),
    raf: window.requestAnimationFrame.bind(window),
    caf: window.cancelAnimationFrame.bind(window),
    setTimeout: window.setTimeout.bind(window),
    clearTimeout: window.clearTimeout.bind(window),
    Date: window.Date,
  };
  let virtual = null;
  let dateOffset = 0;
  let root = null;
  let lastMutation = 0;
  // Virtual callbacks have negative ids, so a real id from before the clock started still cancels.
  let nextId = -1;
  const frames = new Map();
  const timers = new Map();
  const gaveUpOn = new Set();

  performance.now = () => virtual ?? real.now();
  window.Date = class extends real.Date {
    constructor(...args) {
      if (args.length === 0 && virtual !== null) super(virtual + dateOffset);
      else super(...args);
    }
    static now() {
      return virtual !== null ? virtual + dateOffset : real.dateNow();
    }
  };
  window.requestAnimationFrame = (cb) => {
    if (virtual === null) return real.raf(cb);
    const id = nextId--;
    frames.set(id, cb);
    return id;
  };
  window.cancelAnimationFrame = (id) => (id < 0 ? frames.delete(id) : real.caf(id));
  window.setTimeout = (fn, ms = 0, ...args) => {
    if (virtual === null) return real.setTimeout(fn, ms, ...args);
    const id = nextId--;
    timers.set(id, { fn, args, due: virtual + Number(ms || 0) });
    return id;
  };
  window.clearTimeout = (id) => (id < 0 ? timers.delete(id) : real.clearTimeout(id));

  const loadingImages = () => [...root.querySelectorAll("img")].filter((img) => !img.complete && !gaveUpOn.has(img));

  window.__startVirtualClock = (rootSelector) => {
    root = document.querySelector(rootSelector);
    virtual = real.now();
    dateOffset = real.dateNow() - virtual;
    lastMutation = virtual;
    new MutationObserver(() => {
      lastMutation = virtual;
    }).observe(root, { subtree: true, childList: true, attributes: true, characterData: true });
  };

  // The images loading now no longer count: one that never answers mustn't hold up the recording.
  window.__giveUpOnImages = () => {
    for (const img of loadingImages()) gaveUpOn.add(img);
  };

  // Advances the clock: due timers, then animation-frame callbacks, then CSS animations. An animation
  // that never ends is drawn but isn't waited for.
  window.__stepClock = (dt) => {
    virtual += dt;
    for (const [id, timer] of [...timers].sort((a, b) => a[1].due - b[1].due)) {
      if (timer.due > virtual) continue;
      timers.delete(id);
      if (typeof timer.fn === "function") timer.fn(...timer.args);
    }
    const due = [...frames.values()];
    frames.clear();
    for (const cb of due) cb(virtual);
    let running = 0;
    for (const animation of document.getAnimations()) {
      animation.pause();
      const end = animation.effect?.getComputedTiming().endTime ?? 0;
      const at = (animation.currentTime ?? 0) + dt;
      if (end !== Infinity && at >= end) {
        animation.finish();
      } else {
        animation.currentTime = at;
        if (end !== Infinity) running++;
      }
    }
    return { at: virtual, frames: frames.size, timers: timers.size, running, images: loadingImages().length, sinceMutation: virtual - lastMutation };
  };
})();
