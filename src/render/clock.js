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
    setInterval: window.setInterval.bind(window),
    clearInterval: window.clearInterval.bind(window),
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
  const intervals = new Map();
  const gaveUpOn = new Set();

  performance.now = () => virtual ?? real.now();
  // A function, not a class: Date() without new must still return a string, as it does natively.
  const nowMs = () => (virtual !== null ? virtual + dateOffset : real.dateNow());
  function VirtualDate(...args) {
    if (!new.target) return new real.Date(nowMs()).toString();
    return args.length === 0 ? new real.Date(nowMs()) : new real.Date(...args);
  }
  VirtualDate.prototype = real.Date.prototype;
  VirtualDate.now = nowMs;
  VirtualDate.parse = real.Date.parse;
  VirtualDate.UTC = real.Date.UTC;
  window.Date = VirtualDate;
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
  window.setInterval = (fn, ms = 0, ...args) => {
    if (virtual === null) return real.setInterval(fn, ms, ...args);
    const id = nextId--;
    const every = Math.max(1, Number(ms || 0));
    intervals.set(id, { fn, args, every, due: virtual + every });
    return id;
  };
  // Timeouts and intervals share their ids, so either clear function clears either, as in a browser.
  const clear = (id, realClear) => {
    if (id < 0) {
      timers.delete(id);
      intervals.delete(id);
    } else {
      realClear(id);
    }
  };
  window.clearTimeout = (id) => clear(id, real.clearTimeout);
  window.clearInterval = (id) => clear(id, real.clearInterval);

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

  // Advances the clock: due timers and intervals, then animation-frame callbacks, then CSS animations.
  // An animation that never ends is drawn but isn't waited for. `intervals` counts those the click
  // started: an interval running is script motion, as an animation frame is.
  window.__stepClock = (dt) => {
    virtual += dt;
    for (const [id, timer] of [...timers].sort((a, b) => a[1].due - b[1].due)) {
      if (timer.due > virtual) continue;
      timers.delete(id);
      if (typeof timer.fn === "function") timer.fn(...timer.args);
    }
    for (const [id, interval] of [...intervals]) {
      while (intervals.has(id) && interval.due <= virtual) {
        interval.due += interval.every;
        if (typeof interval.fn === "function") interval.fn(...interval.args);
      }
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
    return { at: virtual, frames: frames.size, timers: timers.size, intervals: intervals.size, running, images: loadingImages().length, sinceMutation: virtual - lastMutation };
  };
})();
