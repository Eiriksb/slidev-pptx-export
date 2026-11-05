// Turns a slide's click states into PowerPoint slides for an animated export. A click with no script
// motion is a native Fade on the current still slide. A click with script motion is a slide of its own
// holding only a clip, because on PowerPoint for the web native shapes above a later clip vanish. The
// next reveal opens a new still slide in its own state: the key press that reached it was that reveal.

import { createHash } from "node:crypto";
import { MOTION_FRAMES } from "./record.mjs";

// Fields that say nothing about how a node looks: `above` lists what the page draws over it, and ids
// and labels differ between page loads.
const IGNORED = new Set(["id", "above", "label"]);

const half = (v) => Math.round(v * 2) / 2;

function comparable(value) {
  if (Buffer.isBuffer(value)) return `sha1:${createHash("sha1").update(value).digest("hex")}`;
  if (Array.isArray(value)) return value.map(comparable);
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.keys(value)
        .filter((k) => !IGNORED.has(k))
        .sort()
        .map((k) => [k, comparable(value[k])]),
    );
  }
  return value;
}

// What makes two nodes of one slide's states the same element, for every kind of node. Position counts
// to half a pixel.
export function stateSignature(node) {
  const { rect, ...rest } = node;
  return JSON.stringify(comparable({ ...rest, rect: rect && { x: half(rect.x), y: half(rect.y), w: half(rect.w), h: half(rect.h) } }));
}

// Signatures numbered within a state, so two identical elements stay two.
function numbered(nodes) {
  const seen = new Map();
  return nodes.map((node) => {
    const sig = stateSignature(node);
    const n = seen.get(sig) || 0;
    seen.set(sig, n + 1);
    return { node, key: `${sig}#${n}` };
  });
}

// One z-order for all of `states`, as runs { node, from, to } of consecutive states. A node keeps its
// run only while it still comes after the node beneath it; otherwise it is repeated as a new run, so
// every state paints in its own order.
export function mergeRuns(states) {
  const merged = [];
  let open = new Map();
  states.forEach((nodes, i) => {
    const next = new Map();
    let below = null;
    for (const { node, key } of numbered(nodes)) {
      let run = open.get(key);
      if (run && (!below || merged.indexOf(run) > merged.indexOf(below))) {
        run.to = i;
      } else {
        run = { node, from: i, to: i };
        merged.splice(below ? merged.indexOf(below) + 1 : 0, 0, run);
      }
      next.set(key, run);
      below = run;
    }
    open = next;
  });
  return merged;
}

// Whether a click brought a picture that wasn't there: the only clicks worth recording.
export function pictureChanged(before, after) {
  const known = new Set(before.filter((n) => n.kind === "raster").map(stateSignature));
  return after.some((n) => n.kind === "raster" && !known.has(stateSignature(n)));
}

// What a click becomes. A `reason` means it had motion but is a reveal anyway, which the caller reports.
export function clickKind({ changed, record, encoder }) {
  const reveal = (reason = null) => ({ kind: "reveal", reason });
  if (!changed || record.motionFrames < MOTION_FRAMES) return reveal();
  if (!encoder) return reveal("no encoder to make a clip");
  if (!record.settled) return reveal("the page never settled");
  if (!record.changedVisibly) return reveal("the recording shows no visible change");
  return { kind: "motion", reason: null };
}

// The slide's PowerPoint slides, in order: { kind: "still", first, last, runs } showing states first to
// last natively, or { kind: "clip", click }. `kinds[k]` is click k's kind; kinds[0] is unused.
export function planSlide(states, kinds) {
  const segments = [];
  let still = { kind: "still", first: 0, last: 0 };
  segments.push(still);
  for (let k = 1; k < states.length; k++) {
    if (kinds[k] === "motion") {
      segments.push({ kind: "clip", click: k });
      still = null;
    } else if (still) {
      still.last = k;
    } else {
      still = { kind: "still", first: k, last: k };
      segments.push(still);
    }
  }
  for (const s of segments) {
    if (s.kind === "still") s.runs = mergeRuns(states.slice(s.first, s.last + 1)).map((r) => ({ ...r, from: r.from + s.first, to: r.to + s.first }));
  }
  return segments;
}

// A run's click effects on its still segment: it enters on the click that first shows it and leaves on
// the click after it last shows.
export function effectsOf(run, segment) {
  const effects = [];
  if (run.from > segment.first) effects.push({ click: run.from - segment.first, dir: "in" });
  if (run.to < segment.last) effects.push({ click: run.to + 1 - segment.first, dir: "out" });
  return effects;
}
