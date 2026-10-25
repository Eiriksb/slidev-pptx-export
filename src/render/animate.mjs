// Videos LibreOffice Impress can't show as videos, as looping animated GIFs. Its slide show plays a
// video in a window of its own above the whole slide, so slide content drawn over a video (a cover's
// title over its background clip) is gone while it plays; and it shows one video per slide, so a
// second one is not there at all. An animated picture stays where it is in the stacking order, in
// PowerPoint, Impress, Keynote and Google Slides alike.
//
// A GIF keeps the video's own frame rate (up to 30 a second), so it moves as smoothly. What it costs
// is loading time: Impress decodes every frame before it shows the slide, about 19 million pixels a
// second (a 10 s, 25 fps GIF 800 px wide held the slide back 5 s), so each GIF gets a budget of
// pixels x frames (~1.5 s). A background clip (content drawn over it) loops over a shorter stretch
// first, then gets narrower; a clip that is content of its own keeps its whole length and gets
// narrower. FFmpeg makes the GIF (a palette per clip, no dithering: noise is what LZW packs worst);
// without it the videos stay videos.

import { spawnSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

const intersects = (a, b) => a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h;

// Whether a node draws anything over what is beneath it (an empty box doesn't).
function paints(n) {
  if (!n.rect || n.kind === "missing") return false;
  if (n.kind !== "box") return true;
  return !!(n.fill || n.border || n.sides?.length || n.shadow);
}

// The muted videos that a later node in DOM order - one drawn above them - overlaps. A video
// with sound stays a video: a GIF has none.
export function coveredVideos(nodes) {
  return nodes.filter((n, i) => n.kind === "video" && n.data && n.muted && nodes.slice(i + 1).some((m) => paints(m) && intersects(m.rect, n.rect)));
}

// The videos a slide's GIFs replace: the covered ones (background: true), then, of the videos left,
// every muted one but the longest (the one that stays a video: the costliest as a GIF).
export function videosToAnimate(nodes) {
  const covered = new Set(coveredVideos(nodes));
  const left = nodes.filter((n) => n.kind === "video" && n.data && !covered.has(n));
  const keep = left.reduce((best, n) => (!best || (n.duration || 0) > (best.duration || 0) ? n : best), null);
  return [
    ...[...covered].map((node) => ({ node, background: true })),
    ...left.filter((n) => n !== keep && n.muted).map((node) => ({ node, background: false })),
  ];
}

let ffmpegFound;
export function hasFfmpeg(bin = "ffmpeg") {
  if (ffmpegFound === undefined) ffmpegFound = spawnSync(bin, ["-version"], { stdio: "ignore" }).status === 0;
  return ffmpegFound;
}

function withFile(data, fn) {
  const dir = mkdtempSync(join(tmpdir(), "pptx-gif-"));
  try {
    const src = join(dir, "in.mp4");
    writeFileSync(src, data);
    return fn(src, dir);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

// A video's frame rate, from ffmpeg's description of its stream (null when it gives none).
export function frameRate(data, bin = "ffmpeg") {
  return withFile(data, (src) => {
    const r = spawnSync(bin, ["-hide_banner", "-i", src], { encoding: "utf8" });
    const m = /Video:.*?([\d.]+) fps/.exec(r.stderr || "");
    return m ? Number(m[1]) : null;
  });
}

// GIF bytes for a video: `width` pixels wide, `fps` frames a second, the first `seconds`, looping
// forever when `loop` (else played once).
export function videoGif(data, { width, fps, seconds = 20, loop = true, bin = "ffmpeg" }) {
  return withFile(data, (src, dir) => {
    const out = join(dir, "out.gif");
    const vf = `fps=${fps},scale=${Math.round(width)}:-2:flags=lanczos,split[a][b];[a]palettegen=stats_mode=diff[p];[b][p]paletteuse=dither=none:diff_mode=rectangle`;
    const r = spawnSync(bin, ["-v", "error", "-y", "-t", String(seconds), "-i", src, "-filter_complex", vf, "-loop", loop ? "0" : "-1", out], { encoding: "utf8" });
    if (r.status !== 0) throw new Error((r.stderr || r.error?.message || "ffmpeg failed").trim().split("\n")[0]);
    return readFileSync(out);
  });
}

// Width, frame rate and length of a GIF within `budget` pixels x frames. `shown`: the GIF pixels
// worth having (the box's width at `scale`, never wider than the video). A background loops over a
// shorter stretch, down to `minSeconds`, before it gets narrower, down to `minShare` of `shown`;
// content keeps its length and gets narrower, down to `minShare`, then drops frames (15 a second at
// the fewest) and, past that, length.
export function gifPlan({ shown, aspect, duration, rate, background, budget = 30e6, maxSeconds = 20, minSeconds = 4, minShare = 0.6 }) {
  const fps = Math.round(Math.min(rate || 25, 30));
  const length = Math.min(duration || maxSeconds, maxSeconds);
  const cost = (w, f, s) => w * (w / aspect) * f * s;
  const widthFor = (f, s) => Math.sqrt((budget * aspect) / (f * s));
  const plan = (width, f, s) => ({ width: Math.max(2, Math.round(Math.min(width, shown))), fps: f, seconds: Math.round(Math.min(s, length) * 100) / 100 });
  if (cost(shown, fps, length) <= budget) return plan(shown, fps, length);
  const narrowest = shown * minShare;
  if (background) {
    const s = budget / cost(shown, fps, 1);
    if (s >= minSeconds) return plan(shown, fps, s);
    const short = Math.min(minSeconds, length);
    const w = widthFor(fps, short);
    if (w >= narrowest) return plan(w, fps, short);
    return plan(narrowest, fps, Math.max(1, budget / cost(narrowest, fps, 1)));
  }
  const w = widthFor(fps, length);
  if (w >= narrowest) return plan(w, fps, length);
  const f = Math.floor(budget / cost(narrowest, 1, length));
  if (f >= 15) return plan(narrowest, Math.min(f, fps), length);
  return plan(narrowest, 15, Math.max(1, budget / cost(narrowest, 15, 1)));
}

// Turns a slide's videos that Impress can't show (videosToAnimate) into animated pictures in place
// (same box, same object-fit crop). `scale`: GIF pixels per CSS pixel of the part shown. Returns what
// was done, for notices.
export function animateVideos(scene, { scale = 1.25, budget, gif = videoGif, rate = frameRate } = {}) {
  const done = [];
  for (const { node, background } of videosToAnimate(scene.nodes)) {
    // the part of the video in view under object-fit: cover is narrower than the whole width
    const aspect = node.natural.w / node.natural.h;
    const shown = Math.min(node.natural.w, node.rect.w * scale * Math.max(1, aspect / (node.rect.w / node.rect.h)));
    try {
      const p = gifPlan({ shown, aspect, duration: node.duration, rate: rate(node.data), background, ...(budget ? { budget } : {}) });
      const data = gif(node.data, { ...p, loop: node.loop !== false });
      const natural = { w: p.width, h: Math.round(p.width / aspect) };
      Object.assign(node, { kind: "image", data, ext: "gif", format: "gif", natural, radius: node.radius || 0 });
      done.push({ label: node.label, background, bytes: data.length, ...p, cut: !!node.duration && p.seconds < Math.min(node.duration, 20) - 0.05 });
    } catch (err) {
      done.push({ label: node.label, background, error: String(err.message || err) });
    }
  }
  return done;
}
