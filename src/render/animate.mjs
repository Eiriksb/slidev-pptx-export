// A video with slide content drawn over it (a cover's background clip under its title) as a looping
// animated GIF. PowerPoint draws text over a playing video, but LibreOffice Impress plays a video in
// a window of its own above the whole slide, so the text over it is gone while it plays; an animated
// picture stays where it is in the stacking order in both, and in Keynote and Google Slides too.
// FFmpeg makes the GIF (a palette per clip, no dithering: noise is what LZW packs worst); without it
// the video stays a video. A camera moving over detailed ground changes every pixel of every frame, so
// a GIF is big: it is made smaller and choppier until it fits a size budget.

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

let ffmpegFound;
export function hasFfmpeg(bin = "ffmpeg") {
  if (ffmpegFound === undefined) ffmpegFound = spawnSync(bin, ["-version"], { stdio: "ignore" }).status === 0;
  return ffmpegFound;
}

// GIF bytes for a video: `width` pixels wide at most (never wider than the video), `fps` frames a
// second, looping forever when the video loops (else played once), at most `seconds` long.
export function videoGif(data, { width, fps = 8, seconds = 20, loop = true, bin = "ffmpeg" }) {
  const dir = mkdtempSync(join(tmpdir(), "pptx-gif-"));
  try {
    const src = join(dir, "in.mp4");
    const out = join(dir, "out.gif");
    writeFileSync(src, data);
    const vf = `fps=${fps},scale='min(${Math.round(width)},iw)':-2:flags=lanczos,split[a][b];[a]palettegen=stats_mode=diff[p];[b][p]paletteuse=dither=none:diff_mode=rectangle`;
    const r = spawnSync(bin, ["-v", "error", "-y", "-t", String(seconds), "-i", src, "-filter_complex", vf, "-loop", loop ? "0" : "-1", out], { encoding: "utf8" });
    if (r.status !== 0) throw new Error((r.stderr || r.error?.message || "ffmpeg failed").trim().split("\n")[0]);
    return readFileSync(out);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

// Width (pixels at most) and frame rate tried in turn until a GIF fits the budget.
const STEPS = [{ width: 800, fps: 8 }, { width: 720, fps: 6 }, { width: 560, fps: 6 }, { width: 480, fps: 5 }];

// Turns each covered video of a scene into an animated picture in place (same box, same object-fit
// crop): the first of STEPS (never wider than the video, nor than `scale` GIF pixels per CSS pixel of
// what is shown) whose GIF is at most `maxBytes`, else the last. Returns what was done, for notices.
export function animateCoveredVideos(scene, { scale = 1.25, maxBytes = 12 * 2 ** 20, gif = videoGif } = {}) {
  const done = [];
  for (const node of coveredVideos(scene.nodes)) {
    // the part of the video in view under object-fit: cover is narrower than the whole width
    const shown = node.rect.w * Math.max(1, node.natural.w / node.natural.h / (node.rect.w / node.rect.h));
    const seconds = Math.min(node.duration || 20, 20);
    try {
      let data;
      let width;
      for (const step of STEPS) {
        width = Math.round(Math.min(node.natural.w, shown * scale, step.width));
        data = gif(node.data, { width, fps: step.fps, loop: node.loop !== false, seconds });
        if (data.length <= maxBytes) break;
      }
      const natural = { w: width, h: Math.round((node.natural.h * width) / node.natural.w) };
      Object.assign(node, { kind: "image", data, ext: "gif", format: "gif", natural, radius: 0 });
      done.push({ label: node.label, bytes: data.length, width });
    } catch (err) {
      done.push({ label: node.label, error: String(err.message || err) });
    }
  }
  return done;
}
