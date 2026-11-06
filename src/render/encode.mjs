// Recorded frames to an H.264 MP4, with whichever tool is installed: ffmpeg with libx264, else
// GStreamer with x264enc and mp4mux. Playwright's bundled ffmpeg writes only VP8 WebM, which PowerPoint
// doesn't play, and nothing here is downloaded at install time.

import { spawnSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

export const ENCODER_HINT = "install ffmpeg with libx264, or GStreamer with x264enc and mp4mux";

const CRF = "23";
const KEYFRAME_EVERY = "15";

// `run`: spawnSync, or a stand-in for tests.
export function findEncoder({ run = spawnSync } = {}) {
  const ffmpeg = run("ffmpeg", ["-hide_banner", "-encoders"], { encoding: "utf8" });
  if (ffmpeg.status === 0 && /\blibx264\b/.test(ffmpeg.stdout || "")) return { kind: "ffmpeg" };
  const has = (element) => run("gst-inspect-1.0", [element], { stdio: "ignore" }).status === 0;
  return has("x264enc") && has("mp4mux") ? { kind: "gstreamer" } : null;
}

// The command for frames `<dir>/f%05d.jpg`. GStreamer needs telling the colours after videoconvert: a
// JPEG's range otherwise reaches the clip unconverted, and it plays brighter than the pictures around
// it. It needs no h264parse when x264enc is asked for avc directly, and h264parse is often missing.
// Both tools read the frames' path as a printf pattern, so a % in the folder's own name is doubled.
export function encoderArgs(kind, { dir, count, fps, out }) {
  const frames = join(dir.replace(/%/g, "%%"), "f%05d.jpg");
  if (kind === "ffmpeg") {
    return {
      bin: "ffmpeg",
      args: ["-y", "-hide_banner", "-loglevel", "error", "-framerate", String(fps), "-i", frames, "-frames:v", String(count), "-c:v", "libx264", "-preset", "medium", "-crf", CRF, "-bf", "0", "-g", KEYFRAME_EVERY, "-pix_fmt", "yuv420p", "-movflags", "+faststart", "-an", out],
    };
  }
  return {
    bin: "gst-launch-1.0",
    args: [
      "-q", "multifilesrc", `location=${frames}`, "index=0", `stop-index=${count - 1}`, `caps=image/jpeg,framerate=${fps}/1`,
      "!", "jpegdec", "!", "videoconvert", "!", "video/x-raw,format=I420,colorimetry=bt709",
      "!", "x264enc", "speed-preset=medium", "pass=qual", `quantizer=${CRF}`, `key-int-max=${KEYFRAME_EVERY}`, "bframes=0",
      "!", "video/x-h264,profile=high,stream-format=avc,alignment=au",
      "!", "mp4mux", "faststart=true", "!", "filesink", `location=${out}`,
    ],
  };
}

// The MP4's bytes. Throws with the tool's first error line.
export function encodeClip(encoder, { dir, count, fps }, { run = spawnSync } = {}) {
  const work = mkdtempSync(join(tmpdir(), "pptx-clip-"));
  try {
    const out = join(work, "clip.mp4");
    const { bin, args } = encoderArgs(encoder.kind, { dir, count, fps, out });
    const res = run(bin, args, { encoding: "utf8", stdio: ["ignore", "ignore", "pipe"] });
    if (res.status !== 0) {
      const line = String(res.stderr || "").split("\n").find((l) => l.trim());
      throw new Error(`${bin} failed: ${line || res.error?.message || `exit status ${res.status}`}`);
    }
    return readFileSync(out);
  } finally {
    rmSync(work, { recursive: true, force: true });
  }
}
