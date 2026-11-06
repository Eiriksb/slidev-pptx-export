import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { ENCODER_HINT, encodeClip, encoderArgs, findEncoder } from "../src/render/encode.mjs";

// A 16x16 white JPEG.
const TINY_JPEG = Buffer.from(
  "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAA0JCgsKCA0LCgsODg0PEyAVExISEyccHhcgLikxMC4pLSwzOko+MzZGNywtQFdBRkxOUlNSMj5aYVpQYEpRUk//2wBDAQ4ODhMREyYVFSZPNS01T09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0//wAARCAAQABADAREAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD0KgAoAKACgD//2Q==",
  "base64",
);

function makeFrames(t, count, prefix = "pptx-encode-test-") {
  const dir = mkdtempSync(join(tmpdir(), prefix));
  t.after(() => rmSync(dir, { recursive: true, force: true }));
  for (let i = 0; i < count; i++) writeFileSync(join(dir, `f${String(i).padStart(5, "0")}.jpg`), TINY_JPEG);
  return { dir, count, fps: 30 };
}

test("ffmpeg with libx264 is preferred, then GStreamer with x264enc and mp4mux, then nothing", () => {
  const fake = (present) => (bin) => ({ status: present[bin] ? 0 : 1, stdout: bin === "ffmpeg" && present.x264 ? " V..... libx264 " : "" });
  assert.deepEqual(findEncoder({ run: fake({ ffmpeg: true, x264: true }) }), { kind: "ffmpeg", bin: "ffmpeg" });
  assert.deepEqual(findEncoder({ run: fake({ ffmpeg: true, "gst-inspect-1.0": true }) }), { kind: "gstreamer" });
  assert.equal(findEncoder({ run: fake({}) }), null);
});

test("a named ffmpeg with libx264 is used before the one on PATH", () => {
  const seen = [];
  const run = (bin) => { seen.push(bin); return { status: 0, stdout: " V..... libx264 " }; };
  assert.deepEqual(findEncoder({ run, ffmpeg: "/opt/ff/ffmpeg" }), { kind: "ffmpeg", bin: "/opt/ff/ffmpeg" });
  assert.equal(seen[0], "/opt/ff/ffmpeg");
});

test("a named ffmpeg that is missing or lacks libx264 gives way to the one on PATH", () => {
  const seen = [];
  const run = (bin) => { seen.push(bin); return bin === "ffmpeg" ? { status: 0, stdout: " V..... libx264 " } : { status: 0, stdout: " V..... mpeg4 " }; };
  assert.deepEqual(findEncoder({ run, ffmpeg: "/opt/ff/ffmpeg" }), { kind: "ffmpeg", bin: "ffmpeg" });
  assert.deepEqual(seen, ["/opt/ff/ffmpeg", "ffmpeg"]);
  const missing = (bin) => (bin === "ffmpeg" ? { status: 0, stdout: "libx264" } : { status: null, error: new Error("ENOENT"), stdout: null });
  assert.deepEqual(findEncoder({ run: missing, ffmpeg: "/nope" }), { kind: "ffmpeg", bin: "ffmpeg" });
});

test("ffmpeg's arguments start with the named program", () => {
  assert.equal(encoderArgs("ffmpeg", { dir: "/d", count: 3, fps: 30, out: "/o.mp4", bin: "/opt/ff/ffmpeg" }).bin, "/opt/ff/ffmpeg");
});

test("a clip is encoded with the encoder's own program", (t) => {
  const seen = [];
  const run = (bin) => { seen.push(bin); return { status: 1, stderr: "stop" }; };
  assert.throws(() => encodeClip({ kind: "ffmpeg", bin: "/opt/ff/ffmpeg" }, makeFrames(t, 1), { run }), /\/opt\/ff\/ffmpeg failed: stop/);
  assert.deepEqual(seen, ["/opt/ff/ffmpeg"]);
});

test("GStreamer counts only with both elements", () => {
  const without = (element) => (bin, args) => ({ status: bin === "gst-inspect-1.0" && args[0] !== element ? 0 : 1, stdout: "" });
  assert.equal(findEncoder({ run: without("x264enc") }), null);
  assert.equal(findEncoder({ run: without("mp4mux") }), null);
});

test("a tool that isn't there counts as missing, not as an error", () => {
  assert.equal(findEncoder({ run: () => ({ status: null, error: new Error("ENOENT"), stdout: null }) }), null);
});

test("ffmpeg arguments: H.264, CRF 23, yuv420p, no B-frames, faststart", () => {
  const { bin, args } = encoderArgs("ffmpeg", { dir: "/d", count: 3, fps: 30, out: "/o.mp4" });
  assert.equal(bin, "ffmpeg");
  for (const a of ["libx264", "23", "yuv420p", "+faststart", "/d/f%05d.jpg", "/o.mp4"]) assert.ok(args.includes(a), a);
  assert.equal(args[args.indexOf("-bf") + 1], "0");
  assert.equal(args[args.indexOf("-framerate") + 1], "30");
});

test("GStreamer arguments ask x264enc for avc, so h264parse is not needed", () => {
  const { bin, args } = encoderArgs("gstreamer", { dir: "/d", count: 3, fps: 30, out: "/o.mp4" });
  assert.equal(bin, "gst-launch-1.0");
  assert.ok(args.includes("video/x-h264,profile=high,stream-format=avc,alignment=au"));
  assert.ok(args.includes("quantizer=23") && args.includes("pass=qual"));
  assert.ok(!args.includes("h264parse"));
  assert.ok(args.includes("stop-index=2"));
  assert.ok(args.includes("caps=image/jpeg,framerate=30/1"));
  assert.ok(args.includes("location=/o.mp4"));
});

test("GStreamer is told the frames' colours, or the clip plays brighter than its poster", () => {
  const { args } = encoderArgs("gstreamer", { dir: "/d", count: 3, fps: 30, out: "/o.mp4" });
  assert.ok(args.includes("video/x-raw,format=I420,colorimetry=bt709"));
});

test("a tool that fails: the error names its first line", (t) => {
  const run = () => ({ status: 1, stderr: "\nsomething broke\nand more\n" });
  assert.throws(() => encodeClip({ kind: "ffmpeg" }, makeFrames(t, 1), { run }), /something broke/);
});

test("the hint names both tools", () => {
  assert.match(ENCODER_HINT, /ffmpeg/);
  assert.match(ENCODER_HINT, /GStreamer/);
});

test("a real encode gives an MP4, when an encoder is installed", { skip: findEncoder() ? false : "no encoder installed" }, (t) => {
  const data = encodeClip(findEncoder(), makeFrames(t, 10));
  assert.equal(data.subarray(4, 8).toString(), "ftyp");
});

test("ffmpeg reads a % in the frames' folder as a %, not as part of its pattern", () => {
  const { args } = encoderArgs("ffmpeg", { dir: "/tmp/a%b", count: 3, fps: 30, out: "/o.mp4" });
  assert.ok(args.includes("/tmp/a%%b/f%05d.jpg"));
});

test("a real encode from a folder whose name has a %, a ! and a space, when an encoder is installed", { skip: findEncoder() ? false : "no encoder installed" }, (t) => {
  const data = encodeClip(findEncoder(), makeFrames(t, 4, "pptx 100% ok! "));
  assert.equal(data.subarray(4, 8).toString(), "ftyp");
});

