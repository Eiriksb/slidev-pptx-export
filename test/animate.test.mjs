import { test } from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { animateCoveredVideos, coveredVideos, hasFfmpeg, videoGif } from "../src/render/animate.mjs";

const video = (extra = {}) => ({ kind: "video", data: Buffer.from("mp4"), muted: true, loop: true, duration: 9, rect: { x: 0, y: 0, w: 980, h: 552 }, natural: { w: 1600, h: 900 }, label: "video", ...extra });
const text = (rect) => ({ kind: "text", rect, runs: [{ text: "Title" }] });

test("a muted video with text drawn after it is covered", () => {
  const v = video();
  assert.deepEqual(coveredVideos([v, text({ x: 50, y: 200, w: 300, h: 40 })]), [v]);
});

test("content before the video (beneath it), elsewhere, or an empty box does not cover it", () => {
  const v = video({ rect: { x: 500, y: 0, w: 480, h: 270 } });
  assert.deepEqual(coveredVideos([text({ x: 600, y: 10, w: 100, h: 20 }), v]), []);
  assert.deepEqual(coveredVideos([v, text({ x: 10, y: 400, w: 100, h: 20 })]), []);
  assert.deepEqual(coveredVideos([v, { kind: "box", rect: { x: 500, y: 0, w: 480, h: 270 } }]), []);
  assert.deepEqual(coveredVideos([v, { kind: "box", rect: { x: 500, y: 0, w: 480, h: 270 }, fill: "#000" }]), [v]);
});

test("a video with sound stays a video", () => {
  assert.deepEqual(coveredVideos([video({ muted: false }), text({ x: 50, y: 200, w: 300, h: 40 })]), []);
});

test("a covered video becomes an animated picture in the same box", () => {
  const v = video({ fit: "cover" });
  const calls = [];
  const done = animateCoveredVideos({ nodes: [v, text({ x: 50, y: 200, w: 300, h: 40 })] }, { gif: (data, o) => (calls.push(o), Buffer.from("GIF89a")) });
  assert.equal(v.kind, "image");
  assert.equal(v.ext, "gif");
  assert.equal(v.fit, "cover");
  assert.deepEqual(v.rect, { x: 0, y: 0, w: 980, h: 552 });
  assert.equal(calls[0].loop, true);
  assert.equal(calls[0].seconds, 9);
  assert.equal(calls.length, 1);
  assert.deepEqual(v.natural, { w: 800, h: 450 });
  assert.deepEqual(done, [{ label: "video", bytes: 6, width: 800 }]);
});

test("a GIF over the budget is made again smaller and choppier", () => {
  const v = video();
  const calls = [];
  const done = animateCoveredVideos({ nodes: [v, text({ x: 50, y: 200, w: 300, h: 40 })] }, { maxBytes: 10, gif: (data, o) => (calls.push(o), Buffer.alloc(o.width >= 720 ? 100 : 5)) });
  assert.deepEqual(calls.map((c) => [c.width, c.fps]), [[800, 8], [720, 6], [560, 6]]);
  assert.deepEqual(v.natural, { w: 560, h: 315 });
  assert.equal(done[0].width, 560);
});

test("a small video is never scaled up", () => {
  const v = video({ natural: { w: 640, h: 360 } });
  const calls = [];
  animateCoveredVideos({ nodes: [v, text({ x: 50, y: 200, w: 300, h: 40 })] }, { gif: (data, o) => (calls.push(o), Buffer.from("GIF89a")) });
  assert.equal(calls[0].width, 640);
});

test("a failed conversion leaves the video as it was", () => {
  const v = video();
  const done = animateCoveredVideos({ nodes: [v, text({ x: 50, y: 200, w: 300, h: 40 })] }, { gif: () => { throw new Error("no"); } });
  assert.equal(v.kind, "video");
  assert.equal(done[0].error, "no");
});

test("videoGif makes a looping GIF with ffmpeg", { skip: !hasFfmpeg() }, () => {
  const r = spawnSync("ffmpeg", ["-v", "error", "-f", "lavfi", "-i", "testsrc=size=320x180:rate=25", "-t", "1", "-pix_fmt", "yuv420p", "-f", "mp4", "-movflags", "frag_keyframe+empty_moov", "pipe:1"], { maxBuffer: 1 << 24 });
  const gif = videoGif(r.stdout, { width: 160, fps: 10 });
  assert.equal(gif.subarray(0, 6).toString(), "GIF89a");
  assert.equal(gif.readUInt16LE(6), 160);
  assert.ok(gif.includes(Buffer.from("NETSCAPE2.0")), "loops");
});
