import { test } from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { animateVideos, coveredVideos, frameRate, gifPlan, hasFfmpeg, videoGif, videosToAnimate } from "../src/render/animate.mjs";

const video = (extra = {}) => ({ kind: "video", data: Buffer.from("mp4"), muted: true, loop: true, duration: 9, rect: { x: 0, y: 0, w: 980, h: 552 }, natural: { w: 1600, h: 900 }, label: "video", ...extra });
const text = (rect) => ({ kind: "text", rect, runs: [{ text: "Title" }] });
const title = text({ x: 50, y: 200, w: 300, h: 40 });
const fake = (calls = []) => ({ gif: (data, o) => (calls.push(o), Buffer.from("GIF89a")), rate: () => 25 });

test("a muted video with text drawn after it is covered", () => {
  const v = video();
  assert.deepEqual(coveredVideos([v, title]), [v]);
});

test("content before the video (beneath it), elsewhere, or an empty box does not cover it", () => {
  const v = video({ rect: { x: 500, y: 0, w: 480, h: 270 } });
  assert.deepEqual(coveredVideos([text({ x: 600, y: 10, w: 100, h: 20 }), v]), []);
  assert.deepEqual(coveredVideos([v, text({ x: 10, y: 400, w: 100, h: 20 })]), []);
  assert.deepEqual(coveredVideos([v, { kind: "box", rect: { x: 500, y: 0, w: 480, h: 270 } }]), []);
  assert.deepEqual(coveredVideos([v, { kind: "box", rect: { x: 500, y: 0, w: 480, h: 270 }, fill: "#000" }]), [v]);
});

test("a video with sound stays a video", () => {
  assert.deepEqual(coveredVideos([video({ muted: false }), title]), []);
});

test("of a slide's uncovered videos the longest stays a video, the others are animated", () => {
  const long = video({ duration: 30, rect: { x: 0, y: 100, w: 400, h: 225 } });
  const short = video({ duration: 10, rect: { x: 500, y: 100, w: 400, h: 225 } });
  assert.deepEqual(videosToAnimate([short, long]), [{ node: short, background: false }]);
  assert.deepEqual(videosToAnimate([long]), []);
  const cover = video();
  assert.deepEqual(videosToAnimate([cover, title, long]), [{ node: cover, background: true }]);
});

test("a GIF within its budget keeps size, frame rate and length", () => {
  assert.deepEqual(gifPlan({ shown: 400, aspect: 2, duration: 5, rate: 25, background: false, budget: 1e9 }), { width: 400, fps: 25, seconds: 5 });
  assert.equal(gifPlan({ shown: 400, aspect: 2, duration: 5, rate: 60, background: false, budget: 1e9 }).fps, 30);
});

test("a background over budget loops over a shorter stretch, then gets narrower", () => {
  const p = gifPlan({ shown: 1000, aspect: 2, duration: 10, rate: 25, background: true, budget: 1000 * 500 * 25 * 6 });
  assert.deepEqual(p, { width: 1000, fps: 25, seconds: 6 });
  const q = gifPlan({ shown: 1000, aspect: 2, duration: 10, rate: 25, background: true, budget: 800 * 400 * 25 * 4 });
  assert.equal(q.seconds, 4);
  assert.equal(q.fps, 25);
  assert.equal(q.width, 800);
});

test("content over budget keeps its length and frame rate and gets narrower, then drops frames", () => {
  const p = gifPlan({ shown: 1000, aspect: 2, duration: 10, rate: 30, background: false, budget: 800 * 400 * 30 * 10 });
  assert.deepEqual(p, { width: 800, fps: 30, seconds: 10 });
  const q = gifPlan({ shown: 1000, aspect: 2, duration: 10, rate: 30, background: false, budget: 600 * 300 * 20 * 10 });
  assert.deepEqual(q, { width: 600, fps: 20, seconds: 10 });
});

test("a video becomes an animated picture in the same box", () => {
  const v = video({ fit: "cover" });
  const calls = [];
  const done = animateVideos({ nodes: [v, title] }, { ...fake(calls), budget: 1e12 });
  assert.equal(v.kind, "image");
  assert.equal(v.ext, "gif");
  assert.equal(v.fit, "cover");
  assert.deepEqual(v.rect, { x: 0, y: 0, w: 980, h: 552 });
  assert.deepEqual(calls[0], { width: 1227, fps: 25, seconds: 9, loop: true });
  assert.deepEqual(v.natural, { w: 1227, h: 690 });
  assert.equal(done[0].cut, false);
});

test("a small video is never scaled up", () => {
  const calls = [];
  animateVideos({ nodes: [video({ natural: { w: 640, h: 360 } }), title] }, { ...fake(calls), budget: 1e12 });
  assert.equal(calls[0].width, 640);
});

test("a failed conversion leaves the video as it was", () => {
  const v = video();
  const done = animateVideos({ nodes: [v, title] }, { gif: () => { throw new Error("no"); }, rate: () => 25 });
  assert.equal(v.kind, "video");
  assert.equal(done[0].error, "no");
});

test("videoGif makes a looping GIF at the frame rate asked; frameRate reads a video's", { skip: !hasFfmpeg() }, () => {
  const r = spawnSync("ffmpeg", ["-v", "error", "-f", "lavfi", "-i", "testsrc=size=320x180:rate=25", "-t", "1", "-pix_fmt", "yuv420p", "-f", "mp4", "-movflags", "frag_keyframe+empty_moov", "pipe:1"], { maxBuffer: 1 << 24 });
  assert.equal(frameRate(r.stdout), 25);
  const gif = videoGif(r.stdout, { width: 160, fps: 25 });
  assert.equal(gif.subarray(0, 6).toString(), "GIF89a");
  assert.equal(gif.readUInt16LE(6), 160);
  assert.equal(gif.readUInt16LE(gif.indexOf(Buffer.from([0x21, 0xf9, 0x04])) + 4), 4, "25 fps: 4/100 s a frame");
  assert.ok(gif.includes(Buffer.from("NETSCAPE2.0")), "loops");
});
