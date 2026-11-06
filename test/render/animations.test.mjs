// The whole feature through its front doors, on a small deck with its own Slidev:
// npm --prefix test/fixtures/animated-deck ci
import { test, after } from "node:test";
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { existsSync, mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { promisify } from "node:util";
import { ROOT, usableChromium } from "./helpers.mjs";
import { convert } from "../../src/index.mjs";
import { findEncoder } from "../../src/render/encode.mjs";
import { lintPackage } from "../lint.mjs";
import { unzip } from "../helpers.mjs";

const FIXTURE = join(ROOT, "test/fixtures/animated-deck");
const SLIDES = join(FIXTURE, "slides.md");
const ready = existsSync(join(FIXTURE, "node_modules/@slidev/cli"));
const chromium = ready ? await usableChromium() : null;
const opts = { skip: !ready ? "run npm --prefix test/fixtures/animated-deck ci" : !chromium ? "Playwright/Chromium not available" : false, timeout: 300000 };
const hasEncoder = !!findEncoder();
const needsEncoder = { ...opts, skip: opts.skip || (hasEncoder ? false : "no encoder installed") };

const work = mkdtempSync(join(tmpdir(), "pptx-animations-test-"));
after(() => rmSync(work, { recursive: true, force: true }));

// Each way of exporting runs once: an export takes the best part of a minute.
const WAYS = {
  cli: async (out) => {
    const { stderr } = await promisify(execFile)(process.execPath, [join(ROOT, "bin/cli.mjs"), SLIDES, out, "--animations"], { cwd: work, maxBuffer: 2 ** 24 });
    return stderr;
  },
  config: (out) => convert({ inputPath: SLIDES, outputPath: out, config: { render: { animations: true } } }).then(() => ""),
  off: (out) => convert({ inputPath: SLIDES, outputPath: out, render: true }).then(() => ""),
};
const done = new Map();
function exportFixture(way = "cli") {
  if (!done.has(way)) {
    done.set(
      way,
      (async () => {
        const out = join(work, `${way}.pptx`);
        const notices = await WAYS[way](out);
        const buffer = readFileSync(out);
        return { buffer, z: await unzip(buffer), notices };
      })(),
    );
  }
  return done.get(way);
}

const slideTexts = (z) =>
  z.names
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]))
    .map((n) => z.text(n));
const count = (xml, re) => (xml.match(re) || []).length;
const clickSteps = (xml) => count(xml, /nodeType="clickEffect"/g);
const entrances = (xml) => count(xml, /presetClass="entr"/g);
const exits = (xml) => count(xml, /presetClass="exit"/g);
const tables = (xml) => count(xml, /<a:tbl>/g);

test("the deck exports in Slidev's click order", opts, async () => {
  const { buffer, z } = await exportFixture();
  assert.deepEqual(await lintPackage(buffer), []);
  const slides = slideTexts(z);
  assert.equal(slides.length, hasEncoder ? 4 : 3);
  assert.deepEqual(slides.slice(0, 3).map(clickSteps), [2, 2, hasEncoder ? 0 : 1]);
  for (const s of slides) assert.doesNotMatch(s, /3DA1BC7CA4C3/);
});

test("slide 1's v-clicks each enter on their own click, as native text, and it fades in", opts, async () => {
  const { z } = await exportFixture();
  const [one] = slideTexts(z);
  assert.equal(entrances(one), 2);
  assert.equal(exits(one), 0);
  assert.match(one, /First reveal/);
  assert.match(one, /Second reveal/);
  assert.match(one, /<\/p:clrMapOvr><p:transition><p:fade\/><\/p:transition>/);
});

test("slide 2's table steps one column at a time, one native table per state", opts, async () => {
  const { z } = await exportFixture();
  const two = slideTexts(z)[1];
  assert.equal(tables(two), 3);
  assert.equal(entrances(two), 4);
  assert.equal(exits(two), 4);
  assert.match(two, /Step 2/);
});

test("the motion click is a slide holding only a clip that starts by itself, with no click step and no transition", needsEncoder, async () => {
  const { z } = await exportFixture();
  const clip = slideTexts(z)[3];
  assert.match(clip, /nodeType="afterEffect"/);
  assert.match(clip, /<p:cMediaNode mute="1">/);
  assert.doesNotMatch(clip, /clickEffect|<p:transition/);
  assert.equal(count(clip, /<p:(sp|pic|graphicFrame)>/g), 1);
  const mp4 = z.names.find((n) => n.endsWith(".mp4"));
  assert.equal(z.buffer(mp4).subarray(4, 8).toString(), "ftyp");
});

test("the slide before a clip is native, and what the export did is on stderr", needsEncoder, async () => {
  const { z, notices } = await exportFixture();
  assert.match(slideTexts(z)[2], /Script motion/);
  assert.match(notices, /^Notice: slide 3 click 1: clip, [\d.]+ s, [\d.]+ MB\.$/m);
  assert.doesNotMatch(notices, /last frame differs/, "a clip that ends as the page loads gets no notice");
});

test("animations in the config imply render mode too", opts, async () => {
  const { z } = await exportFixture("config");
  assert.equal(slideTexts(z).length, hasEncoder ? 4 : 3);
});

test("without the flag the same deck has one slide each and no effects", opts, async () => {
  const { z } = await exportFixture("off");
  const slides = slideTexts(z);
  assert.equal(slides.length, 3);
  for (const s of slides) assert.doesNotMatch(s, /<p:timing|<p:transition/);
  assert.equal(z.names.filter((n) => n.endsWith(".mp4")).length, 0);
});

// --progress json: stdout is the progress events and nothing else, one JSON object a line.
const jsonLines = (stdout) => stdout.split("\n").filter(Boolean).map((l) => JSON.parse(l));
const cli = (args) => promisify(execFile)(process.execPath, [join(ROOT, "bin/cli.mjs"), ...args], { cwd: work, maxBuffer: 2 ** 24 });

test("--progress json writes only JSON lines to stdout: start first, done last", opts, async () => {
  const out = join(work, "progress.pptx");
  const { stdout, stderr } = await cli([SLIDES, out, "--animations", "--progress", "json"]);
  const events = jsonLines(stdout);
  assert.deepEqual(events[0], { event: "start", slides: 3 });
  const last = events.at(-1);
  assert.deepEqual(Object.keys(last), ["event", "path", "slides", "clips", "bytes"]);
  assert.equal(last.event, "done");
  assert.equal(last.path, out);
  assert.equal(last.slides, hasEncoder ? 4 : 3);
  assert.equal(last.clips, hasEncoder ? 1 : 0);
  assert.equal(last.bytes, readFileSync(out).length);
  assert.equal(events.at(-2).event, "write");
  assert.ok(events.some((e) => e.event === "slide" && e.states > 1));
  assert.ok(events.some((e) => e.event === "state"));
  // What stderr says is also a notice event, and "Wrote N slides" is not on stdout.
  const notices = events.filter((e) => e.event === "notice").map((e) => e.text);
  assert.ok(notices.length > 0);
  for (const text of notices) assert.ok(stderr.includes(text));
  assert.doesNotMatch(stdout, /Wrote /);
});

test("--progress json: a deck that doesn't exist ends with an error line and exit code 1", async () => {
  await assert.rejects(cli([join(work, "missing.md"), join(work, "x.pptx"), "--progress", "json"]), (err) => {
    assert.equal(err.code, 1);
    const events = jsonLines(err.stdout);
    assert.equal(events.length, 1);
    assert.equal(events[0].event, "error");
    assert.match(events[0].message, /not found.*missing\.md/);
    return true;
  });
});
