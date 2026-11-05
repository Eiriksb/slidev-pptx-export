// Render mode's capture for an animated export: every click state of every slide, directly loaded as
// /<n>?clicks=k, and a clip for each click whose page moves by script.

import { readFileSync, rmSync } from "node:fs";
import { join } from "node:path";
import { clickKind, pictureChanged } from "./animation-plan.mjs";
import { captureSlide, openDeck } from "./capture.mjs";
import { compareImages } from "./compare.mjs";
import { ENCODER_HINT, encodeClip, findEncoder } from "./encode.mjs";
import { clipFrame, recordClick } from "./record.mjs";

// Below this, a clip's last frame isn't what a fresh load of that state shows.
const SAME_PICTURE = 0.95;
const RECORD_TIMEOUT_MS = 30000;

const firstLine = (err) => String(err?.message || err).split("\n")[0];
const frameFile = (dir, i) => join(dir, `f${String(i).padStart(5, "0")}.jpg`);

// `nav.clicksTotal` is a number or, depending on how Slidev exposes it, a ref.
const clicksTotal = (page) =>
  page.evaluate(() => {
    const total = window.__slidev__?.nav?.clicksTotal;
    return Number(total && typeof total === "object" ? total.value : total) || 0;
  });

async function statesOf(page, baseUrl, s, canvas, timeout) {
  const at = (k) => ({ url: `${baseUrl}/${s.no}?clicks=${k}`, jpeg: true });
  const states = [await captureSlide(page, baseUrl, s, canvas, true, timeout, at(0))];
  const total = await clicksTotal(page);
  for (let k = 1; k <= total; k++) states.push(await captureSlide(page, baseUrl, s, canvas, true, timeout, at(k)));
  return states;
}

// A click that changed a picture: records it to see whether the page moves by script, and if it does,
// makes the clip. Anything that goes wrong costs the clip, not the export: the click is then a Fade.
async function recordedClick({ page, baseUrl, s, k, canvas, state, record, encoder, encode, notice }) {
  const where = `slide ${s.no} click ${k}`;
  const reveal = { kind: "reveal" };
  let rec;
  try {
    rec = await record({ browser: page.context().browser(), url: `${baseUrl}/${s.no}?clicks=${k - 1}`, sel: `.slidev-page-${s.no}`, canvas, timeout: RECORD_TIMEOUT_MS });
  } catch (err) {
    notice(`Notice: ${where}: couldn't record it (${firstLine(err)}) - exported as a Fade.`);
    return reveal;
  }
  try {
    const { kind, reason } = clickKind({ changed: true, record: rec, encoder });
    if (kind === "reveal") {
      if (reason) notice(`Notice: ${where}: ${reason}${encoder ? "" : ` (${ENCODER_HINT})`} - exported as a Fade.`);
      return reveal;
    }
    const data = encode(encoder, rec);
    const { similarity } = await compareImages(page, readFileSync(frameFile(rec.dir, rec.count - 1)), state.screenshot);
    if (similarity < SAME_PICTURE) notice(`Notice: ${where}: the clip's last frame differs from a fresh load of the slide (the page's live state differs from a fresh load, or a recording is wrong).`);
    const duration = rec.count / rec.fps;
    notice(`Notice: ${where}: clip, ${duration.toFixed(1)} s, ${(data.length / 2 ** 20).toFixed(1)} MB.`);
    const { width, height } = clipFrame(canvas);
    return { kind: "motion", clip: { data, poster: readFileSync(frameFile(rec.dir, 0)), duration, natural: { w: width, h: height } } };
  } catch (err) {
    notice(`Notice: ${where}: couldn't make a clip (${firstLine(err)}) - exported as a Fade.`);
    return reveal;
  } finally {
    rmSync(rec.dir, { recursive: true, force: true });
  }
}

async function captureOne({ page, baseUrl, s, canvas, timeout, record, encoder, encode, notice }) {
  const source = { index: s.no, sourceIndex: s.sourceIndex, filepath: s.filepath };
  const slide = { ...source, notes: s.notes, transition: s.transition };
  let states;
  try {
    states = await statesOf(page, baseUrl, s, canvas, timeout);
  } catch (err) {
    try {
      const final = await captureSlide(page, baseUrl, s, canvas, false, timeout, { jpeg: true });
      notice(`Notice: slide ${s.no}: its click states couldn't be captured (${firstLine(err)}) - exported in its final state.`);
      return { ...slide, states: [final], clicks: [] };
    } catch (failed) {
      return { ...source, error: firstLine(failed) };
    }
  }
  const clicks = states.length > 1 ? [null] : [];
  for (let k = 1; k < states.length; k++) {
    const changed = pictureChanged(states[k - 1].nodes, states[k].nodes);
    clicks.push(changed ? await recordedClick({ page, baseUrl, s, k, canvas, state: states[k], record, encoder, encode, notice }) : { kind: "reveal" });
  }
  for (const state of states) delete state.screenshot;
  return { ...slide, states, clicks };
}

// { canvas, slides }. A slide is { index, sourceIndex, filepath, notes, transition, states, clicks }, with
// clicks[k] for k >= 1 being { kind: "reveal" } or { kind: "motion", clip: { data, poster, duration, natural } };
// or, when it couldn't be captured, { index, sourceIndex, filepath, error }.
export async function captureClickStates({ baseUrl, browser, timeout = 15000, slidevVersion = "unknown", record = recordClick, encoder = findEncoder(), encode = encodeClip, notice = console.warn }) {
  const page = await browser.newPage({ viewport: { width: 980, height: 552 }, deviceScaleFactor: 2 });
  try {
    page.setDefaultTimeout(timeout);
    const { canvas, list } = await openDeck(page, baseUrl, timeout, slidevVersion);
    const slides = [];
    for (const s of list) slides.push(await captureOne({ page, baseUrl, s, canvas, timeout, record, encoder, encode, notice }));
    return { canvas, slides };
  } finally {
    await page.close();
  }
}
