// Records what a click does on the page, frame by frame on a stepped clock (clock.js). Real-time capture
// isn't good enough: Chromium's screencast managed about 9.5 fps on a map flight, with gaps of half a
// second. The frames show the whole slide, as a clip slide holds nothing else.

import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { open, waitForQuietSlide } from "./capture.mjs";

const CLOCK = readFileSync(new URL("./clock.js", import.meta.url), "utf8");

const CLIP_WIDTH = 1920;
// How long the page must stay unchanged, on its own clock, to count as finished.
const QUIET_MS = 600;
const IMAGE_POLL_MS = 60;

// Scripts animating for this many frames after a click (0.4 s at 30 fps) make the click motion.
export const MOTION_FRAMES = 12;

// How tall the slide is at the clip's width. Rounded first, so a slide that fits exactly isn't pushed
// up to the next even number by floating point.
const slideHeight = ({ w, h }) => Math.round(((CLIP_WIDTH * h) / w) * 100) / 100;

// The window to record in: the slide at the clip's width, with even dimensions because H.264 requires
// them.
export function clipViewport(canvas) {
  return { width: CLIP_WIDTH, height: 2 * Math.ceil(slideHeight(canvas) / 2) };
}

// The part of that window a frame shows. Slidev scales the slide to the width and centres it on black,
// so a slide of 1081.47 in a window of 1082 would leave a black row at the top and bottom of every frame.
export function clipFrame(canvas) {
  const height = 2 * Math.floor(slideHeight(canvas) / 2);
  return { x: 0, y: (clipViewport(canvas).height - height) / 2, width: CLIP_WIDTH, height };
}

const frameName = (i) => `f${String(i).padStart(5, "0")}.jpg`;

// Opens `url` (the state before the click), clicks, and steps until the page settles or `maxSeconds`
// runs out. `dir` holds the frames, frame 0 being the state before the click; the caller removes it.
// `setup` (optional) runs on the new browser context.
export async function recordClick({ browser, url, sel, canvas, tmpRoot = tmpdir(), fps = 30, maxSeconds = 40, timeout = 30000, imageGrace = 3000, setup = async () => {} }) {
  const context = await browser.newContext({ viewport: clipViewport(canvas), deviceScaleFactor: 1 });
  let dir = null;
  try {
    await setup(context);
    const page = await context.newPage();
    page.setDefaultTimeout(timeout);
    await page.addInitScript(CLOCK);
    await open(page, url);
    await page.waitForFunction((s) => document.querySelector(s)?.getBoundingClientRect().width > 0, sel);
    await page.evaluate(() => document.fonts.ready.then(() => true));
    await waitForQuietSlide(page, sel, timeout);

    dir = mkdtempSync(join(tmpRoot, "pptx-record-"));
    const cdp = await context.newCDPSession(page);
    const clip = clipFrame(canvas);
    let count = 0;
    let last = null;
    const frame = async () => {
      const { data } = await cdp.send("Page.captureScreenshot", { format: "jpeg", quality: 92, optimizeForSpeed: true, clip: { ...clip, scale: 1 } });
      last = Buffer.from(data, "base64");
      writeFileSync(join(dir, frameName(count++)), last);
      return last;
    };

    await page.evaluate((s) => window.__startVirtualClock(s), sel);
    const first = await frame();
    // Not awaited in the page: Slidev's next() may wait on timers, which now run only when stepped.
    await page.evaluate(() => {
      window.__slidev__.nav.next();
    });

    const dt = 1000 / fps;
    let motionFrames = 0;
    let settled = false;
    let imagesWaited = 0;
    for (let i = 0; i < maxSeconds * fps; i++) {
      const state = await page.evaluate((ms) => window.__stepClock(ms), dt);
      if (state.images) {
        await page.waitForTimeout(IMAGE_POLL_MS);
        imagesWaited += IMAGE_POLL_MS;
        if (imagesWaited >= imageGrace) await page.evaluate(() => window.__giveUpOnImages());
      } else {
        imagesWaited = 0;
      }
      await frame();
      if (state.frames > 0 || state.intervals > 0) motionFrames++;
      if (state.frames === 0 && state.running === 0 && state.images === 0 && state.sinceMutation > QUIET_MS) {
        settled = true;
        break;
      }
    }
    return { dir, count, fps, motionFrames, settled, changedVisibly: !first.equals(last) };
  } catch (err) {
    if (dir) rmSync(dir, { recursive: true, force: true });
    throw err;
  } finally {
    await context.close();
  }
}
