// A clip must play in the colours of the pictures around it, or the slide visibly jumps where a clip
// starts and ends. Playwright's own Chromium can't decode H.264, so this needs Google Chrome.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { encodeClip, findEncoder } from "../../src/render/encode.mjs";
import { ROOT } from "./helpers.mjs";
import { loadChromium } from "../../src/render/browser.mjs";

async function chrome() {
  try {
    return await (await loadChromium(ROOT)).launch({ channel: "chrome" });
  } catch {
    return null;
  }
}

const encoder = findEncoder();
const browser = encoder ? await chrome() : null;
const opts = { skip: !encoder ? "no encoder installed" : !browser ? "needs Google Chrome: Playwright's Chromium can't decode H.264" : false };

test.after(() => browser?.close());

// Each channel's mean over a whole picture or a video's first frame, drawn to a canvas.
const means = (page, src, isVideo) =>
  page.evaluate(
    async ({ src, isVideo }) => {
      let el;
      if (isVideo) {
        el = document.createElement("video");
        el.muted = true;
        el.src = src;
        await new Promise((ok, fail) => {
          el.onloadeddata = ok;
          el.onerror = () => fail(new Error("the video didn't decode"));
        });
        // A frame can only be drawn once it is decoded, which loadeddata doesn't promise.
        const shown = new Promise((ok) => {
          el.requestVideoFrameCallback(ok);
          setTimeout(ok, 1000);
        });
        el.currentTime = 0.02;
        await new Promise((ok) => (el.onseeked = ok));
        await shown;
      } else {
        el = new Image();
        el.src = src;
        await el.decode();
      }
      const c = document.createElement("canvas");
      c.width = el.videoWidth || el.width;
      c.height = el.videoHeight || el.height;
      const x = c.getContext("2d", { willReadFrequently: true });
      x.drawImage(el, 0, 0);
      const d = x.getImageData(0, 0, c.width, c.height).data;
      const sum = [0, 0, 0];
      for (let i = 0; i < d.length; i += 4) for (let k = 0; k < 3; k++) sum[k] += d[i + k];
      return sum.map((v) => v / (d.length / 4));
    },
    { src, isVideo },
  );

test("a clip plays in the colours of its frames", opts, async (t) => {
  const dir = mkdtempSync(join(tmpdir(), "pptx-colour-test-"));
  t.after(() => rmSync(dir, { recursive: true, force: true }));
  const page = await browser.newPage({ viewport: { width: 640, height: 360 } });
  t.after(() => page.close());
  await page.setContent(
    '<!doctype html><body style="margin:0;background:#f2f4ee">' +
      '<div style="position:absolute;left:40px;top:40px;width:200px;height:120px;background:#3b6ea5"></div>' +
      '<div style="position:absolute;left:300px;top:60px;width:280px;height:220px;background:linear-gradient(135deg,#d7e8c8,#7fa86a)"></div>' +
      '<div style="position:absolute;left:0;top:320px;width:640px;height:40px;background:#6e2a4a"></div>',
  );
  const frame = await page.screenshot({ type: "jpeg", quality: 92 });
  for (let i = 0; i < 6; i++) writeFileSync(join(dir, `f${String(i).padStart(5, "0")}.jpg`), frame);
  const clip = encodeClip(encoder, { dir, count: 6, fps: 30 });
  const source = await means(page, `data:image/jpeg;base64,${frame.toString("base64")}`, false);
  const played = await means(page, `data:video/mp4;base64,${clip.toString("base64")}`, true);
  for (let k = 0; k < 3; k++) assert.ok(Math.abs(source[k] - played[k]) <= 3, `channel ${k}: frames ${source[k].toFixed(1)}, clip ${played[k].toFixed(1)}`);
});
