// Visual check for render mode: two scores for how alike two slide images are, both computed on
// 160x90 grayscale thumbnails (a transparent PNG counts as on white). Decoded and compared inside a
// Playwright page, so no image library is needed.
//
//   similarity  1 minus the mean absolute difference (1 = identical, 0 = black vs white). It sees
//               a wrong background but not a blank slide: most slides are mostly background.
//   inkOverlap  the overlap of what is drawn on the background. Per image, ink is each pixel's
//               distance from that image's median gray; the score is sum(min) / sum(max) of the two
//               inks (1 = same marks in the same places, 0 = none shared; two blank images are 1).
//               Content missing or shifted by a few pixels drops it sharply.

// Both scores for one pair of PNG buffers: { similarity, ink }.
export async function compareImages(page, a, b) {
  return page.evaluate(
    async ({ a: pngA, b: pngB }) => {
      const load = async (b64) => {
        const img = new Image();
        img.src = `data:image/png;base64,${b64}`;
        await img.decode();
        return img;
      };
      const gray = (img) => {
        const c = document.createElement("canvas");
        c.width = 160;
        c.height = 90;
        const x = c.getContext("2d");
        x.fillStyle = "#ffffff";
        x.fillRect(0, 0, 160, 90);
        x.drawImage(img, 0, 0, 160, 90);
        const d = x.getImageData(0, 0, 160, 90).data;
        const out = new Float64Array(160 * 90);
        for (let i = 0; i < out.length; i++) out[i] = 0.299 * d[i * 4] + 0.587 * d[i * 4 + 1] + 0.114 * d[i * 4 + 2];
        return out;
      };
      const inkOf = (g) => {
        const sorted = Float64Array.from(g).sort();
        const median = sorted[sorted.length >> 1];
        return g.map((v) => Math.abs(v - median));
      };
      const ga = gray(await load(pngA));
      const gb = gray(await load(pngB));
      let diff = 0;
      for (let i = 0; i < ga.length; i++) diff += Math.abs(ga[i] - gb[i]);
      const ia = inkOf(ga);
      const ib = inkOf(gb);
      let lo = 0;
      let hi = 0;
      for (let i = 0; i < ia.length; i++) {
        lo += Math.min(ia[i], ib[i]);
        hi += Math.max(ia[i], ib[i]);
      }
      return { similarity: 1 - diff / ga.length / 255, ink: hi === 0 ? 1 : lo / hi };
    },
    { a: a.toString("base64"), b: b.toString("base64") },
  );
}

export async function similarity(page, a, b) {
  return (await compareImages(page, a, b)).similarity;
}

export async function inkOverlap(page, a, b) {
  return (await compareImages(page, a, b)).ink;
}
