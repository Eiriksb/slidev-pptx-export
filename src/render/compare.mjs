// Visual check for render mode: how alike two slide images are, as 1 minus the mean absolute
// difference of their 160x90 grayscale thumbnails (1 = identical, 0 = black vs white). Decoded
// and compared inside a Playwright page, so no image library is needed.

export async function similarity(page, a, b) {
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
      const ga = gray(await load(pngA));
      const gb = gray(await load(pngB));
      let sum = 0;
      for (let i = 0; i < ga.length; i++) sum += Math.abs(ga[i] - gb[i]);
      return 1 - sum / ga.length / 255;
    },
    { a: a.toString("base64"), b: b.toString("base64") },
  );
}
