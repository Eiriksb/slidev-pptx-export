// writePptx(): SlideSpec[] -> .pptx bytes. This file owns part naming, relationship wiring
// and id allocation; the XML of each part comes from the modules next to it. It knows nothing
// about Slidev or markdown - see ../layout.mjs for that.

import { readFileSync } from "node:fs";
import { extname } from "node:path";
import { createHash } from "node:crypto";
import { Package, REL, CT } from "./package.mjs";
import { themeXml } from "./theme.mjs";
import { layoutDefs, layoutXml, masterXml } from "./master.mjs";
import { placeholderXml, shapeXml, pictureXml, slideXml } from "./shapes.mjs";
import { presentationXml, presPropsXml, viewPropsXml, EMPTY_TABLE_STYLES, appXml, coreXml } from "./presentation.mjs";

const PRES = "/ppt/presentation.xml";
const MASTER = "/ppt/slideMasters/slideMaster1.xml";
const THEME = "/ppt/theme/theme1.xml";

// Each distinct image is stored once, however many slides (or the master) use it.
class MediaStore {
  constructor(pkg) {
    this.pkg = pkg;
    this.parts = new Map();
  }

  add(path) {
    const buf = readFileSync(path);
    const ext = extname(path).slice(1).toLowerCase().replace("jpeg", "jpg");
    const key = `${createHash("sha1").update(buf).digest("hex")}.${ext}`;
    if (!this.parts.has(key)) {
      const part = `/ppt/media/image${this.parts.size + 1}.${ext}`;
      this.pkg.addMedia(part, buf);
      this.parts.set(key, part);
    }
    return this.parts.get(key);
  }
}

async function shapeToXml(shape, id, env) {
  const { pkg, part, media, ctx } = env;
  switch (shape.kind) {
    case "text":
      return shapeXml(shape, id, ctx);
    case "image": {
      const rId = pkg.rel(part, media.add(shape.path), REL.image);
      return pictureXml({ id, name: `Picture ${id}`, descr: shape.alt, rId, box: shape });
    }
    default:
      throw new Error(`Unknown shape kind ${JSON.stringify(shape.kind)}`);
  }
}

export async function writePptx(slides, deck) {
  const pkg = new Package();
  const media = new MediaStore(pkg);
  const ctx = { lang: deck.lang || "en-US", fonts: deck.fonts };
  const counters = { chart: 0 };
  // Slide masters and layouts share one id space that starts at 2^31.
  let nextMasterId = 2147483648;

  pkg.rel("", PRES, REL.officeDocument);
  pkg.rel("", "/docProps/core.xml", REL.coreProps);
  pkg.rel("", "/docProps/app.xml", REL.extProps);

  pkg.addXml(THEME, CT.theme, themeXml({ name: deck.title || "Deck", colors: deck.theme, fonts: deck.fonts }));
  const masterRId = pkg.rel(PRES, MASTER, REL.slideMaster);
  const masterId = nextMasterId++;
  pkg.rel(MASTER, THEME, REL.theme);
  const layouts = new Map();
  const layoutIds = [];
  layoutDefs(deck.geometry).forEach((def, i) => {
    const part = `/ppt/slideLayouts/slideLayout${i + 1}.xml`;
    pkg.addXml(part, CT.slideLayout, layoutXml(def, ctx));
    pkg.rel(part, MASTER, REL.slideMaster);
    layoutIds.push({ id: nextMasterId++, rId: pkg.rel(MASTER, part, REL.slideLayout) });
    layouts.set(def.key, { part, def });
  });
  const logo = deck.logo ? { rId: pkg.rel(MASTER, media.add(deck.logo.path), REL.image), box: deck.logo.box } : null;
  pkg.addXml(MASTER, CT.slideMaster, masterXml({ geometry: deck.geometry, ctx, layoutIds, logo }));

  const slideRefs = [];
  for (const [i, spec] of slides.entries()) {
    const part = `/ppt/slides/slide${i + 1}.xml`;
    const layout = layouts.get(spec.layout);
    if (!layout) throw new Error(`Slide ${i + 1}: unknown layout ${JSON.stringify(spec.layout)}`);
    pkg.rel(part, layout.part, REL.slideLayout);
    const sctx = { ...ctx, link: (url) => pkg.rel(part, url, REL.hyperlink, { external: true }) };
    const env = { pkg, part, media, ctx: sctx, counters };

    let nextId = 2;
    let xml = "";
    // Only placeholders that have content are put on the slide; empty ones would show
    // "Click to add text" prompts in the editor.
    const ph = spec.placeholders || {};
    for (const p of layout.def.placeholders) {
      const paras = ph[p.role];
      if (!paras || !paras.length) continue;
      const box = p.role === "title" ? ph.titleBox : p.role === "body" ? ph.bodyBox : undefined;
      const body = p.role === "body" && ph.bodyFontScale ? { fontScale: ph.bodyFontScale } : {};
      xml += placeholderXml({ id: nextId++, name: p.name, ph: p.ph, box, paras, ctx: sctx, body });
    }
    for (const shape of spec.shapes || []) xml += await shapeToXml(shape, nextId++, env);

    pkg.addXml(part, CT.slide, slideXml(xml));
    slideRefs.push({ id: 256 + i, rId: pkg.rel(PRES, part, REL.slide) });
  }

  pkg.rel(PRES, THEME, REL.theme);
  for (const [name, type, rel, xml] of [
    ["/ppt/presProps.xml", CT.presProps, REL.presProps, presPropsXml()],
    ["/ppt/viewProps.xml", CT.viewProps, REL.viewProps, viewPropsXml()],
    ["/ppt/tableStyles.xml", CT.tableStyles, REL.tableStyles, EMPTY_TABLE_STYLES],
  ]) {
    pkg.addXml(name, type, xml);
    pkg.rel(PRES, name, rel);
  }
  pkg.addXml(PRES, CT.presentation, presentationXml({
    masterId, masterRId, notesMasterRId: null, slides: slideRefs,
    width: deck.geometry.width, height: deck.geometry.height,
  }));
  pkg.addXml("/docProps/core.xml", CT.coreProps, coreXml({ title: deck.title, author: deck.author, date: deck.timestamps ? new Date() : null }));
  pkg.addXml("/docProps/app.xml", CT.extProps, appXml({ slides: slides.length, notes: 0 }));
  return pkg.toBuffer();
}
