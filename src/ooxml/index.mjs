// writePptx(): SlideSpec[] -> .pptx bytes. This file owns part naming, relationship wiring
// and id allocation; the XML of each part comes from the modules next to it. It knows nothing
// about Slidev or markdown - see ../layout.mjs for that.

import { readFileSync } from "node:fs";
import { extname } from "node:path";
import { createHash } from "node:crypto";
import { Package, REL, CT } from "./package.mjs";
import { esc } from "./xml.mjs";
import { themeXml } from "./theme.mjs";
import { layoutDefs, layoutXml, masterXml } from "./master.mjs";
import { placeholderXml, shapeXml, pictureXml, slideXml, graphicFrameXml, backgroundXml, timingXml } from "./shapes.mjs";
import { presentationXml, presPropsXml, viewPropsXml, appXml, coreXml } from "./presentation.mjs";
import { TABLE_URI, tableXml, tableStylesXml } from "./table.mjs";
import { CHART_URI, chartXml, chartRefXml } from "./chart.mjs";
import { chartWorkbook } from "./xlsx.mjs";
import { notesMasterXml, notesSlideXml } from "./notes.mjs";

const PRES = "/ppt/presentation.xml";
const MASTER = "/ppt/slideMasters/slideMaster1.xml";
const THEME = "/ppt/theme/theme1.xml";
const NOTES_MASTER = "/ppt/notesMasters/notesMaster1.xml";

const VIDEO_EXTS = new Set(["mp4", "mov"]);

// Each distinct image (and video) is stored once, however many slides (or the master) use it.
class MediaStore {
  constructor(pkg) {
    this.pkg = pkg;
    this.parts = new Map();
  }

  // `source` is a file path, or { data: Buffer, ext } for bytes captured in render mode.
  add(source) {
    const fromFile = typeof source === "string";
    const buf = fromFile ? readFileSync(source) : source.data;
    const ext = (fromFile ? extname(source).slice(1) : source.ext).toLowerCase().replace("jpeg", "jpg");
    const key = `${createHash("sha1").update(buf).digest("hex")}.${ext}`;
    if (!this.parts.has(key)) {
      const part = `/ppt/media/${VIDEO_EXTS.has(ext) ? "media" : "image"}${this.parts.size + 1}.${ext}`;
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
      const rId = pkg.rel(part, media.add(shape.path ?? { data: shape.data, ext: shape.ext }), REL.image);
      const svgRId = shape.svg ? pkg.rel(part, media.add({ data: shape.svg, ext: "svg" }), REL.image) : null;
      return pictureXml({ id, name: `Picture ${id}`, descr: shape.alt, rId, svgRId, box: shape, crop: shape.crop, line: shape.line, shadow: shape.shadow });
    }
    case "video": {
      // one file, two relationships: the linked video (a:videoFile) and the embedded media (p14:media)
      const file = media.add({ data: shape.data, ext: shape.ext });
      const video = { linkRId: pkg.rel(part, file, REL.video), embedRId: pkg.rel(part, file, REL.media) };
      const rId = pkg.rel(part, media.add(shape.poster), REL.image);
      env.videos.push({ spid: id, autoplay: shape.autoplay, loop: shape.loop, muted: shape.muted, duration: shape.duration });
      return pictureXml({ id, name: `Video ${id}`, descr: shape.alt, rId, box: shape, crop: shape.crop, video });
    }
    case "table":
      return graphicFrameXml({ id, name: `Table ${id}`, box: shape, uri: TABLE_URI, inner: tableXml(shape, ctx) });
    case "chart": {
      const n = ++env.counters.chart;
      const chartPart = `/ppt/charts/chart${n}.xml`;
      const workbookPart = `/ppt/embeddings/Microsoft_Excel_Worksheet${n}.xlsx`;
      pkg.addMedia(workbookPart, await chartWorkbook(shape.chart));
      const workbookRId = pkg.rel(chartPart, workbookPart, REL.package);
      pkg.addXml(chartPart, CT.chart, chartXml(shape.chart, { lang: ctx.lang, workbookRId }));
      const rId = pkg.rel(part, chartPart, REL.chart);
      return graphicFrameXml({ id, name: `Chart ${id}`, box: shape, uri: CHART_URI, inner: chartRefXml(rId) });
    }
    default:
      throw new Error(`Unknown shape kind ${JSON.stringify(shape.kind)}`);
  }
}

const FONT_STYLES = ["regular", "bold", "italic", "boldItalic"];

// Fonts carried in the file: every font name written so far (theme, master, layouts, slides, notes)
// that `resolve` finds embeddable is renamed to the family it found (a web font's name becomes the
// installed family's, the name inside the embedded font) and its styles are added as
// /ppt/fonts/fontN.fntdata. Theme references (+mj-lt ...) and names it can't find are left as they are.
function embedFonts(pkg, resolve, notice = (msg) => console.warn(msg)) {
  const names = new Set();
  for (const [, part] of pkg.parts) {
    if (typeof part.data !== "string") continue;
    for (const m of part.data.matchAll(/typeface="([^"+][^"]*)"/g)) names.add(m[1].replace(/&amp;/g, "&"));
  }
  const families = new Map(); // family -> rIds per style
  const rename = new Map();
  let n = 0;
  for (const name of [...names].sort()) {
    const found = resolve(name);
    if (!found) {
      notice(`Notice: font "${name}" not embedded: no installed font of that family.`);
      continue;
    }
    for (const note of found.notes) notice(`Notice: font "${name}": ${note}`);
    if (!found.fonts.regular) continue;
    if (found.family !== name) rename.set(name, found.family);
    if (families.has(found.family)) continue;
    const rIds = {};
    for (const style of FONT_STYLES) {
      if (!found.fonts[style]) continue;
      const part = `/ppt/fonts/font${++n}.fntdata`;
      pkg.addMedia(part, found.fonts[style]);
      rIds[style] = pkg.rel(PRES, part, REL.font);
    }
    families.set(found.family, rIds);
  }
  if (rename.size) {
    for (const [, part] of pkg.parts) {
      if (typeof part.data !== "string") continue;
      part.data = part.data.replace(/typeface="([^"]*)"/g, (m, f) => {
        const to = rename.get(f.replace(/&amp;/g, "&"));
        return to ? `typeface="${esc(to)}"` : m;
      });
    }
  }
  return [...families].map(([typeface, rIds]) => ({ typeface, ...rIds }));
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
  // Render mode can put elements every slide shares (a footer logo) on the master. Their ids
  // follow the master's own title (2), body (3) and logo (4).
  const masterEnv = { pkg, part: MASTER, media, ctx, counters, videos: [] };
  let extra = "";
  let masterShapeId = 5;
  for (const shape of deck.masterShapes || []) extra += await shapeToXml(shape, masterShapeId++, masterEnv);
  pkg.addXml(MASTER, CT.slideMaster, masterXml({ geometry: deck.geometry, ctx, layoutIds, logo, extra }));

  // The notes master (with its own theme - every master needs one) only exists if some slide
  // has notes.
  const hasNotes = slides.some((s) => s.notes);
  let notesMasterRId = null;
  if (hasNotes) {
    const NOTES_THEME = "/ppt/theme/theme2.xml";
    pkg.addXml(NOTES_THEME, CT.theme, themeXml({ name: "Notes", colors: deck.theme, fonts: deck.fonts }));
    pkg.addXml(NOTES_MASTER, CT.notesMaster, notesMasterXml(ctx));
    pkg.rel(NOTES_MASTER, NOTES_THEME, REL.theme);
    notesMasterRId = pkg.rel(PRES, NOTES_MASTER, REL.notesMaster);
  }
  let notesCount = 0;

  const slideRefs = [];
  for (const [i, spec] of slides.entries()) {
    const part = `/ppt/slides/slide${i + 1}.xml`;
    const layout = layouts.get(spec.layout);
    if (!layout) throw new Error(`Slide ${i + 1}: unknown layout ${JSON.stringify(spec.layout)}`);
    pkg.rel(part, layout.part, REL.slideLayout);
    const sctx = { ...ctx, link: (url) => pkg.rel(part, url, REL.hyperlink, { external: true }) };
    const env = { pkg, part, media, ctx: sctx, counters, videos: [] };

    let nextId = 2;
    let xml = "";
    // Only placeholders that have content are put on the slide; empty ones would show
    // "Click to add text" prompts in the editor.
    const ph = spec.placeholders || {};
    // placeholders.titleZ (render mode) is how many shapes are drawn beneath the title, in DOM
    // order. Without it the title comes first, as every placeholder does.
    let heldTitle = null;
    for (const p of layout.def.placeholders) {
      const paras = ph[p.role];
      if (!paras || !paras.length) continue;
      const box = p.role === "title" ? ph.titleBox : p.role === "body" ? ph.bodyBox : undefined;
      const body = p.role === "body" && ph.bodyFontScale ? { fontScale: ph.bodyFontScale } : p.role === "title" && ph.titleBody ? ph.titleBody : {};
      const phXml = placeholderXml({ id: nextId++, name: p.name, ph: p.ph, box, paras, ctx: sctx, body });
      if (p.role === "title" && ph.titleZ) heldTitle = phXml;
      else xml += phXml;
    }
    const shapes = spec.shapes || [];
    for (const [i, shape] of shapes.entries()) {
      if (heldTitle !== null && i === ph.titleZ) {
        xml += heldTitle;
        heldTitle = null;
      }
      xml += await shapeToXml(shape, nextId++, env);
    }
    if (heldTitle !== null) xml += heldTitle;

    let bg = "";
    if (spec.background?.image) bg = backgroundXml(spec.background, pkg.rel(part, media.add(spec.background.image), REL.image));
    else if (spec.background?.color || spec.background?.gradient) bg = backgroundXml(spec.background);
    pkg.addXml(part, CT.slide, slideXml(xml, bg, timingXml(env.videos)));
    if (spec.notes) {
      const notesPart = `/ppt/notesSlides/notesSlide${i + 1}.xml`;
      pkg.addXml(notesPart, CT.notesSlide, notesSlideXml(spec.notes, ctx));
      pkg.rel(notesPart, NOTES_MASTER, REL.notesMaster);
      pkg.rel(notesPart, part, REL.slide);
      pkg.rel(part, notesPart, REL.notesSlide);
      notesCount++;
    }
    slideRefs.push({ id: 256 + i, rId: pkg.rel(PRES, part, REL.slide) });
  }

  const embedded = deck.fontResolver ? embedFonts(pkg, deck.fontResolver, deck.onFontNotice) : [];

  pkg.rel(PRES, THEME, REL.theme);
  for (const [name, type, rel, xml] of [
    ["/ppt/presProps.xml", CT.presProps, REL.presProps, presPropsXml()],
    ["/ppt/viewProps.xml", CT.viewProps, REL.viewProps, viewPropsXml()],
    ["/ppt/tableStyles.xml", CT.tableStyles, REL.tableStyles, tableStylesXml()],
  ]) {
    pkg.addXml(name, type, xml);
    pkg.rel(PRES, name, rel);
  }
  pkg.addXml(PRES, CT.presentation, presentationXml({
    masterId, masterRId, notesMasterRId, slides: slideRefs, embeddedFonts: embedded,
    width: deck.geometry.width, height: deck.geometry.height,
  }));
  pkg.addXml("/docProps/core.xml", CT.coreProps, coreXml({ title: deck.title, author: deck.author, date: deck.timestamps ? new Date() : null }));
  pkg.addXml("/docProps/app.xml", CT.extProps, appXml({ slides: slides.length, notes: notesCount }));
  return pkg.toBuffer();
}
