// PresentationML shape builders: placeholders, free text/geometry shapes, pictures and
// graphic frames (tables, charts). Geometry arrives in inches; ids come from the caller's
// per-part counter so they stay unique within a slide.

import { XML_DECL, NS_PML, esc, emu, emuSize, solidFill, colorXml, fillXml, gradFillXml } from "./xml.mjs";
import { txBodyXml } from "./text.mjs";

export const SP_TREE_HEADER =
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

export function xfrmXml(box, tag = "a:xfrm") {
  // rot is in 60000ths of a degree, 0..21599999.
  const rot = box.rot ? ` rot="${Math.round((((box.rot % 360) + 360) % 360) * 60000) % 21600000}"` : "";
  return `<${tag}${rot}><a:off x="${emu(box.x)}" y="${emu(box.y)}"/><a:ext cx="${emuSize(box.w)}" cy="${emuSize(box.h)}"/></${tag}>`;
}

export function phXml(ph) {
  const type = ph.type ? ` type="${ph.type}"` : "";
  const idx = ph.idx != null ? ` idx="${ph.idx}"` : "";
  const sz = ph.sz ? ` sz="${ph.sz}"` : "";
  return `<p:ph${type}${idx}${sz}/>`;
}

export function placeholderXml({ id, name, ph, box, paras, ctx, body = {} }) {
  // No xfrm means "inherit position and size from the layout" - that's what makes Reset Slide
  // and layout switching work.
  const spPr = box ? `<p:spPr>${xfrmXml(box)}</p:spPr>` : "<p:spPr/>";
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
    `<p:nvPr>${phXml(ph)}</p:nvPr></p:nvSpPr>${spPr}${txBodyXml(paras, ctx, { ...body, placeholder: true })}</p:sp>`
  );
}

function geometryXml(shape) {
  if (shape.geometry === "roundRect") {
    // roundRect's adj is the corner radius as a fraction (x100000) of the shorter side.
    const side = Math.max(0.0001, Math.min(shape.w, shape.h));
    const adj = Math.round(Math.min(50000, Math.max(0, ((shape.radius ?? 0.1) / side) * 100000)));
    return `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val ${adj}"/></a:avLst></a:prstGeom>`;
  }
  return `<a:prstGeom prst="${shape.geometry === "ellipse" ? "ellipse" : "rect"}"><a:avLst/></a:prstGeom>`;
}

function lineXml(line) {
  // Explicit "no line": without a p:style there is no line anyway in PowerPoint, but Keynote
  // and Google Slides apply their own defaults to a shape with no a:ln at all.
  if (!line) return "<a:ln><a:noFill/></a:ln>";
  const dash = line.dash ? `<a:prstDash val="${line.dash}"/>` : "";
  return `<a:ln w="${Math.round(line.width * 12700)}">${solidFill(line.color, line.alpha)}${dash}</a:ln>`;
}

const SHADOW =
  '<a:effectLst><a:outerShdw blurRad="114300" dist="38100" dir="5400000" algn="t" rotWithShape="0">' +
  `${colorXml("dk2", 0.16)}</a:outerShdw></a:effectLst>`;

// `true` = the deck's soft card shadow; an object = a measured CSS box-shadow.
export function shadowXml(shadow) {
  if (shadow === true) return SHADOW;
  const dir = Math.round(((((shadow.dir ?? 90) % 360) + 360) % 360) * 60000) % 21600000;
  return (
    `<a:effectLst><a:outerShdw blurRad="${emuSize(shadow.blur ?? 0)}" dist="${emuSize(shadow.dist ?? 0)}" dir="${dir}" algn="ctr" rotWithShape="0">` +
    `${colorXml(shadow.color, shadow.alpha)}</a:outerShdw></a:effectLst>`
  );
}

export function shapeXml(shape, id, ctx) {
  const geometry = shape.geometry || "rect";
  const isTextBox = geometry === "rect" && !shape.fill && !shape.line;
  const nvSpPr =
    `<p:nvSpPr><p:cNvPr id="${id}" name="${isTextBox ? "TextBox" : "Shape"} ${id}"/>` +
    `<p:cNvSpPr${isTextBox ? ' txBox="1"' : ""}/><p:nvPr/></p:nvSpPr>`;
  // spPr child order: xfrm, geometry, fill, ln, effectLst.
  const spPr =
    `<p:spPr>${xfrmXml(shape)}${geometryXml({ ...shape, geometry })}` +
    `${shape.fill ? fillXml(shape.fill, shape.fillAlpha) : "<a:noFill/>"}${lineXml(shape.line)}${shape.shadow ? shadowXml(shape.shadow) : ""}</p:spPr>`;
  const body = txBodyXml(shape.paras || [], ctx, { anchor: shape.anchor || "t", inset: shape.inset, wrap: shape.wrap });
  return `<p:sp>${nvSpPr}${spPr}${body}</p:sp>`;
}

// Slide background (p:bg): a color, a gradient, or a picture (rId) stretched over the slide.
export function backgroundXml(bg, rId) {
  const fill = rId
    ? `<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="${rId}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`
    : bg.gradient ? gradFillXml(bg.gradient) : solidFill(bg.color, bg.alpha);
  return `<p:bg><p:bgPr>${fill}<a:effectLst/></p:bgPr></p:bg>`;
}

// PowerPoint 2016+ draws the SVG; older PowerPoint, Google Slides, Keynote and LibreOffice use the PNG.
const SVG_EXT = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";

export function pictureXml({ id, name, descr, rId, svgRId, box, crop, line, shadow }) {
  const alt = descr ? ` descr="${esc(descr)}"` : "";
  // srcRect crops in 1/1000 of a percent of the image, per side.
  const sides = crop ? ["l", "t", "r", "b"].filter((k) => crop[k] > 0) : [];
  const srcRect = sides.length ? `<a:srcRect${sides.map((k) => ` ${k}="${Math.round(crop[k] * 100000)}"`).join("")}/>` : "";
  const blip = svgRId
    ? `<a:blip r:embed="${rId}"><a:extLst><a:ext uri="${SVG_EXT}"><asvg:svgBlip xmlns:asvg="${ASVG}" r:embed="${svgRId}"/></a:ext></a:extLst></a:blip>`
    : `<a:blip r:embed="${rId}"/>`;
  return (
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${esc(name)}"${alt}/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill>${blip}${srcRect}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr>${xfrmXml(box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${line ? lineXml(line) : ""}${shadow ? shadowXml(shadow) : ""}</p:spPr></p:pic>`
  );
}

export function graphicFrameXml({ id, name, box, uri, inner }) {
  return (
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${esc(name)}"/>` +
    `<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
    `${xfrmXml(box, "p:xfrm")}<a:graphic><a:graphicData uri="${uri}">${inner}</a:graphicData></a:graphic></p:graphicFrame>`
  );
}

export function slideXml(shapesXml, bgXml = "") {
  return (
    `${XML_DECL}<p:sld ${NS_PML}><p:cSld>${bgXml}<p:spTree>${SP_TREE_HEADER}${shapesXml}</p:spTree></p:cSld>` +
    "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
  );
}
