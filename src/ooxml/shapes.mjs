// PresentationML shape builders: placeholders, free text/geometry shapes, pictures and
// graphic frames (tables, charts). Geometry arrives in inches; ids come from the caller's
// per-part counter so they stay unique within a slide.

import { XML_DECL, NS_PML, esc, emu, emuSize, solidFill, colorXml } from "./xml.mjs";
import { txBodyXml } from "./text.mjs";

export const SP_TREE_HEADER =
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

export function xfrmXml(box, tag = "a:xfrm") {
  return `<${tag}><a:off x="${emu(box.x)}" y="${emu(box.y)}"/><a:ext cx="${emuSize(box.w)}" cy="${emuSize(box.h)}"/></${tag}>`;
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
  return `<a:ln w="${Math.round(line.width * 12700)}">${solidFill(line.color)}${dash}</a:ln>`;
}

const SHADOW =
  '<a:effectLst><a:outerShdw blurRad="114300" dist="38100" dir="5400000" algn="t" rotWithShape="0">' +
  `${colorXml("dk2", 0.16)}</a:outerShdw></a:effectLst>`;

export function shapeXml(shape, id, ctx) {
  const geometry = shape.geometry || "rect";
  const isTextBox = geometry === "rect" && !shape.fill && !shape.line;
  const nvSpPr =
    `<p:nvSpPr><p:cNvPr id="${id}" name="${isTextBox ? "TextBox" : "Shape"} ${id}"/>` +
    `<p:cNvSpPr${isTextBox ? ' txBox="1"' : ""}/><p:nvPr/></p:nvSpPr>`;
  // spPr child order: xfrm, geometry, fill, ln, effectLst.
  const spPr =
    `<p:spPr>${xfrmXml(shape)}${geometryXml({ ...shape, geometry })}` +
    `${shape.fill ? solidFill(shape.fill) : "<a:noFill/>"}${lineXml(shape.line)}${shape.shadow ? SHADOW : ""}</p:spPr>`;
  const body = txBodyXml(shape.paras || [], ctx, { anchor: shape.anchor || "t", inset: shape.inset, wrap: shape.wrap });
  return `<p:sp>${nvSpPr}${spPr}${body}</p:sp>`;
}

export function pictureXml({ id, name, descr, rId, box }) {
  const alt = descr ? ` descr="${esc(descr)}"` : "";
  return (
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${esc(name)}"${alt}/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr>${xfrmXml(box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
  );
}

export function graphicFrameXml({ id, name, box, uri, inner }) {
  return (
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${esc(name)}"/>` +
    `<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
    `${xfrmXml(box, "p:xfrm")}<a:graphic><a:graphicData uri="${uri}">${inner}</a:graphicData></a:graphic></p:graphicFrame>`
  );
}

export function slideXml(shapesXml) {
  return (
    `${XML_DECL}<p:sld ${NS_PML}><p:cSld><p:spTree>${SP_TREE_HEADER}${shapesXml}</p:spTree></p:cSld>` +
    "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
  );
}
