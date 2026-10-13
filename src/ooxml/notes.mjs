// Speaker notes: a notes master (written only when some slide has notes) and one notes slide
// per slide that has them. Positions are PowerPoint's own defaults for a 7.5x10in notes page.

import { XML_DECL, NS_PML } from "./xml.mjs";
import { SP_TREE_HEADER, placeholderXml, xfrmXml } from "./shapes.mjs";

export function notesParas(text) {
  const paras = [];
  for (const raw of String(text ?? "").split("\n")) {
    const line = raw.trim();
    if (line === "") {
      if (paras.length && paras[paras.length - 1].runs.length) paras.push({ runs: [] });
      continue;
    }
    paras.push({ runs: [{ text: line }] });
  }
  while (paras.length && !paras[paras.length - 1].runs.length) paras.pop();
  return paras;
}

const CLR_MAP =
  '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" ' +
  'accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';

const SLIDE_IMAGE_LOCKS = '<a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/>';

export function notesMasterXml(ctx) {
  const image =
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr>${SLIDE_IMAGE_LOCKS}</p:cNvSpPr>` +
    '<p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr>' +
    `<p:spPr>${xfrmXml({ x: 0.75, y: 1.25, w: 6, h: 3.375 })}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>` +
    '<a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr></p:sp>';
  const body = placeholderXml({
    id: 3, name: "Notes Placeholder 2", ph: { type: "body", idx: 3, sz: "quarter" }, ctx,
    box: { x: 0.75, y: 4.8125, w: 6, h: 3.9375 },
    paras: [{ runs: [{ text: "Click to edit Master text styles" }] }],
  });
  const style =
    '<p:notesStyle><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">' +
    '<a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>' +
    '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:notesStyle>';
  return (
    `${XML_DECL}<p:notesMaster ${NS_PML}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>` +
    `<p:spTree>${SP_TREE_HEADER}${image}${body}</p:spTree></p:cSld>${CLR_MAP}${style}</p:notesMaster>`
  );
}

export function notesSlideXml(text, ctx) {
  const image =
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr>${SLIDE_IMAGE_LOCKS}</p:cNvSpPr>` +
    '<p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>';
  const body = placeholderXml({ id: 3, name: "Notes Placeholder 2", ph: { type: "body", idx: 1 }, ctx, paras: notesParas(text) });
  return (
    `${XML_DECL}<p:notes ${NS_PML}><p:cSld><p:spTree>${SP_TREE_HEADER}${image}${body}</p:spTree></p:cSld>` +
    "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>"
  );
}
