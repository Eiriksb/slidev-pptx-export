// presentation.xml plus the small package-level parts PowerPoint expects beside it.

import { XML_DECL, NS_PML, esc, emuSize } from "./xml.mjs";

const DEFAULT_TEXT_STYLE =
  '<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>' +
  '<a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">' +
  '<a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>' +
  '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:defaultTextStyle>';

// PowerPoint only accepts slide sizes between 1" and 56".
const clampSize = (inches) => Math.min(51206400, Math.max(914400, emuSize(inches)));

export function presentationXml({ masterId, masterRId, notesMasterRId, slides, width, height, embeddedFonts = [] }) {
  // embeddedFontLst comes after notesSz; embedTrueTypeFonts tells PowerPoint the file carries them.
  const fonts = embeddedFonts.length
    ? `<p:embeddedFontLst>${embeddedFonts.map((f) => `<p:embeddedFont><p:font typeface="${esc(f.typeface)}"/>` +
      ["regular", "bold", "italic", "boldItalic"].filter((k) => f[k]).map((k) => `<p:${k} r:id="${f[k]}"/>`).join("") +
      "</p:embeddedFont>").join("")}</p:embeddedFontLst>`
    : "";
  // Child order per CT_Presentation: sldMasterIdLst, notesMasterIdLst, sldIdLst, sldSz,
  // notesSz, defaultTextStyle. (pptxgenjs writes sldIdLst before notesMasterIdLst.)
  const notes = notesMasterRId ? `<p:notesMasterIdLst><p:notesMasterId r:id="${notesMasterRId}"/></p:notesMasterIdLst>` : "";
  const sldIds = slides.length ? `<p:sldIdLst>${slides.map((s) => `<p:sldId id="${s.id}" r:id="${s.rId}"/>`).join("")}</p:sldIdLst>` : "";
  return (
    `${XML_DECL}<p:presentation ${NS_PML}${embeddedFonts.length ? ' embedTrueTypeFonts="1"' : ""}><p:sldMasterIdLst><p:sldMasterId id="${masterId}" r:id="${masterRId}"/></p:sldMasterIdLst>` +
    `${notes}${sldIds}<p:sldSz cx="${clampSize(width)}" cy="${clampSize(height)}"/><p:notesSz cx="6858000" cy="9144000"/>` +
    `${fonts}${DEFAULT_TEXT_STYLE}</p:presentation>`
  );
}

export function presPropsXml() {
  return `${XML_DECL}<p:presentationPr ${NS_PML}/>`;
}

export function viewPropsXml() {
  return (
    `${XML_DECL}<p:viewPr ${NS_PML}><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr>` +
    '<p:gridSpacing cx="76200" cy="76200"/></p:viewPr>'
  );
}

export function appXml({ slides, notes }) {
  return (
    `${XML_DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ` +
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    `<Application>slidev-pptx-export</Application><PresentationFormat>Custom</PresentationFormat>` +
    `<Slides>${slides}</Slides><Notes>${notes}</Notes></Properties>`
  );
}

// Dates are left out unless asked for, so two exports of the same deck are byte-identical.
export function coreXml({ title, author, date }) {
  const iso = date ? date.toISOString().replace(/\.\d{3}Z$/, "Z") : null;
  const dates = iso
    ? `<dcterms:created xsi:type="dcterms:W3CDTF">${iso}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${iso}</dcterms:modified>`
    : "";
  return (
    `${XML_DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
    'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    `<dc:title>${esc(title)}</dc:title><dc:creator>${esc(author)}</dc:creator><cp:lastModifiedBy>${esc(author)}</cp:lastModifiedBy>` +
    `<cp:revision>1</cp:revision>${dates}</cp:coreProperties>`
  );
}
