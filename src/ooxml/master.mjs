// Slide master + the layout catalogue. Layouts carry real types (title, obj, secHead, ...) so
// PowerPoint, Google Slides and Keynote all recognize them, and their placeholders carry the
// geometry, anchors and autofit that slide placeholders inherit.

import { XML_DECL, NS_PML, esc, solidFill } from "./xml.mjs";
import { SP_TREE_HEADER, placeholderXml, pictureXml } from "./shapes.mjs";

// Level-1 list style override for one layout placeholder (e.g. the cover's big centered title).
function lvl1({ algn, sz, b, color, buNone }) {
  const attrs = [buNone ? 'marL="0" indent="0"' : "", algn ? `algn="${algn}"` : ""].filter(Boolean).join(" ");
  const rAttrs = `${sz ? ` sz="${sz * 100}"` : ""}${b ? ' b="1"' : ""}`;
  const defRPr = color ? `<a:defRPr${rAttrs}>${solidFill(color)}</a:defRPr>` : `<a:defRPr${rAttrs}/>`;
  return `<a:lvl1pPr${attrs ? ` ${attrs}` : ""}>${buNone ? "<a:buNone/>" : ""}${defRPr}</a:lvl1pPr>`;
}

const TITLE_PROMPT = "Click to edit Master title style";
const BODY_PROMPT = "Click to edit Master text styles";

export function layoutDefs(g) {
  const title = { role: "title", ph: { type: "title" }, name: "Title 1", box: g.title, anchor: "b", prompt: TITLE_PROMPT };
  const body = { role: "body", ph: { idx: 1 }, name: "Content Placeholder 2", box: g.body, anchor: "t", prompt: BODY_PROMPT };
  const half = (g.contentW - 0.3) / 2;
  return [
    {
      key: "title", name: "Title Slide", type: "title",
      placeholders: [
        { role: "title", ph: { type: "ctrTitle" }, name: "Title 1", box: g.coverTitle, anchor: "ctr", prompt: TITLE_PROMPT,
          lstStyle: lvl1({ algn: "ctr", sz: 40, b: true, color: "accent2" }) },
        { role: "subtitle", ph: { type: "subTitle", idx: 1 }, name: "Subtitle 2", box: g.coverSubtitle, anchor: "t", prompt: "Click to edit Master subtitle style",
          lstStyle: lvl1({ algn: "ctr", sz: 18, buNone: true }) },
      ],
    },
    { key: "content", name: "Title and Content", type: "obj", placeholders: [title, body] },
    {
      key: "section", name: "Section Header", type: "secHead",
      placeholders: [
        { ...title, box: g.sectionTitle, lstStyle: lvl1({ sz: 34, b: true }) },
        { role: "body", ph: { type: "body", idx: 1 }, name: "Text Placeholder 2", box: g.sectionBody, anchor: "t", prompt: BODY_PROMPT,
          lstStyle: lvl1({ sz: 16, color: "accent3", buNone: true }) },
      ],
    },
    {
      key: "twoContent", name: "Two Content", type: "twoObj",
      placeholders: [
        title,
        { ...body, ph: { sz: "half", idx: 1 }, box: { ...g.body, w: half } },
        { ...body, role: "body2", ph: { sz: "half", idx: 2 }, name: "Content Placeholder 3", box: { ...g.body, x: g.marginX + half + 0.3, w: half } },
      ],
    },
    { key: "titleOnly", name: "Title Only", type: "titleOnly", placeholders: [title] },
    {
      key: "centered", name: "Centered", type: "cust",
      placeholders: [{ ...title, box: g.centeredTitle }, { ...body, box: g.centeredBody }],
    },
    { key: "blank", name: "Blank", type: "blank", placeholders: [] },
  ];
}

export function layoutXml(def, ctx) {
  const sps = def.placeholders
    .map((p, i) =>
      placeholderXml({
        id: i + 2, name: p.name, ph: p.ph, box: p.box, ctx,
        paras: [{ runs: [{ text: p.prompt }] }],
        body: { anchor: p.anchor, autofit: "norm", lstStyle: p.lstStyle },
      })
    )
    .join("");
  return (
    `${XML_DECL}<p:sldLayout ${NS_PML} type="${def.type}" preserve="1"><p:cSld name="${esc(def.name)}">` +
    `<p:spTree>${SP_TREE_HEADER}${sps}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`
  );
}

const CLR_MAP =
  '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" ' +
  'accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';

const P_ATTRS = 'algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"';

function defRPr(sz, { b = false, font = "mn" } = {}) {
  return (
    `<a:defRPr sz="${sz * 100}"${b ? ' b="1"' : ""} kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>` +
    `<a:latin typeface="+${font}-lt"/><a:ea typeface="+${font}-ea"/><a:cs typeface="+${font}-cs"/></a:defRPr>`
  );
}

const TITLE_STYLE =
  `<p:titleStyle><a:lvl1pPr ${P_ATTRS}><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef>` +
  `<a:buNone/>${defRPr(26, { b: true, font: "mj" })}</a:lvl1pPr></p:titleStyle>`;

// Body levels 1-3: 0.25"/0.5"/0.75" left margin with a 0.2" hanging bullet, 13/12/11 pt,
// accent1 bullets - the deck's brand color comes from the master, not from each slide.
const BODY_LEVELS = [[228600, 13], [457200, 12], [685800, 11]];
const BODY_STYLE = `<p:bodyStyle>${BODY_LEVELS.map(
  ([marL, sz], i) =>
    `<a:lvl${i + 1}pPr marL="${marL}" indent="-182880" ${P_ATTRS}><a:lnSpc><a:spcPct val="100000"/></a:lnSpc>` +
    '<a:spcBef><a:spcPts val="600"/></a:spcBef><a:buClr><a:schemeClr val="accent1"/></a:buClr><a:buFont typeface="Arial"/><a:buChar char="•"/>' +
    `${defRPr(sz)}</a:lvl${i + 1}pPr>`
).join("")}</p:bodyStyle>`;

const OTHER_STYLE =
  `<p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr marL="0" ${P_ATTRS}>${defRPr(18)}</a:lvl1pPr></p:otherStyle>`;

export function masterXml({ geometry: g, ctx, layoutIds, logo }) {
  const title = placeholderXml({
    id: 2, name: "Title Placeholder 1", ph: { type: "title" }, box: g.title, ctx,
    paras: [{ runs: [{ text: TITLE_PROMPT }] }], body: { anchor: "b", autofit: "norm" },
  });
  const body = placeholderXml({
    id: 3, name: "Text Placeholder 2", ph: { type: "body", idx: 1 }, box: g.body, ctx,
    paras: [
      { bullet: "char", runs: [{ text: BODY_PROMPT }] },
      { bullet: "char", level: 1, runs: [{ text: "Second level" }] },
      { bullet: "char", level: 2, runs: [{ text: "Third level" }] },
    ],
    body: { anchor: "t", autofit: "norm" },
  });
  const pic = logo ? pictureXml({ id: 4, name: "Logo", descr: "Logo", rId: logo.rId, box: logo.box }) : "";
  const ids = layoutIds.map(({ id, rId }) => `<p:sldLayoutId id="${id}" r:id="${rId}"/>`).join("");
  return (
    `${XML_DECL}<p:sldMaster ${NS_PML}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>` +
    `<p:spTree>${SP_TREE_HEADER}${title}${body}${pic}</p:spTree></p:cSld>${CLR_MAP}` +
    `<p:sldLayoutIdLst>${ids}</p:sldLayoutIdLst><p:txStyles>${TITLE_STYLE}${BODY_STYLE}${OTHER_STYLE}</p:txStyles></p:sldMaster>`
  );
}
