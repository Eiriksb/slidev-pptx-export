// DrawingML text bodies. Every paragraph gets at most one <a:pPr>, always first, and every
// pPr/rPr child is written in schema order - pptxgenjs 4.0.1 writes a pPr per *run*, which
// is what dropped bullets in LibreOffice and is a likely PowerPoint Online repair trigger.

import { esc, emu, pt100, solidFill } from "./xml.mjs";

const THEME_FONTS = {
  heading: ["+mj-lt", "+mj-ea", "+mj-cs"],
  body: ["+mn-lt", "+mn-ea", "+mn-cs"],
};

function fontXml(font, ctx) {
  if (!font) return "";
  if (font === "mono") {
    const f = esc(ctx.fonts.mono);
    return `<a:latin typeface="${f}"/><a:cs typeface="${f}"/>`;
  }
  const refs = THEME_FONTS[font];
  if (!refs) throw new Error(`Unknown font role ${JSON.stringify(font)}`);
  return `<a:latin typeface="${refs[0]}"/><a:ea typeface="${refs[1]}"/><a:cs typeface="${refs[2]}"/>`;
}

export function rPrXml(run, ctx, tag = "a:rPr") {
  let attrs = ` lang="${esc(ctx.lang)}"`;
  if (run.size) attrs += ` sz="${Math.min(400000, Math.max(100, pt100(run.size)))}"`;
  if (run.b) attrs += ' b="1"';
  if (run.i) attrs += ' i="1"';
  if (run.spacing) attrs += ` spc="${pt100(run.spacing)}"`;
  attrs += ' dirty="0"';
  // Child order per CT_TextCharacterProperties: fill, ..., latin, ea, cs, sym, hlinkClick.
  let kids = "";
  if (run.color) kids += solidFill(run.color);
  kids += fontXml(run.font, ctx);
  if (run.link) {
    if (!ctx.link) throw new Error("Run has a link but the text context has no link() callback");
    kids += `<a:hlinkClick r:id="${ctx.link(run.link)}"/>`;
  }
  return kids ? `<${tag}${attrs}>${kids}</${tag}>` : `<${tag}${attrs}/>`;
}

function runsXml(runs, ctx) {
  let out = "";
  for (const run of runs) {
    const pieces = String(run.text ?? "").split("\n");
    pieces.forEach((piece, i) => {
      if (i > 0) out += `<a:br>${rPrXml({ ...run, link: undefined }, ctx)}</a:br>`;
      if (piece !== "") out += `<a:r>${rPrXml(run, ctx)}<a:t>${esc(piece)}</a:t></a:r>`;
    });
  }
  return out;
}

export function pPrXml(para, ctx, { placeholder = false } = {}) {
  const level = Math.max(0, Math.min(8, para.level || 0));
  const attrs = [];
  const kids = [];
  const bullet = para.bullet;

  if (bullet === "none" && placeholder) {
    // Body placeholders inherit bullets from the master's bodyStyle, so a plain paragraph has
    // to switch them off and drop the hanging indent explicitly.
    attrs.push('marL="0"', 'indent="0"');
  } else if ((bullet === "char" || bullet === "number") && !placeholder) {
    // Free text boxes inherit no list style, so they carry their own indent.
    attrs.push(`marL="${emu(0.25 * (level + 1))}"`, `indent="${-emu(0.2)}"`);
  }
  if (level > 0) attrs.push(`lvl="${level}"`);
  if (para.align && para.align !== "l") attrs.push(`algn="${para.align}"`);

  // Child order per CT_TextParagraphProperties: lnSpc, spcBef, spcAft, buClr, buSz, buFont, bullet.
  if (para.spaceBefore != null) kids.push(`<a:spcBef><a:spcPts val="${pt100(para.spaceBefore)}"/></a:spcBef>`);
  if (bullet === "char" && !placeholder) {
    kids.push('<a:buClr><a:schemeClr val="accent1"/></a:buClr>', '<a:buFont typeface="Arial"/>', '<a:buChar char="•"/>');
  } else if (bullet === "number") {
    const startAt = para.number && para.number !== 1 ? ` startAt="${para.number}"` : "";
    kids.push('<a:buFont typeface="+mj-lt"/>', `<a:buAutoNum type="arabicPeriod"${startAt}/>`);
  } else if (bullet === "none" && placeholder) {
    kids.push("<a:buNone/>");
  }

  if (!attrs.length && !kids.length) return "";
  const a = attrs.length ? ` ${attrs.join(" ")}` : "";
  return kids.length ? `<a:pPr${a}>${kids.join("")}</a:pPr>` : `<a:pPr${a}/>`;
}

function paraXml(para, ctx, opts) {
  const runs = para.runs || [];
  const last = runs[runs.length - 1] || {};
  // endParaRPr keeps an empty trailing line (and an empty paragraph) at the run's size/font
  // instead of PowerPoint's 18pt default.
  const end = rPrXml({ size: last.size, font: last.font }, ctx, "a:endParaRPr");
  return `<a:p>${pPrXml(para, ctx, opts)}${runsXml(runs, ctx)}${end}</a:p>`;
}

export function bodyPrXml(body = {}) {
  const attrs = [];
  if (!body.placeholder) {
    if (body.wrap === false) attrs.push('wrap="none"');
    else attrs.push('wrap="square"');
  }
  if (body.inset != null) {
    const i = emu(body.inset);
    attrs.push(`lIns="${i}" tIns="${i}" rIns="${i}" bIns="${i}"`);
  }
  if (body.anchor) attrs.push(`anchor="${body.anchor}"`);
  if (!body.placeholder) attrs.push('rtlCol="0"');
  let kid = "";
  if (body.fontScale != null && body.fontScale < 1) {
    kid = `<a:normAutofit fontScale="${Math.round(body.fontScale * 100000)}" lnSpcReduction="10000"/>`;
  } else if (body.autofit === "norm") kid = "<a:normAutofit/>";
  else if (body.autofit === "shape") kid = "<a:spAutoFit/>";
  else if (body.autofit === "none") kid = "<a:noAutofit/>";
  const a = attrs.length ? ` ${attrs.join(" ")}` : "";
  return kid ? `<a:bodyPr${a}>${kid}</a:bodyPr>` : `<a:bodyPr${a}/>`;
}

export function txBodyXml(paras, ctx, body = {}, tag = "p:txBody") {
  const list = paras && paras.length ? paras : [{ runs: [] }];
  const ps = list.map((p) => paraXml(p, ctx, { placeholder: !!body.placeholder })).join("");
  const lst = body.lstStyle ? `<a:lstStyle>${body.lstStyle}</a:lstStyle>` : "<a:lstStyle/>";
  return `<${tag}>${bodyPrXml(body)}${lst}${ps}</${tag}>`;
}
