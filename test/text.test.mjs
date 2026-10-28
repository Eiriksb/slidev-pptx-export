import { test } from "node:test";
import assert from "node:assert/strict";
import { txBodyXml, bodyPrXml, rPrXml, pPrXml } from "../src/ooxml/text.mjs";
import { parseXml, byLocal } from "./helpers.mjs";

const ctx = { lang: "en-US", fonts: { mono: "JetBrains Mono" }, link: (url) => (url === "https://x.test" ? "rId7" : "rId8") };
const wrap = (xml) => `<p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">${xml}</p:sp>`;

test("a bullet with mixed formatting is ONE paragraph with at most one pPr, first", () => {
  const xml = txBodyXml(
    [{ bullet: "char", runs: [{ text: "Bold", b: true, color: "dk2" }, { text: " for emphasis, " }, { text: "code", font: "mono" }] }],
    ctx,
    {}
  );
  const doc = parseXml(wrap(xml));
  const ps = byLocal(doc, "p");
  assert.equal(ps.length, 1);
  const kids = [];
  for (let c = ps[0].firstChild; c; c = c.nextSibling) if (c.nodeType === 1) kids.push(c.localName);
  assert.equal(kids.filter((k) => k === "pPr").length, 1);
  assert.equal(kids[0], "pPr");
  assert.deepEqual(kids.slice(1), ["r", "r", "r", "endParaRPr"]);
});

test("free text box bullets write indent and bullet children in schema order", () => {
  const xml = txBodyXml([{ bullet: "char", runs: [{ text: "x" }] }], ctx, {});
  assert.match(xml, /<a:pPr marL="228600" indent="-182880"><a:buClr><a:schemeClr val="accent1"\/><\/a:buClr><a:buFont typeface="Arial"\/><a:buChar char="•"\/><\/a:pPr>/);
});

test("placeholder bullets inherit; none and number are explicit", () => {
  const xml = txBodyXml(
    [
      { bullet: "char", runs: [{ text: "a" }] },
      { bullet: "char", level: 1, runs: [{ text: "b" }] },
      { bullet: "none", runs: [{ text: "c" }] },
      { bullet: "number", number: 3, runs: [{ text: "d" }] },
    ],
    ctx,
    { placeholder: true }
  );
  assert.match(xml, /<a:p><a:r><a:rPr lang="en-US" dirty="0"\/><a:t>a<\/a:t>/);
  assert.match(xml, /<a:pPr lvl="1"\/><a:r>/);
  assert.match(xml, /<a:pPr marL="0" indent="0"><a:buNone\/><\/a:pPr>/);
  assert.match(xml, /<a:buFont typeface="\+mj-lt"\/><a:buAutoNum type="arabicPeriod" startAt="3"\/>/);
});

test("spaceBefore and alignment go on pPr", () => {
  const xml = txBodyXml([{ align: "ctr", spaceBefore: 12, runs: [{ text: "x" }] }], ctx, {});
  assert.match(xml, /<a:pPr algn="ctr"><a:spcBef><a:spcPts val="1200"\/><\/a:spcBef><\/a:pPr>/);
});

test("runs: size, bold, italic, spacing, color, theme fonts, mono, link", () => {
  const xml = txBodyXml(
    [{ runs: [
      { text: "T", size: 26, b: true, i: true, spacing: 1.5, color: "accent1", font: "heading" },
      { text: "c", font: "mono" },
      { text: "l", font: "body", link: "https://x.test" },
    ] }],
    ctx,
    {}
  );
  assert.match(xml, /<a:rPr lang="en-US" sz="2600" b="1" i="1" spc="150" dirty="0"><a:solidFill><a:schemeClr val="accent1"\/><\/a:solidFill><a:latin typeface="\+mj-lt"\/><a:ea typeface="\+mj-ea"\/><a:cs typeface="\+mj-cs"\/><\/a:rPr><a:t>T<\/a:t>/);
  assert.match(xml, /<a:latin typeface="JetBrains Mono"\/><a:cs typeface="JetBrains Mono"\/><\/a:rPr><a:t>c<\/a:t>/);
  assert.match(xml, /<a:latin typeface="\+mn-lt"\/><a:ea typeface="\+mn-ea"\/><a:cs typeface="\+mn-cs"\/><a:hlinkClick r:id="rId7"\/><\/a:rPr><a:t>l<\/a:t>/);
});

test("b: false is written as b=\"0\" (regular text under a bold title style); an unset b inherits", () => {
  assert.equal(rPrXml({ text: "x", b: false }, ctx), '<a:rPr lang="en-US" b="0" dirty="0"/>');
  assert.equal(rPrXml({ text: "x" }, ctx), '<a:rPr lang="en-US" dirty="0"/>');
});

test("newlines become a:br and text is escaped", () => {
  const xml = txBodyXml([{ runs: [{ text: "a < b\nc & d" }] }], ctx, {});
  assert.match(xml, /<a:t>a &lt; b<\/a:t><\/a:r><a:br><a:rPr lang="en-US" dirty="0"\/><\/a:br><a:r><a:rPr lang="en-US" dirty="0"\/><a:t>c &amp; d<\/a:t>/);
});

test("empty bodies still contain one paragraph", () => {
  assert.match(txBodyXml([], ctx, {}), /<a:lstStyle\/><a:p><a:endParaRPr lang="en-US" dirty="0"\/><\/a:p><\/p:txBody>$/);
});

test("bodyPr: attributes only when set, autofit children", () => {
  assert.equal(bodyPrXml({ placeholder: true }), "<a:bodyPr/>");
  assert.equal(bodyPrXml({ placeholder: true, fontScale: 0.85 }), '<a:bodyPr><a:normAutofit fontScale="85000" lnSpcReduction="10000"/></a:bodyPr>');
  assert.equal(bodyPrXml({ anchor: "ctr", inset: 0.1 }), '<a:bodyPr wrap="square" lIns="91440" tIns="91440" rIns="91440" bIns="91440" anchor="ctr" rtlCol="0"/>');
  assert.equal(bodyPrXml({ placeholder: true, autofit: "norm", anchor: "b" }), '<a:bodyPr anchor="b"><a:normAutofit/></a:bodyPr>');
  assert.equal(bodyPrXml({ wrap: false }), '<a:bodyPr wrap="none" rtlCol="0"/>');
});

test("a run with a link but no link callback throws", () => {
  assert.throws(() => txBodyXml([{ runs: [{ text: "x", link: "https://x" }] }], { lang: "en-US", fonts: { mono: "M" } }, {}), /link/);
});

test("runs can name any typeface, and carry underline, strike, alpha and a highlight in schema order", () => {
  const xml = rPrXml({ text: "x", typeface: "Plus Jakarta Sans", color: "#112233", alpha: 0.5, highlight: "#F5F5F5", u: true, strike: true }, ctx);
  assert.equal(
    xml,
    '<a:rPr lang="en-US" u="sng" strike="sngStrike" dirty="0"><a:solidFill><a:srgbClr val="112233"><a:alpha val="50000"/></a:srgbClr></a:solidFill>' +
      '<a:highlight><a:srgbClr val="F5F5F5"/></a:highlight><a:latin typeface="Plus Jakarta Sans"/><a:cs typeface="Plus Jakarta Sans"/></a:rPr>',
  );
});

test("a typeface wins over a font role, and the last run's typeface carries into endParaRPr", () => {
  const xml = txBodyXml([{ runs: [{ text: "a", typeface: "Inter", font: "heading", size: 12 }] }], ctx);
  assert.match(xml, /<a:endParaRPr lang="en-US" sz="1200" dirty="0"><a:latin typeface="Inter"\/><a:cs typeface="Inter"\/><\/a:endParaRPr>/);
  assert.doesNotMatch(xml, /\+mj-lt/);
});

test("free text box paragraphs take exact line spacing, a bullet color and measured margins", () => {
  assert.equal(
    pPrXml({ bullet: "char", bulletColor: "#1A833B", lineSpacing: 21, marL: 0.3, indent: -0.15 }, ctx),
    '<a:pPr marL="274320" indent="-137160"><a:lnSpc><a:spcPts val="2100"/></a:lnSpc><a:buClr><a:srgbClr val="1A833B"/></a:buClr><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>',
  );
  assert.equal(
    pPrXml({ bullet: "number", number: 3, bulletColor: "accent1" }, ctx),
    '<a:pPr marL="228600" indent="-182880"><a:buClr><a:schemeClr val="accent1"/></a:buClr><a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod" startAt="3"/></a:pPr>',
  );
  assert.equal(pPrXml({ lineSpacing: 18.5 }, ctx), '<a:pPr><a:lnSpc><a:spcPts val="1850"/></a:lnSpc></a:pPr>');
  assert.equal(pPrXml({ bullet: "none", marL: 0.5 }, ctx), '<a:pPr marL="457200" indent="0"/>');
});

test("placeholders can switch off wrapping and insets too (render mode's measured titles)", () => {
  assert.equal(bodyPrXml({ placeholder: true, wrap: false, inset: 0, anchor: "t" }), '<a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"/>');
  assert.equal(bodyPrXml({ placeholder: true }), "<a:bodyPr/>");
});
