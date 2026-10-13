import { test } from "node:test";
import assert from "node:assert/strict";
import { txBodyXml, bodyPrXml } from "../src/ooxml/text.mjs";
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
