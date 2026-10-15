import { test } from "node:test";
import assert from "node:assert/strict";
import { placeholderXml, shapeXml, pictureXml, graphicFrameXml, slideXml, phXml } from "../src/ooxml/shapes.mjs";
import { parseXml } from "./helpers.mjs";

const ctx = { lang: "en-US", fonts: { mono: "M" } };
const box = { x: 1, y: 0.5, w: 2, h: 1 };

test("placeholder without a box inherits geometry from the layout", () => {
  const xml = placeholderXml({ id: 2, name: "Title 1", ph: { type: "title" }, paras: [{ runs: [{ text: "Hi" }] }], ctx });
  assert.match(xml, /^<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"\/><p:cNvSpPr><a:spLocks noGrp="1"\/><\/p:cNvSpPr><p:nvPr><p:ph type="title"\/><\/p:nvPr><\/p:nvSpPr><p:spPr\/><p:txBody><a:bodyPr\/>/);
});

test("placeholder with a box writes a slide-level xfrm", () => {
  const xml = placeholderXml({ id: 3, name: "Body", ph: { idx: 1 }, box, paras: [], ctx, body: { fontScale: 0.9 } });
  assert.match(xml, /<p:ph idx="1"\/>/);
  assert.match(xml, /<p:spPr><a:xfrm><a:off x="914400" y="457200"\/><a:ext cx="1828800" cy="914400"\/><\/a:xfrm><\/p:spPr>/);
  assert.match(xml, /<a:normAutofit fontScale="90000"/);
});

test("phXml writes only the attributes that are set", () => {
  assert.equal(phXml({ type: "body", idx: 1, sz: "half" }), '<p:ph type="body" idx="1" sz="half"/>');
  assert.equal(phXml({ type: "ctrTitle" }), '<p:ph type="ctrTitle"/>');
});

test("plain text shape is a txBox with no fill and no line", () => {
  const xml = shapeXml({ kind: "text", ...box, paras: [{ runs: [{ text: "x" }] }] }, 4, ctx);
  assert.match(xml, /<p:cNvSpPr txBox="1"\/>/);
  assert.match(xml, /<a:prstGeom prst="rect"><a:avLst\/><\/a:prstGeom><a:noFill\/><a:ln><a:noFill\/><\/a:ln><\/p:spPr>/);
  assert.match(xml, /<a:bodyPr wrap="square" anchor="t" rtlCol="0"\/>/);
});

test("card shape: roundRect radius, fill, dashed line, shadow, in spPr schema order", () => {
  const xml = shapeXml({ kind: "text", x: 0, y: 0, w: 2, h: 1, geometry: "roundRect", radius: 0.1, fill: "lt2", line: { color: "accent4", width: 1, dash: "dash" }, shadow: true }, 5, ctx);
  assert.match(xml, /<p:cNvSpPr\/>/);
  assert.match(xml, /<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"\/><\/a:avLst><\/a:prstGeom><a:solidFill><a:schemeClr val="bg2"\/><\/a:solidFill><a:ln w="12700"><a:solidFill><a:schemeClr val="accent4"\/><\/a:solidFill><a:prstDash val="dash"\/><\/a:ln><a:effectLst><a:outerShdw /);
  assert.match(xml, /<a:outerShdw blurRad="114300" dist="38100" dir="5400000" algn="t" rotWithShape="0"><a:schemeClr val="tx2"><a:alpha val="16000"\/><\/a:schemeClr><\/a:outerShdw>/);
});

test("picture carries alt text as descr", () => {
  const xml = pictureXml({ id: 6, name: "Picture 6", descr: 'A "quoted" <alt>', rId: "rId3", box });
  assert.match(xml, /<p:cNvPr id="6" name="Picture 6" descr="A &quot;quoted&quot; &lt;alt&gt;"\/>/);
  assert.match(xml, /<a:blip r:embed="rId3"\/><a:stretch><a:fillRect\/><\/a:stretch>/);
});

test("graphic frame wraps graphicData with the given uri", () => {
  const xml = graphicFrameXml({ id: 7, name: "Chart 7", box, uri: "urn:x", inner: "<x/>" });
  assert.match(xml, /<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"\/><\/p:cNvGraphicFramePr>/);
  assert.match(xml, /<p:xfrm><a:off x="914400" y="457200"\/><a:ext cx="1828800" cy="914400"\/><\/p:xfrm><a:graphic><a:graphicData uri="urn:x"><x\/><\/a:graphicData><\/a:graphic>/);
});

test("slideXml is a well-formed slide with clrMapOvr", () => {
  const xml = slideXml(shapeXml({ kind: "text", ...box, paras: [] }, 2, ctx));
  parseXml(xml);
  assert.match(xml, /<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""\/>/);
  assert.match(xml, /<\/p:cSld><p:clrMapOvr><a:masterClrMapping\/><\/p:clrMapOvr><\/p:sld>$/);
});

test("wrap: false on a free text shape writes wrap=\"none\"", () => {
  const xml = shapeXml({ x: 0, y: 0, w: 1, h: 1, wrap: false, paras: [{ runs: [{ text: "x" }] }] }, 5, ctx);
  assert.match(xml, /<a:bodyPr wrap="none" anchor="t" rtlCol="0"\/>/);
});

test("shape fill and outline alpha, a measured shadow and a rotation", () => {
  const xml = shapeXml({
    x: 1, y: 1, w: 2, h: 1, rot: -90, fill: "lt2", fillAlpha: 0.5,
    line: { width: 1, color: "#D9E3DB", alpha: 0.4 },
    shadow: { blur: 0.1, dist: 0.05, dir: 90, color: "#14181A", alpha: 0.14 }, paras: [],
  }, 3, ctx);
  assert.match(xml, /<a:xfrm rot="16200000">/);
  assert.match(xml, /<a:solidFill><a:schemeClr val="bg2"><a:alpha val="50000"\/><\/a:schemeClr><\/a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="D9E3DB"><a:alpha val="40000"\/><\/a:srgbClr><\/a:solidFill><\/a:ln>/);
  assert.match(xml, /<a:effectLst><a:outerShdw blurRad="91440" dist="45720" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="14181A"><a:alpha val="14000"\/><\/a:srgbClr><\/a:outerShdw><\/a:effectLst>/);
});

test("pictures can be cropped, outlined and shadowed; a plain picture is unchanged", () => {
  const xml = pictureXml({ id: 4, name: "P", rId: "rId2", box: { x: 0, y: 0, w: 1, h: 1 }, crop: { l: 0.1, t: 0, r: 0.1, b: 0 }, line: { width: 0.75, color: "accent4" }, shadow: true });
  assert.match(xml, /<a:blip r:embed="rId2"\/><a:srcRect l="10000" r="10000"\/><a:stretch>/);
  assert.match(xml, /<\/a:prstGeom><a:ln w="9525"><a:solidFill><a:schemeClr val="accent4"\/><\/a:solidFill><\/a:ln><a:effectLst>/);
  const plain = pictureXml({ id: 4, name: "P", rId: "rId2", box: { x: 0, y: 0, w: 1, h: 1 } });
  assert.match(plain, /<a:blip r:embed="rId2"\/><a:stretch><a:fillRect\/><\/a:stretch><\/p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"\/><a:ext cx="914400" cy="914400"\/><\/a:xfrm><a:prstGeom prst="rect"><a:avLst\/><\/a:prstGeom><\/p:spPr>/);
});
