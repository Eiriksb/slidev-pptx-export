import { test } from "node:test";
import assert from "node:assert/strict";
import { slideGeometry } from "../src/geometry.mjs";
import { layoutDefs, layoutXml, masterXml } from "../src/ooxml/master.mjs";
import { parseXml, byLocal } from "./helpers.mjs";

const g = slideGeometry({ width: 13.333, height: 7.5, marginX: 0.6 });
const ctx = { lang: "en-US", fonts: { mono: "M" } };

test("geometry reserves the eyebrow band above the title", () => {
  assert.equal(g.eyebrow.x, 0.6);
  assert.equal(g.eyebrow.y, 0.45);
  assert.equal(g.eyebrow.h, 0.3);
  assert.ok(Math.abs(g.eyebrow.w - 12.133) < 1e-9);
  assert.equal(g.title.y, 0.8);
  assert.equal(g.body.y, 1.6);
  assert.ok(Math.abs(g.body.h - 5.55) < 1e-9);
  assert.ok(Math.abs(g.coverTitle.y - 2.7) < 1e-9);
  assert.ok(Math.abs(g.coverSubtitle.y - 4.1) < 1e-9);
});

test("layout catalogue has the expected keys, types and placeholders", () => {
  const defs = layoutDefs(g);
  assert.deepEqual(defs.map((d) => [d.key, d.type]), [
    ["title", "title"], ["content", "obj"], ["section", "secHead"], ["twoContent", "twoObj"],
    ["titleOnly", "titleOnly"], ["centered", "cust"], ["blank", "blank"],
  ]);
  const byKey = Object.fromEntries(defs.map((d) => [d.key, d]));
  assert.deepEqual(byKey.title.placeholders.map((p) => p.ph), [{ type: "ctrTitle" }, { type: "subTitle", idx: 1 }]);
  assert.deepEqual(byKey.content.placeholders.map((p) => p.role), ["title", "body"]);
  assert.deepEqual(byKey.blank.placeholders, []);
});

test("every layout part is well-formed and typed", () => {
  for (const def of layoutDefs(g)) {
    const xml = layoutXml(def, ctx);
    const doc = parseXml(xml, def.name);
    assert.equal(doc.documentElement.getAttribute("type"), def.type);
    assert.equal(byLocal(doc, "cSld")[0].getAttribute("name"), def.name);
    assert.equal(byLocal(doc, "ph").length, def.placeholders.length);
  }
});

test("cover title style lives in the layout's lstStyle", () => {
  const def = layoutDefs(g).find((d) => d.key === "title");
  const xml = layoutXml(def, ctx);
  assert.match(xml, /<a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="4000" b="1"><a:solidFill><a:schemeClr val="accent2"\/><\/a:solidFill><\/a:defRPr><\/a:lvl1pPr><\/a:lstStyle>/);
});

test("F3: section body lstStyle no longer sets buNone, so char bullets inherit a real bullet", () => {
  const def = layoutDefs(g).find((d) => d.key === "section");
  const xml = layoutXml(def, ctx);
  const bodyDef = def.placeholders.find((p) => p.role === "body");
  assert.doesNotMatch(bodyDef.lstStyle, /buNone/);
  assert.match(xml, /<a:lstStyle><a:lvl1pPr><a:defRPr sz="1600"><a:solidFill><a:schemeClr val="accent3"\/><\/a:solidFill><\/a:defRPr><\/a:lvl1pPr><\/a:lstStyle>/);
});

test("master: clrMap, layout id list, txStyles with theme fonts and accent bullets, logo", () => {
  const xml = masterXml({ geometry: g, ctx, layoutIds: [{ id: 2147483649, rId: "rId2" }], logo: { rId: "rId9", box: { x: 0.5, y: 6.94, w: 0.98, h: 0.27 } } });
  parseXml(xml);
  assert.match(xml, /<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1"/);
  assert.match(xml, /<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId2"\/><\/p:sldLayoutIdLst>/);
  assert.match(xml, /<p:titleStyle><a:lvl1pPr[^>]*>.*<a:defRPr sz="2600" b="1" kern="1200">.*<a:latin typeface="\+mj-lt"\/>/);
  assert.match(xml, /<p:bodyStyle><a:lvl1pPr marL="228600" indent="-182880"[^>]*>.*<a:buClr><a:schemeClr val="accent1"\/><\/a:buClr><a:buFont typeface="Arial"\/><a:buChar char="•"\/><a:defRPr sz="1300"/);
  assert.match(xml, /<a:blip r:embed="rId9"\/>/);
  // cSld, clrMap, sldLayoutIdLst, txStyles - in schema order.
  const order = ["<p:cSld>", "<p:clrMap ", "<p:sldLayoutIdLst>", "<p:txStyles>"].map((s) => xml.indexOf(s));
  assert.deepEqual([...order].sort((a, b) => a - b), order);
});
