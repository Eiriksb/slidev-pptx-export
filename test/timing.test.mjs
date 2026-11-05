import { test } from "node:test";
import assert from "node:assert/strict";
import { animationTimingXml, transitionXml } from "../src/ooxml/timing.mjs";
import { writePptx } from "../src/ooxml/index.mjs";
import { DECK, PIC } from "./deck.mjs";
import { lintPackage } from "./lint.mjs";
import { unzip } from "./helpers.mjs";

const count = (xml, re) => (xml.match(re) || []).length;
const fx = (spid, click, dir, kind = "image", fill = false) => ({ spid, kind, fill, click, dir });

test("no videos and no effects: no timing at all", () => {
  assert.equal(animationTimingXml({ videos: [], effects: [] }), "");
});

test("effects of one click: the first is On Click, the rest With Previous", () => {
  const xml = animationTimingXml({ videos: [], effects: [fx(4, 1, "in"), fx(5, 1, "out")] });
  assert.equal(count(xml, /nodeType="clickEffect"/g), 1);
  assert.equal(count(xml, /nodeType="withEffect"/g), 1);
  assert.match(xml, /presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect"/);
  assert.match(xml, /<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="\d+" dur="400"\/>/);
  assert.match(xml, /presetID="10" presetClass="exit"/);
  assert.match(xml, /<p:animEffect transition="out" filter="fade"><p:cBhvr><p:cTn id="\d+" dur="400"\/>/);
  assert.match(xml, /<p:strVal val="hidden"\/>/);
});

test("a shape that fades out hides itself only once the fade has finished", () => {
  const xml = animationTimingXml({ videos: [], effects: [fx(5, 1, "out")] });
  assert.match(xml, /<p:cond delay="399"\/>.*<p:strVal val="hidden"\/>/);
});

test("each click is its own step, in order", () => {
  const xml = animationTimingXml({ videos: [], effects: [fx(5, 2, "in"), fx(4, 1, "in"), fx(6, 3, "in")] });
  assert.equal(count(xml, /nodeType="clickEffect"/g), 3);
  assert.equal(count(xml, /<p:cond delay="indefinite"\/>/g), 3);
  assert.ok(xml.indexOf('spid="4"') < xml.indexOf('spid="5"') && xml.indexOf('spid="5"') < xml.indexOf('spid="6"'));
});

test("a text shape with two effects gets grpId 0 and 1 and a bldP for each; a fill adds animBg", () => {
  const xml = animationTimingXml({ videos: [], effects: [fx(7, 1, "in", "text", true), fx(7, 2, "out", "text", true), fx(8, 1, "in")] });
  assert.equal(count(xml, /grpId="0"/g), 2);
  assert.equal(count(xml, /grpId="1"/g), 2);
  assert.match(xml, /<p:bldLst><p:bldP spid="7" grpId="0" animBg="1"\/><p:bldP spid="7" grpId="1" animBg="1"\/><\/p:bldLst>/);
  assert.doesNotMatch(xml, /spid="8" grpId/);
});

test("a text shape with no fill has no animBg", () => {
  const xml = animationTimingXml({ videos: [], effects: [fx(7, 1, "in", "text")] });
  assert.match(xml, /<p:bldP spid="7" grpId="0"\/>/);
});

test("a video that starts by itself is written as PowerPoint for the web writes it", () => {
  const xml = animationTimingXml({ videos: [{ spid: 6, autoplay: true, loop: false, muted: true, duration: 3 }], effects: [] });
  assert.match(xml, /<p:cond evt="onBegin" delay="0"><p:tn val="2"\/><\/p:cond>/);
  assert.match(xml, /presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect"/);
  assert.match(xml, /cmd="playFrom\(0\.0\)"><p:cBhvr><p:cTn id="\d+" dur="1" fill="hold"\/>/);
  assert.match(xml, /<p:video><p:cMediaNode mute="1"><p:cTn id="\d+" fill="hold" display="0">/);
  assert.doesNotMatch(xml, /togglePause|repeatCount/);
});

test("a looping video keeps repeatCount; one that does not autoplay has only its media node", () => {
  const looping = animationTimingXml({ videos: [{ spid: 6, autoplay: true, loop: true, muted: true }], effects: [] });
  assert.match(looping, /repeatCount="indefinite"/);
  const waiting = animationTimingXml({ videos: [{ spid: 6, autoplay: false, muted: false }], effects: [] });
  assert.doesNotMatch(waiting, /mediacall/);
  assert.match(waiting, /<p:video><p:cMediaNode>/);
});

test("the clip starts first and On Click effects follow in the same main sequence, with unique ids", () => {
  const xml = animationTimingXml({ videos: [{ spid: 6, autoplay: true, muted: true }], effects: [fx(9, 1, "in", "text")] });
  assert.equal(count(xml, /nodeType="mainSeq"/g), 1);
  assert.ok(xml.indexOf("mediacall") < xml.indexOf("clickEffect"));
  const ids = [...xml.matchAll(/<p:cTn id="(\d+)"/g)].map((m) => m[1]);
  assert.equal(new Set(ids).size, ids.length);
});

test("transitions", () => {
  assert.equal(transitionXml("fade"), "<p:transition><p:fade/></p:transition>");
  assert.equal(transitionXml({ push: "l" }), '<p:transition><p:push dir="l"/></p:transition>');
  assert.equal(transitionXml(undefined), "");
});

test("the writer: effects reach the timing, the transition comes before it, the package is clean", async () => {
  const shapes = [{ kind: "image", x: 1, y: 1, w: 2, h: 1, path: PIC, effects: [{ click: 1, dir: "in" }] }];
  const buffer = await writePptx([{ layout: "blank", shapes, animated: true, transition: "fade" }], DECK);
  assert.deepEqual(await lintPackage(buffer), []);
  const slide = (await unzip(buffer)).text("ppt/slides/slide1.xml");
  assert.match(slide, /<\/p:clrMapOvr><p:transition><p:fade\/><\/p:transition><p:timing>/);
  assert.match(slide, /nodeType="clickEffect"/);
});

test("the writer: a spec that is not animated keeps the existing video timing", async () => {
  const shapes = [{ kind: "image", x: 1, y: 1, w: 2, h: 1, path: PIC, effects: [{ click: 1, dir: "in" }] }];
  const slide = (await unzip(await writePptx([{ layout: "blank", shapes }], DECK))).text("ppt/slides/slide1.xml");
  assert.doesNotMatch(slide, /p:timing|p:transition/);
});
