import { readFileSync } from "node:fs";
import { test } from "node:test";
import assert from "node:assert/strict";
import { writePptx } from "../src/ooxml/index.mjs";
import { timingXml } from "../src/ooxml/shapes.mjs";
import { shapesOf } from "../src/render/scene.mjs";
import { DECK, PIC } from "./deck.mjs";
import { lintPackage } from "./lint.mjs";
import { unzip } from "./helpers.mjs";

// Not a real video: the writer stores the bytes as they are, PowerPoint reads them when played.
const MP4 = Buffer.from("00000018667479706d703432000000006d703432", "hex");
const POSTER = { data: readFileSync(PIC), ext: "png" };
const video = (over = {}) => ({ kind: "video", x: 1, y: 1, w: 4, h: 2.25, data: MP4, ext: "mp4", poster: POSTER, alt: "Clip", autoplay: true, loop: true, muted: true, duration: 12.5, ...over });

async function build(shapes) {
  const buffer = await writePptx([{ layout: "blank", shapes }], DECK);
  return { buffer, z: await unzip(buffer) };
}

test("a video is a picture with the file linked and embedded, the poster as its picture", async () => {
  const { buffer, z } = await build([video()]);
  assert.deepEqual(await lintPackage(buffer), []);
  const slide = z.text("ppt/slides/slide1.xml");
  assert.match(slide, /<p:cNvPr id="2" name="Video 2" descr="Clip"><a:hlinkClick r:id="" action="ppaction:\/\/media"\/><\/p:cNvPr>/);
  const link = /<a:videoFile r:link="(rId\d+)"\/>/.exec(slide)[1];
  const embed = /<p14:media xmlns:p14="[^"]+" r:embed="(rId\d+)"\/>/.exec(slide)[1];
  const blip = /<a:blip r:embed="(rId\d+)"\/>/.exec(slide)[1];
  const rels = z.text("ppt/slides/_rels/slide1.xml.rels");
  const target = (id) => new RegExp(`Id="${id}" Type="([^"]+)" Target="([^"]+)"`).exec(rels).slice(1);
  const [vType, vFile] = target(link);
  assert.equal(vType, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video");
  assert.match(vFile, /^\.\.\/media\/media\d+\.mp4$/);
  assert.deepEqual(target(embed), ["http://schemas.microsoft.com/office/2007/relationships/media", vFile]);
  assert.match(target(blip)[1], /^\.\.\/media\/image\d+\.png$/);
  assert.match(z.text("[Content_Types].xml"), /<Default Extension="mp4" ContentType="video\/mp4"\/>/);
  assert.deepEqual(z.buffer(`ppt/media/${vFile.split("/").pop()}`), MP4);
});

test("an autoplaying, looping, muted video starts with the slide and repeats", async () => {
  const { z } = await build([video()]);
  const slide = z.text("ppt/slides/slide1.xml");
  assert.match(slide, /<\/p:clrMapOvr><p:timing>/);
  assert.match(slide, /<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">/);
  assert.match(slide, /<p:cond evt="onBegin" delay="0"><p:tn val="2"\/><\/p:cond>/);
  assert.match(slide, /presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect">.*<p:cmd type="call" cmd="playFrom\(0.0\)"><p:cBhvr><p:cTn id="\d+" dur="12500" fill="hold"\/><p:tgtEl><p:spTgt spid="2"\/>/);
  assert.match(slide, /<p:video><p:cMediaNode vol="80000" mute="1"><p:cTn id="\d+" repeatCount="indefinite" fill="hold" display="0">.*<p:spTgt spid="2"\/><\/p:tgtEl><\/p:cMediaNode><\/p:video>/);
});

test("timing: several autoplaying videos start together; one that does not autoplay waits for a click", () => {
  const xml = timingXml([{ spid: 4, autoplay: true }, { spid: 5, autoplay: true, loop: true }, { spid: 6, autoplay: false }]);
  assert.equal((xml.match(/nodeType="afterEffect"/g) || []).length, 1);
  assert.equal((xml.match(/nodeType="withEffect"/g) || []).length, 1);
  assert.doesNotMatch(xml, /spid="6"\/><\/p:tgtEl><\/p:cBhvr>/);
  assert.equal((xml.match(/<p:video>/g) || []).length, 3);
  const ids = [...xml.matchAll(/<p:cTn id="(\d+)"/g)].map((m) => Number(m[1]));
  assert.equal(new Set(ids).size, ids.length);
  assert.equal(timingXml([]), "");
});

test("the same video twice is stored once; slides without videos have no timing", async () => {
  const { z } = await build([video(), video({ x: 6, autoplay: false })]);
  assert.equal(z.names.filter((n) => n.endsWith(".mp4")).length, 1);
  const plain = await unzip(await writePptx([{ layout: "blank", shapes: [{ kind: "image", x: 1, y: 1, w: 2, h: 1, path: PIC }] }], DECK));
  assert.doesNotMatch(plain.text("ppt/slides/slide1.xml"), /p:timing/);
});

test("scene: a video node becomes a video shape, cropped as object-fit: cover draws it", () => {
  const u = { box: (r) => ({ x: r.x / 100, y: r.y / 100, w: r.w / 100, h: r.h / 100 }) };
  const node = { kind: "video", rect: { x: 0, y: 0, w: 400, h: 400 }, data: MP4, ext: "mp4", posterImage: POSTER, fit: "cover", natural: { w: 1600, h: 900 }, autoplay: true, loop: true, muted: true, duration: 3 };
  const [shape] = shapesOf(node, u, {});
  assert.equal(shape.kind, "video");
  assert.deepEqual([shape.x, shape.y, shape.w, shape.h], [0, 0, 4, 4]);
  assert.ok(shape.crop.l > 0.2 && shape.crop.r > 0.2 && !shape.crop.t);
  assert.equal(shape.poster, POSTER);
  assert.deepEqual(shapesOf({ ...node, data: undefined }, u, {}), []);
});
