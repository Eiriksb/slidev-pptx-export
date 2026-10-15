import { readFileSync } from "node:fs";
import { test } from "node:test";
import assert from "node:assert/strict";
import { writePptx } from "../src/ooxml/index.mjs";
import { DECK, PIC, geometry } from "./deck.mjs";
import { lintPackage } from "./lint.mjs";
import { unzip, byLocal } from "./helpers.mjs";

const t = (text) => [{ runs: [{ text, font: "heading" }] }];

const SLIDES = [
  { layout: "title", placeholders: { title: t("Cover"), subtitle: [{ runs: [{ text: "Sub" }] }] }, shapes: [] },
  {
    layout: "content",
    placeholders: {
      title: t("Bullets"),
      body: [
        { bullet: "char", runs: [{ text: "One " }, { text: "bold", b: true, color: "dk2" }] },
        { bullet: "number", runs: [{ text: "Two" }] },
      ],
      bodyFontScale: 0.85,
    },
    shapes: [{ kind: "text", ...geometry.eyebrow, paras: [{ runs: [{ text: "EYEBROW", link: "https://example.com/?a=1&b=2" }] }] }],
  },
  {
    layout: "titleOnly",
    placeholders: { title: t("Pictures") },
    shapes: [
      { kind: "image", x: 1, y: 2, w: 2, h: 1, path: PIC, alt: "A picture" },
      { kind: "image", x: 4, y: 2, w: 2, h: 1, path: PIC, alt: "Same picture again" },
    ],
  },
  {
    layout: "centered",
    placeholders: { title: t("Centered"), titleBox: { x: 0.6, y: 3, w: 12.133, h: 0.6 }, body: [{ bullet: "none", align: "ctr", runs: [{ text: "Questions?" }] }], bodyBox: { x: 0.6, y: 3.8, w: 12.133, h: 0.5 } },
    shapes: [],
  },
  { layout: "blank", shapes: [{ kind: "text", x: 1, y: 1, w: 3, h: 1, geometry: "roundRect", fill: "lt2", shadow: true, paras: [] }] },
];

async function build(slides = SLIDES, deck = DECK) {
  const buffer = await writePptx(slides, deck);
  return { buffer, z: await unzip(buffer) };
}

function layoutNameOf(z, slideNo) {
  const rels = z.text(`ppt/slides/_rels/slide${slideNo}.xml.rels`);
  const target = /Type="[^"]+\/slideLayout" Target="..\/slideLayouts\/(slideLayout\d+\.xml)"/.exec(rels)[1];
  return byLocal(z.xml(`ppt/slideLayouts/${target}`), "cSld")[0].getAttribute("name");
}

test("package passes lint", async () => {
  const { buffer } = await build();
  assert.deepEqual(await lintPackage(buffer), []);
});

test("each slide uses the requested layout", async () => {
  const { z } = await build();
  assert.deepEqual([1, 2, 3, 4, 5].map((n) => layoutNameOf(z, n)), ["Title Slide", "Title and Content", "Title Only", "Centered", "Blank"]);
});

test("titles live in real title placeholders that inherit geometry", async () => {
  const { z } = await build();
  const s1 = z.text("ppt/slides/slide1.xml");
  assert.match(s1, /<p:ph type="ctrTitle"\/><\/p:nvPr><\/p:nvSpPr><p:spPr\/>.*<a:t>Cover<\/a:t>/);
  assert.match(s1, /<p:ph type="subTitle" idx="1"\/>.*<a:t>Sub<\/a:t>/);
  const s2 = z.text("ppt/slides/slide2.xml");
  assert.match(s2, /<p:ph type="title"\/>.*<a:t>Bullets<\/a:t>/);
  assert.match(s2, /<p:ph idx="1"\/>.*<a:normAutofit fontScale="85000"/);
});

test("centered slide overrides title and body position on the slide", async () => {
  const { z } = await build();
  assert.match(z.text("ppt/slides/slide4.xml"), /<p:ph type="title"\/><\/p:nvPr><\/p:nvSpPr><p:spPr><a:xfrm><a:off x="548640" y="2743200"\/>/);
});

test("hyperlinks become external relationships", async () => {
  const { z } = await build();
  assert.match(z.text("ppt/slides/_rels/slide2.xml.rels"), /Target="https:\/\/example.com\/\?a=1&amp;b=2" TargetMode="External"/);
  assert.match(z.text("ppt/slides/slide2.xml"), /<a:hlinkClick r:id="rId2"\/>/);
});

test("identical images are stored once and carry alt text", async () => {
  const { z } = await build();
  assert.deepEqual(z.names.filter((n) => n.startsWith("ppt/media/")), ["ppt/media/image1.png"]);
  assert.match(z.text("ppt/slides/slide3.xml"), /descr="A picture"/);
  assert.match(z.text("ppt/slideMasters/slideMaster1.xml"), /name="Logo" descr="Logo"/);
});

test("presentation ids, size and theme", async () => {
  const { z } = await build();
  const pres = z.text("ppt/presentation.xml");
  assert.match(pres, /<p:sldMasterId id="2147483648" r:id="rId1"\/>/);
  assert.deepEqual([...pres.matchAll(/<p:sldId id="(\d+)"/g)].map((m) => Number(m[1])), [256, 257, 258, 259, 260]);
  assert.match(pres, /<p:sldSz cx="12191695" cy="6858000"\/><p:notesSz cx="6858000" cy="9144000"\/>/);
  assert.match(z.text("ppt/slideMasters/slideMaster1.xml"), /<p:sldLayoutId id="2147483649" /);
  assert.match(z.text("ppt/theme/theme1.xml"), /<a:accent1><a:srgbClr val="1A833B"\/><\/a:accent1>/);
});

test("core.xml has no dates unless timestamps are on", async () => {
  const off = await build();
  assert.doesNotMatch(off.z.text("docProps/core.xml"), /dcterms:created/);
  assert.match(off.z.text("docProps/core.xml"), /<dc:title>Writer test<\/dc:title><dc:creator>Tester<\/dc:creator>/);
  const on = await build(SLIDES, { ...DECK, timestamps: true });
  assert.match(on.z.text("docProps/core.xml"), /<dcterms:created xsi:type="dcterms:W3CDTF">\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ<\/dcterms:created>/);
});

test("same input gives byte-identical output", async () => {
  const a = await writePptx(SLIDES, DECK);
  const b = await writePptx(SLIDES, DECK);
  assert.ok(a.equals(b));
});

test("unknown layouts and shape kinds are rejected", async () => {
  await assert.rejects(writePptx([{ layout: "nope" }], DECK), /unknown layout "nope"/);
  await assert.rejects(writePptx([{ layout: "blank", shapes: [{ kind: "video" }] }], DECK), /Unknown shape kind "video"/);
});

test("a slide's title placeholder can carry its own body properties", async () => {
  const buffer = await writePptx([{ layout: "titleOnly", placeholders: { title: t("Measured"), titleBox: { x: 1, y: 1, w: 3, h: 0.5 }, titleBody: { wrap: false, inset: 0, anchor: "t" } }, shapes: [] }], DECK);
  assert.match((await unzip(buffer)).text("ppt/slides/slide1.xml"), /<p:ph type="title"\/>[\s\S]*?<a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"\/>/);
});

test("images from buffers share media with file images; slides take solid and picture backgrounds", async () => {
  const data = readFileSync(PIC);
  const slides = [
    { layout: "blank", background: { color: "#FAFAFA" }, shapes: [{ kind: "image", x: 1, y: 1, w: 2, h: 1, path: PIC }, { kind: "image", x: 4, y: 1, w: 2, h: 1, data, ext: "png", crop: { l: 0.25, t: 0, r: 0.25, b: 0 } }] },
    { layout: "blank", background: { image: { data, ext: "png" } }, shapes: [] },
  ];
  const buffer = await writePptx(slides, { ...DECK, logo: null });
  assert.deepEqual(await lintPackage(buffer), []);
  const z = await unzip(buffer);
  assert.deepEqual(z.names.filter((n) => n.startsWith("ppt/media/")), ["ppt/media/image1.png"]);
  assert.match(z.text("ppt/slides/slide1.xml"), /<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FAFAFA"\/><\/a:solidFill><a:effectLst\/><\/p:bgPr><\/p:bg><p:spTree>/);
  assert.match(z.text("ppt/slides/slide1.xml"), /<a:srcRect l="25000" r="25000"\/>/);
  assert.match(z.text("ppt/slides/slide2.xml"), /<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rId2"\/><a:srcRect\/><a:stretch><a:fillRect\/><\/a:stretch><\/a:blipFill><a:effectLst\/><\/p:bgPr><\/p:bg>/);
});

test("a title with titleZ is drawn among the shapes at its DOM position, not beneath them", async () => {
  const slides = [
    {
      layout: "titleOnly",
      placeholders: { title: t("TITLE"), titleZ: 1 },
      shapes: [
        { kind: "text", x: 1, y: 1, w: 3, h: 1, geometry: "rect", fill: "lt2", paras: [{ runs: [{ text: "FIRST" }] }] },
        { kind: "text", x: 1, y: 2, w: 3, h: 1, paras: [{ runs: [{ text: "SECOND" }] }] },
      ],
    },
    { layout: "titleOnly", placeholders: { title: t("LATE"), titleZ: 5 }, shapes: [{ kind: "text", x: 1, y: 1, w: 3, h: 1, paras: [{ runs: [{ text: "ONLY" }] }] }] },
  ];
  const z = await unzip(await writePptx(slides, DECK));
  const s1 = z.text("ppt/slides/slide1.xml");
  const titleAt = s1.indexOf('<p:ph type="title"/>');
  assert.equal(s1.split('<p:ph type="title"/>').length - 1, 1);
  assert.ok(s1.indexOf("<a:t>FIRST</a:t>") < titleAt, "the title is drawn above the first shape");
  assert.ok(titleAt < s1.indexOf("<a:t>SECOND</a:t>"), "the title is drawn beneath the second shape");
  const s2 = z.text("ppt/slides/slide2.xml");
  assert.ok(s2.indexOf("<a:t>ONLY</a:t>") < s2.indexOf('<p:ph type="title"/>'), "a titleZ past the last shape goes on top");
});
