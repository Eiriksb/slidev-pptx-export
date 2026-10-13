import { test } from "node:test";
import assert from "node:assert/strict";
import { Package, REL, CT } from "../src/ooxml/package.mjs";
import { unzip } from "./helpers.mjs";

function samplePackage() {
  const pkg = new Package();
  pkg.addXml("/ppt/presentation.xml", CT.presentation, "<x/>");
  pkg.addXml("/ppt/slides/slide1.xml", CT.slide, "<s/>");
  pkg.addMedia("/ppt/media/image1.png", Buffer.from([1, 2, 3]));
  pkg.rel("", "/ppt/presentation.xml", REL.officeDocument);
  pkg.rel("/ppt/presentation.xml", "/ppt/slides/slide1.xml", REL.slide);
  pkg.rel("/ppt/slides/slide1.xml", "/ppt/media/image1.png", REL.image);
  return pkg;
}

test("content types list exactly the parts that were added", async () => {
  const z = await unzip(await samplePackage().toBuffer());
  const ct = z.text("[Content_Types].xml");
  assert.match(ct, /<Override PartName="\/ppt\/presentation.xml" ContentType="application\/vnd.openxmlformats-officedocument.presentationml.presentation.main\+xml"\/>/);
  assert.match(ct, /<Override PartName="\/ppt\/slides\/slide1.xml"/);
  assert.match(ct, /<Default Extension="png" ContentType="image\/png"\/>/);
  assert.equal((ct.match(/<Override /g) || []).length, 2);
  assert.equal(z.names[0], "[Content_Types].xml");
});

test("relationship targets are relative to the source part", async () => {
  const z = await unzip(await samplePackage().toBuffer());
  assert.match(z.text("_rels/.rels"), /Target="ppt\/presentation.xml"/);
  assert.match(z.text("ppt/_rels/presentation.xml.rels"), /Id="rId1" Type="[^"]+\/slide" Target="slides\/slide1.xml"/);
  assert.match(z.text("ppt/slides/_rels/slide1.xml.rels"), /Target="..\/media\/image1.png"/);
});

test("rel() reuses identical relationships and numbers new ones per part", () => {
  const pkg = new Package();
  assert.equal(pkg.rel("/a.xml", "/b.xml", REL.slide), "rId1");
  assert.equal(pkg.rel("/a.xml", "/b.xml", REL.slide), "rId1");
  assert.equal(pkg.rel("/a.xml", "https://example.com", REL.hyperlink, { external: true }), "rId2");
  assert.equal(pkg.rel("/c.xml", "/b.xml", REL.slide), "rId1");
});

test("external relationships carry TargetMode and are not resolved", async () => {
  const pkg = samplePackage();
  pkg.rel("/ppt/slides/slide1.xml", "https://example.com/?a=1&b=2", REL.hyperlink, { external: true });
  const z = await unzip(await pkg.toBuffer());
  assert.match(z.text("ppt/slides/_rels/slide1.xml.rels"), /Target="https:\/\/example.com\/\?a=1&amp;b=2" TargetMode="External"/);
});

test("toBuffer refuses a relationship to a missing part", async () => {
  const pkg = samplePackage();
  pkg.rel("/ppt/presentation.xml", "/ppt/slides/slide2.xml", REL.slide);
  await assert.rejects(pkg.toBuffer(), /missing part \/ppt\/slides\/slide2.xml/);
});

test("duplicate parts and unknown media types are rejected", () => {
  const pkg = samplePackage();
  assert.throws(() => pkg.addXml("/ppt/slides/slide1.xml", CT.slide, "<s/>"), /Duplicate part/);
  assert.throws(() => pkg.addMedia("/ppt/media/x.bmp", Buffer.from([])), /Unsupported media/);
});

test("output is byte-identical across runs", async () => {
  const a = await samplePackage().toBuffer();
  const b = await samplePackage().toBuffer();
  assert.ok(a.equals(b));
});
