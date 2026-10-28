import { existsSync, readFileSync } from "node:fs";
import { test } from "node:test";
import assert from "node:assert/strict";
import { eotFromTtf, readEot, fontInfo, embeddable } from "../src/ooxml/eot.mjs";
import { writePptx } from "../src/ooxml/index.mjs";
import { DECK } from "./deck.mjs";
import { lintPackage } from "./lint.mjs";
import { unzip } from "./helpers.mjs";

// A minimal TrueType file: the tables EOT needs (name, OS/2, head) and an empty glyf.
function tinyTtf({ family = "Tiny Sans", style = "Regular", weight = 400, italic = false, fsType = 0 } = {}) {
  const utf16be = (s) => Buffer.from(s, "utf16le").swap16();
  const names = [[1, family], [2, style], [4, `${family} ${style}`], [5, "Version 1.0"]];
  const strings = Buffer.concat(names.map(([, s]) => utf16be(s)));
  const name = Buffer.alloc(6 + 12 * names.length);
  name.writeUInt16BE(0, 0);
  name.writeUInt16BE(names.length, 2);
  name.writeUInt16BE(name.length, 4);
  let off = 0;
  names.forEach(([id, s], i) => {
    const r = 6 + 12 * i;
    [3, 1, 0x409, id, utf16be(s).length, off].forEach((v, k) => name.writeUInt16BE(v, r + 2 * k));
    off += utf16be(s).length;
  });
  const os2 = Buffer.alloc(96);
  os2.writeUInt16BE(4, 0);
  os2.writeUInt16BE(weight, 4);
  os2.writeUInt16BE(fsType, 8);
  os2.writeUInt8(2, 32); // PANOSE family kind
  os2.writeUInt32BE(0x00000001, 42); // Basic Latin
  os2.writeUInt16BE(italic ? 1 : 0x40, 62);
  os2.writeUInt32BE(0x00000001, 78); // Latin 1
  const head = Buffer.alloc(54);
  head.writeUInt32BE(0x12345678, 8);
  const tables = [["OS/2", os2], ["glyf", Buffer.alloc(4)], ["head", head], ["name", Buffer.concat([name, strings])]];
  const dir = Buffer.alloc(12 + 16 * tables.length);
  dir.writeUInt32BE(0x00010000, 0);
  dir.writeUInt16BE(tables.length, 4);
  let at = dir.length;
  const bodies = [];
  tables.forEach(([tag, body], i) => {
    const o = 12 + 16 * i;
    dir.write(tag, o, "latin1");
    dir.writeUInt32BE(at, o + 8);
    dir.writeUInt32BE(body.length, o + 12);
    const padded = Buffer.concat([body, Buffer.alloc((4 - (body.length % 4)) % 4)]);
    bodies.push(padded);
    at += padded.length;
  });
  return Buffer.concat([dir, ...bodies]);
}

test("EOT: the header PowerPoint reads, then the TrueType file as it is", () => {
  const ttf = tinyTtf({ weight: 700, italic: true });
  const eot = eotFromTtf(ttf);
  const h = readEot(eot);
  assert.equal(h.eotSize, eot.length);
  assert.equal(h.fontDataSize, ttf.length);
  assert.equal(h.version, 0x00020001);
  assert.equal(h.flags, 0);
  assert.equal(h.magic, 0x504c);
  assert.equal(h.weight, 700);
  assert.equal(h.italic, 1);
  assert.deepEqual(h.names, ["Tiny Sans", "Regular", "Version 1.0", "Tiny Sans Regular"]);
  assert.deepEqual(Buffer.from(h.fontData), ttf);
  assert.equal(eot.readUInt32LE(60), 0x12345678); // head.checkSumAdjustment
});

test("EOT: a font whose licence forbids embedding, or with PostScript outlines, is refused", () => {
  assert.equal(embeddable(0), true);
  assert.equal(embeddable(4), true);
  assert.equal(embeddable(2), false);
  assert.throws(() => eotFromTtf(tinyTtf({ fsType: 2 })), /license does not allow embedding/);
  const otf = tinyTtf();
  otf.write("OTTO", 0, "latin1");
  assert.throws(() => eotFromTtf(otf), /TrueType fonts only/);
  assert.deepEqual(fontInfo(tinyTtf({ family: "X", style: "Bold", weight: 700 })), { family: "X", style: "Bold", weight: 700, fsType: 0, italic: false });
});

test("embedded fonts: parts, relationships, embeddedFontLst, and web font names renamed to the family", async () => {
  const slides = [{ layout: "blank", shapes: [{ kind: "text", x: 1, y: 1, w: 3, h: 1, paras: [{ runs: [{ text: "Hei", typeface: "Tiny Sans Variable" }, { text: "!", typeface: "Missing Font" }] }] }] }];
  const resolver = (name) => (name === "Tiny Sans Variable" ? { family: "Tiny Sans", fonts: { regular: eotFromTtf(tinyTtf()), bold: eotFromTtf(tinyTtf({ weight: 700 })) }, notes: [] } : null);
  const notices = [];
  const buffer = await writePptx(slides, { ...DECK, fontResolver: resolver, onFontNotice: (m) => notices.push(m) });
  assert.deepEqual(await lintPackage(buffer), []);
  const z = await unzip(buffer);
  const pres = z.text("ppt/presentation.xml");
  assert.match(pres, /<p:presentation [^>]*embedTrueTypeFonts="1"/);
  assert.match(pres, /<p:notesSz [^>]*\/><p:embeddedFontLst><p:embeddedFont><p:font typeface="Tiny Sans"\/><p:regular r:id="(rId\d+)"\/><p:bold r:id="rId\d+"\/><\/p:embeddedFont><\/p:embeddedFontLst>/);
  const rels = z.text("ppt/_rels/presentation.xml.rels");
  assert.match(rels, /Type="http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships\/font" Target="fonts\/font1.fntdata"/);
  assert.match(z.text("[Content_Types].xml"), /<Default Extension="fntdata" ContentType="application\/x-fontdata"\/>/);
  assert.equal(readEot(z.buffer("ppt/fonts/font1.fntdata")).weight, 400);
  const slide = z.text("ppt/slides/slide1.xml");
  assert.match(slide, /typeface="Tiny Sans"/);
  assert.doesNotMatch(slide, /Tiny Sans Variable/);
  assert.match(slide, /typeface="Missing Font"/);
  assert.ok(notices.some((m) => /"Missing Font" not embedded/.test(m)));
});

test("without a resolver nothing is embedded", async () => {
  const z = await unzip(await writePptx([{ layout: "blank", shapes: [] }], DECK));
  assert.doesNotMatch(z.text("ppt/presentation.xml"), /embeddedFontLst|embedTrueTypeFonts/);
});

const INTER = "/usr/share/fonts/rsms-inter-fonts/Inter-Regular.ttf";
test("a real font: Inter, if installed", { skip: !existsSync(INTER) }, () => {
  const h = readEot(eotFromTtf(readFileSync(INTER)));
  assert.equal(h.names[0], "Inter");
  assert.equal(h.weight, 400);
});
