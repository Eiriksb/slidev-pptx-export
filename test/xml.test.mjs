import { test } from "node:test";
import assert from "node:assert/strict";
import { esc, emu, emuSize, pt100, guid, colorXml, solidFill } from "../src/ooxml/xml.mjs";

test("esc escapes markup characters", () => {
  assert.equal(esc(`a<b>&"c"'`), "a&lt;b&gt;&amp;&quot;c&quot;&apos;");
});

test("esc strips characters XML 1.0 forbids", () => {
  assert.equal(esc("a\u0001b\u0008c\u000Bd"), "abcd");
  assert.equal(esc("x\uFFFEy\uFFFFz"), "xyz");
  assert.equal(esc("lone\uD800surrogate"), "lonesurrogate");
  assert.equal(esc("keep\t\n\r"), "keep\t\n\r");
  assert.equal(esc("emoji 😀"), "emoji 😀");
  assert.equal(esc("\uFB01\uFF21\uFFFC\uFFFD"), "\uFB01\uFF21\uFFFC\uFFFD"); // valid BMP text above U+E000 is kept
});

test("esc handles null/undefined/numbers", () => {
  assert.equal(esc(undefined), "");
  assert.equal(esc(null), "");
  assert.equal(esc(42), "42");
});

test("emu converts inches and guards bad numbers", () => {
  assert.equal(emu(1), 914400);
  assert.equal(emu(0.5), 457200);
  assert.equal(emu(NaN), 0);
  assert.equal(emuSize(-1), 0);
  assert.equal(pt100(13), 1300);
});

test("guid is deterministic and well-formed", () => {
  const a = guid("table-default");
  assert.equal(a, guid("table-default"));
  assert.notEqual(a, guid("table-deflist"));
  assert.match(a, /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/);
});

test("colorXml writes PowerPoint's mapped scheme names", () => {
  assert.equal(colorXml("dk1"), '<a:schemeClr val="tx1"/>');
  assert.equal(colorXml("lt2"), '<a:schemeClr val="bg2"/>');
  assert.equal(colorXml("accent3"), '<a:schemeClr val="accent3"/>');
  assert.equal(colorXml("#1a833b"), '<a:srgbClr val="1A833B"/>');
  assert.equal(colorXml("dk2", 0.16), '<a:schemeClr val="tx2"><a:alpha val="16000"/></a:schemeClr>');
  assert.equal(solidFill("accent1"), '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>');
  assert.throws(() => colorXml("purple"), /Unknown color/);
});
