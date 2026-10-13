import { test } from "node:test";
import assert from "node:assert/strict";
import { normalizeHex, themeColors, toColor, hexOf, blendWithWhite } from "../src/colors.mjs";
import { themeXml } from "../src/ooxml/theme.mjs";
import { parseXml, byLocal } from "./helpers.mjs";

const COLORS = { accent: "1A833B", accentDark: "14532D", accent2: "1A589F", panel: "EEF4F0", border: "D9E3DB", text: "1F2937", muted: "6B7280" };

test("normalizeHex accepts #, lowercase and 3-digit forms", () => {
  assert.equal(normalizeHex("#1a833b"), "1A833B");
  assert.equal(normalizeHex("fff"), "FFFFFF");
  assert.throws(() => normalizeHex("green"), /Invalid color/);
});

test("themeColors maps config colors onto the 12 slots", () => {
  assert.deepEqual(themeColors(COLORS), {
    dk1: "1F2937", lt1: "FFFFFF", dk2: "14532D", lt2: "EEF4F0",
    accent1: "1A833B", accent2: "1A589F", accent3: "6B7280", accent4: "D9E3DB",
    accent5: "9CA3AF", accent6: "F59E0B", hlink: "1A833B", folHlink: "14532D",
  });
});

test("toColor resolves theme colors to slots and others to hex", () => {
  const theme = themeColors(COLORS);
  assert.equal(toColor("1F2937", theme), "dk1");
  assert.equal(toColor("#1a833b", theme), "accent1");
  assert.equal(toColor("FFFFFF", theme), "lt1");
  assert.equal(toColor("9CA3AF", theme), "accent5");
  assert.equal(toColor("123456", theme), "#123456");
  assert.equal(hexOf("dk2", theme), "14532D");
  assert.equal(hexOf("#abcdef", theme), "ABCDEF");
});

test("blendWithWhite reproduces CSS opacity over white", () => {
  assert.equal(blendWithWhite("000000", 0.5), "808080");
  assert.equal(blendWithWhite("1F2937", 1), "1F2937");
});

test("themeXml writes slots in schema order, fonts and 3-entry style lists", () => {
  const xml = themeXml({ name: "Deck", colors: themeColors(COLORS), fonts: { heading: "Plus Jakarta Sans", body: "Inter" } });
  const doc = parseXml(xml);
  const scheme = byLocal(doc, "clrScheme")[0];
  const slots = [];
  for (let c = scheme.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) slots.push(c.localName);
  assert.deepEqual(slots, ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]);
  assert.match(xml, /<a:majorFont><a:latin typeface="Plus Jakarta Sans"\/>/);
  assert.match(xml, /<a:minorFont><a:latin typeface="Inter"\/>/);
  for (const list of ["fillStyleLst", "lnStyleLst", "effectStyleLst", "bgFillStyleLst"]) {
    const el = byLocal(doc, list)[0];
    let count = 0;
    for (let c = el.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) count++;
    assert.equal(count, 3, list);
  }
});
