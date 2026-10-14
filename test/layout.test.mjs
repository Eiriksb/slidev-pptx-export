import { test } from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";
import { join } from "node:path";
import { splitSlides, parseSlide } from "../src/parse.mjs";
import { mergeConfig } from "../src/config.mjs";
import { layoutDeck } from "../src/layout.mjs";

const FIXTURE_DIR = fileURLToPath(new URL("./fixtures/kitchen-sink/", import.meta.url));

function specsFor(md, userConfig = {}) {
  const config = mergeConfig(userConfig);
  config.inputDir = FIXTURE_DIR;
  config.publicDir = join(FIXTURE_DIR, "public");
  config.resolvedDataDir = null;
  const codeStore = [];
  const { slides } = splitSlides(md);
  const parsed = slides.map((s) => ({ ...parseSlide(s.content, codeStore, config), frontmatter: s.frontmatter }));
  return layoutDeck(parsed, codeStore, config);
}
const texts = (paras) => paras.map((p) => p.runs.map((r) => r.text).join(""));
const COVER = "# Cover\n\n";

test("first slide is a Title Slide with title and subtitle placeholders", () => {
  const [cover] = specsFor("# Deck\n\nThe **subtitle**\n");
  assert.equal(cover.layout, "title");
  assert.deepEqual(texts(cover.placeholders.title), ["Deck"]);
  assert.deepEqual(cover.placeholders.subtitle[0].runs, [
    { font: "body", text: "The " },
    { font: "body", text: "subtitle", b: true, color: "dk2" },
  ]);
});

test("text-only body goes into the content placeholder with native bullets", () => {
  const [, s] = specsFor(`${COVER}---\n\n<div class="eyebrow">Intro</div>\n\n# Plain\n\nPara with **bold**\n\n- one\n- two\n\n1. first\n2. second\n\n<!-- say this -->\n`);
  assert.equal(s.layout, "content");
  assert.deepEqual(texts(s.placeholders.title), ["Plain"]);
  assert.deepEqual(s.placeholders.body.map((p) => p.bullet), ["none", "char", "char", "number", "number"]);
  assert.equal(s.placeholders.body[0].runs[0].size, undefined); // inherits the master's 13pt
  assert.equal(s.placeholders.body[0].runs[0].color, undefined);
  assert.equal(s.placeholders.body[0].runs[1].color, "dk2");
  assert.equal(s.placeholders.body[1].spaceBefore, 12);
  assert.equal(s.placeholders.body[2].spaceBefore, undefined);
  assert.equal(s.placeholders.bodyFontScale, undefined);
  assert.equal(s.notes, "say this");
  const eyebrow = s.shapes[0];
  assert.deepEqual(eyebrow.paras[0].runs[0], { text: "INTRO", font: "heading", size: 9, b: true, color: "accent1", spacing: 1.5 });
});

test("F5: an ordered list starting above 1 carries startAt on only its first paragraph", () => {
  const [, s] = specsFor(`${COVER}---\n\n# Numbered\n\n3. third\n4. fourth\n`);
  assert.equal(s.placeholders.body[0].number, 3);
  assert.equal(s.placeholders.body[1].number, undefined);
});

test("an overflowing text-only body gets a shrink-on-overflow font scale", () => {
  const bullets = Array.from({ length: 40 }, (_, i) => `- bullet number ${i}`).join("\n");
  const [, s] = specsFor(`${COVER}---\n\n# Dense\n\n${bullets}\n`);
  assert.equal(s.layout, "content");
  assert.ok(s.placeholders.bodyFontScale >= 0.55 && s.placeholders.bodyFontScale < 1, String(s.placeholders.bodyFontScale));
});

test("a rich body uses Title Only plus positioned shapes", () => {
  const [, s] = specsFor(`${COVER}---\n\n# Table\n\n| a | b |\n| --- | --- |\n| 1 | 2 |\n`);
  assert.equal(s.layout, "titleOnly");
  const table = s.shapes.find((x) => x.kind === "table");
  assert.equal(table.header, true);
  assert.equal(table.rows.length, 2);
  assert.equal(table.y, 1.6);
});

test("F2: table header-row runs carry b:true so the header stays bold even without the table style", () => {
  const [, s] = specsFor(`${COVER}---\n\n# Table\n\n| a | b |\n| --- | --- |\n| 1 | 2 |\n`);
  const table = s.shapes.find((x) => x.kind === "table");
  assert.equal(table.rows[0][0][0].runs[0].b, true);
  assert.equal(table.rows[0][1][0].runs[0].b, true);
  assert.equal(table.rows[1][0][0].runs[0].b, undefined);
});

test("layout: center with text-center centers the stack and the text", () => {
  const [, s] = specsFor(`${COVER}---\nlayout: center\nclass: text-center\n---\n\n# Thanks\n\nQuestions?\n`);
  assert.equal(s.layout, "centered");
  assert.equal(s.placeholders.title[0].align, "ctr");
  assert.equal(s.placeholders.body[0].align, "ctr");
  assert.ok(s.placeholders.titleBox.y > 2.5, String(s.placeholders.titleBox.y));
  assert.ok(s.placeholders.bodyBox.y > s.placeholders.titleBox.y);
});

test("layout: section maps to Section Header", () => {
  const [, s] = specsFor(`${COVER}---\nlayout: section\n---\n\n# Part two\n\nWhat comes next\n`);
  assert.equal(s.layout, "section");
  assert.deepEqual(texts(s.placeholders.body), ["What comes next"]);
  assert.equal(s.placeholders.bodyFontScale, undefined);
});

test("F4: an overflowing section body gets a shrink-on-overflow font scale too", () => {
  const bullets = Array.from({ length: 10 }, (_, i) => `- bullet number ${i}`).join("\n");
  const [, s] = specsFor(`${COVER}---\nlayout: section\n---\n\n# Part two\n\n${bullets}\n`);
  assert.equal(s.layout, "section");
  assert.ok(s.placeholders.bodyFontScale >= 0.55 && s.placeholders.bodyFontScale < 1, String(s.placeholders.bodyFontScale));
});

test("a slide without a title is Blank with everything as shapes", () => {
  const [, s] = specsFor(`${COVER}---\n\nJust text\n`);
  assert.equal(s.layout, "blank");
  assert.equal(s.shapes[0].kind, "text");
});

test("charts: resolver colors map to theme slots, unresolved tags become placeholders", () => {
  const C = () => ({ kind: "barchart", labels: ["a", "b"], datasets: [{ name: "x", color: "2563EB", data: [1, 2] }, { name: "y", data: [3, 4] }], min: 0, max: 10 });
  const [, s] = specsFor(`${COVER}---\n\n# Charts\n\n<C />\n\n<Missing />\n`, { chartResolvers: { C } });
  const chart = s.shapes.find((x) => x.kind === "chart").chart;
  assert.equal(chart.type, "column");
  assert.deepEqual(chart.series.map((x) => x.color), ["accent1", "accent2"]);
  const missing = s.shapes.find((x) => x.kind === "text" && x.line && x.line.dash === "dash");
  assert.match(missing.paras[0].runs[0].text, /Missing — interactive chart, no data directory configured/);
});

test("F6: a malformed resolver color falls back to the palette instead of throwing, and warns once", () => {
  const warnings = [];
  const orig = console.warn;
  console.warn = (...args) => warnings.push(args.join(" "));
  try {
    const C = () => ({
      kind: "barchart", labels: ["a", "b"],
      datasets: [{ name: "x", color: "green", data: [1, 2] }, { name: "y", color: "green", data: [3, 4] }],
      min: 0, max: 10,
    });
    const [, s] = specsFor(`${COVER}---\n\n# Charts\n\n<C />\n`, { chartResolvers: { C } });
    const chart = s.shapes.find((x) => x.kind === "chart").chart;
    assert.deepEqual(chart.series.map((x) => x.color), ["accent1", "accent2"]);
    assert.equal(warnings.length, 1, warnings.join("\n"));
    assert.match(warnings[0], /green/);
  } finally {
    console.warn = orig;
  }
});

test("images are contain-fitted with alt text; missing ones become placeholders", () => {
  const [, s] = specsFor(`${COVER}---\n\n# Pics\n\n![Alt text](/pic.png)\n\n![nope](/missing.png)\n`);
  const img = s.shapes.find((x) => x.kind === "image");
  assert.equal(img.alt, "Alt text");
  assert.ok(Math.abs(img.w / img.h - 2) < 1e-9);
  assert.ok(s.shapes.some((x) => x.kind === "text" && /image not found: \/missing.png/.test(x.paras?.[0]?.runs?.[0]?.text || "")));
});

test("opacity fades become blended hex colors", () => {
  const [, s] = specsFor(`${COVER}---\n\n# Fade\n\n<div class="opacity-50">Faded</div>\n`);
  assert.equal(s.placeholders.body[0].runs[0].color, "#8F949B");
});

test("card rows become rounded panels with their own content inside", () => {
  const [, s] = specsFor(`${COVER}---\n\n# Cards\n\n<div class="grid grid-cols-2 gap-4">\n<div>\n\n### Left\n\nText\n\n</div>\n<div>\n\n### Right\n\nMore\n\n</div>\n</div>\n`);
  assert.equal(s.layout, "titleOnly");
  assert.equal(s.shapes.filter((x) => x.geometry === "roundRect").length, 2);
});
