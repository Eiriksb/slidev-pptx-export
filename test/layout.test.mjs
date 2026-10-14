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
  const bullets = Array.from({ length: 30 }, (_, i) => `- bullet number ${i}`).join("\n");
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

// A 25-line code block: too tall for the body, so the free shapes are scaled to fill it exactly.
const TALL_CODE = "```\n" + Array.from({ length: 25 }, (_, i) => `line ${i}`).join("\n") + "\n```\n";
const bottomOf = (shapes) => Math.max(...shapes.map((x) => x.y + x.h));

test("with a footer logo, positioned shapes end 0.1in above the logo", () => {
  const md = `${COVER}---\n\n# Code\n\n${TALL_CODE}`;
  const [, plain] = specsFor(md);
  assert.ok(bottomOf(plain.shapes) > 6.84, "without a logo the body runs past the logo's top");
  const [, s] = specsFor(md, { logo: { path: "logo.png", height: 0.27, aspect: 2 } });
  assert.ok(bottomOf(s.shapes) <= 6.84 + 1e-9, `bottom ${bottomOf(s.shapes)}`);
});

const bigTable = (n) => "| Name | Value |\n| --- | --- |\n" + Array.from({ length: n }, (_, i) => `| row ${i} | ${i} |`).join("\n") + "\n";
const cellText = (row) => row[0][0].runs.map((r) => r.text).join("");
const hasEyebrow = (spec, text) => spec.shapes.some((x) => x.kind === "text" && x.paras[0]?.runs[0]?.text === text);

test("a table too tall for its slide continues on extra slides with the header repeated", () => {
  const specs = specsFor(`${COVER}---\n\n<div class="eyebrow">Data</div>\n\n# Big table\n\n${bigTable(40)}\nAfter the table.\n\n<!-- table notes -->\n`);
  const pages = specs.slice(1);
  assert.equal(pages.length, 3);
  assert.deepEqual(pages.map((p) => texts(p.placeholders.title)[0]), ["Big table", "Big table (cont.)", "Big table (cont.)"]);
  assert.deepEqual(pages.map((p) => p.notes ?? null), ["table notes", null, null]);
  assert.ok(pages.every((p) => p.layout === "titleOnly" && hasEyebrow(p, "DATA")));
  const tables = pages.map((p) => p.shapes.find((x) => x.kind === "table"));
  for (const t of tables) {
    assert.equal(cellText(t.rows[0]), "Name");
    assert.ok(t.y + t.h <= 7.5 - 0.35 + 1e-9, `table ends at ${t.y + t.h}`);
  }
  assert.deepEqual(tables.flatMap((t) => t.rows.slice(1).map(cellText)), Array.from({ length: 40 }, (_, i) => `row ${i}`));
  // Content after the table follows the last chunk.
  const last = pages[2].shapes.filter((x) => x.kind === "text").map((x) => texts(x.paras).join(""));
  assert.ok(last.includes("After the table."));
  assert.ok(!pages[0].shapes.some((x) => x.kind === "text" && texts(x.paras).join("") === "After the table."));
});

test("a short table that fits stays on one slide", () => {
  const specs = specsFor(`${COVER}---\n\n# Small table\n\n${bigTable(8)}`);
  assert.equal(specs.length, 2);
  assert.equal(specs[1].shapes.find((x) => x.kind === "table").rows.length, 9);
});

test("a table that would start with fewer than 2 data rows starts on the next slide instead", () => {
  // A 19-line code block leaves room for the header row and at most one data row.
  const code = "```\n" + Array.from({ length: 19 }, (_, i) => `line ${i}`).join("\n") + "\n```\n";
  const specs = specsFor(`${COVER}---\n\n# Code then table\n\n${code}\n${bigTable(5)}`);
  assert.equal(specs.length, 3);
  assert.ok(!specs[1].shapes.some((x) => x.kind === "table"));
  assert.deepEqual(texts(specs[2].placeholders.title), ["Code then table (cont.)"]);
  assert.equal(specs[2].shapes.find((x) => x.kind === "table").rows.length, 6);
});

function captureWarnings(fn) {
  const warnings = [];
  const orig = console.warn;
  console.warn = (...args) => warnings.push(args.join(" "));
  try {
    fn();
  } finally {
    console.warn = orig;
  }
  return warnings;
}
const manyBullets = Array.from({ length: 60 }, (_, i) => `- point ${i}`).join("\n") + "\n";
const longCode = "```\n" + Array.from({ length: 60 }, (_, i) => `line ${i}`).join("\n") + "\n```\n";

test("slides still too full after shrinking to 55% print a warning naming the slide", () => {
  const cases = [
    ["content placeholder", `# Packed\n\n${manyBullets}`],
    ["section placeholder", `---\nlayout: section\n---\n\n# Packed\n\n${manyBullets}`],
    ["centered placeholder", `---\nlayout: center\n---\n\n# Packed\n\n${manyBullets}`],
    ["title only shapes", `# Packed\n\n${longCode}`],
    ["centered shapes", `---\nlayout: center\n---\n\n# Packed\n\n${longCode}`],
    ["blank shapes", longCode],
  ];
  for (const [label, slide] of cases) {
    const md = slide.startsWith("---") ? `${COVER}${slide}` : `${COVER}---\n\n${slide}`;
    const warnings = captureWarnings(() => specsFor(md));
    assert.equal(warnings.length, 1, `${label}: ${JSON.stringify(warnings)}`);
    const title = label === "blank shapes" ? "" : ' "Packed"';
    assert.equal(warnings[0], `Warning: slide 2${title} is still too full after shrinking to 55% - content may run off the slide.`, label);
  }
});

test("slides that fit, and long tables split across slides, don't warn", () => {
  const warnings = captureWarnings(() => specsFor(`${COVER}---\n\n# Fine\n\n- one\n- two\n\n---\n\n# Big table\n\n${bigTable(40)}`));
  assert.deepEqual(warnings, []);
});

test("slides with a split table are never shrunk, so nothing overlaps the table's rows", () => {
  // 30 rows: the second chunk nearly fills its slide, leaving no room for the closing line.
  const specs = specsFor(`${COVER}---\n\n# Thirty rows\n\n${bigTable(30)}\nClosing line.\n`);
  const pages = specs.slice(1);
  const tables = pages.flatMap((p) => p.shapes.filter((x) => x.kind === "table"));
  assert.deepEqual(tables.flatMap((t) => t.rows.slice(1).map(cellText)), Array.from({ length: 30 }, (_, i) => `row ${i}`));
  for (const t of tables) assert.ok(t.h / t.rows.length >= 0.32, `rows squeezed to ${t.h / t.rows.length}in`);
  for (const p of pages) {
    const shapes = [...p.shapes].sort((a, b) => a.y - b.y);
    for (let i = 1; i < shapes.length; i++) assert.ok(shapes[i].y >= shapes[i - 1].y + shapes[i - 1].h - 1e-9, "shapes overlap");
    assert.ok(bottomOf(p.shapes) <= 7.5 - 0.35 + 1e-9);
  }
  const lastTablePage = pages.findLastIndex((p) => p.shapes.some((x) => x.kind === "table"));
  const closingPage = pages.findIndex((p) => p.shapes.some((x) => x.kind === "text" && texts(x.paras).join("") === "Closing line."));
  assert.ok(closingPage >= lastTablePage);
});

test("warnings name the exported slide number, counting continuation slides before it", () => {
  const warnings = captureWarnings(() => specsFor(`${COVER}---\n\n# Big table\n\n${bigTable(40)}\n---\n\n# Packed\n\n${manyBullets}`));
  assert.equal(warnings.length, 1, JSON.stringify(warnings));
  assert.match(warnings[0], /^Warning: slide 5 "Packed"/);
});

test("rows whose cell text wraps get taller rows, and pagination plans with those heights", () => {
  const long = "a description long enough to wrap onto several lines in a third-width column cell";
  const md = "| Name | Description | Value |\n| --- | --- | --- |\n" + Array.from({ length: 20 }, (_, i) => `| row ${i} | ${long} | ${i} |`).join("\n") + "\n";
  const pages = specsFor(`${COVER}---\n\n# Wrapped\n\n${md}`).slice(1);
  const tables = pages.map((p) => p.shapes.find((x) => x.kind === "table"));
  assert.ok(pages.length > 2, `${pages.length} pages`);
  assert.deepEqual(tables.flatMap((t) => t.rows.slice(1).map(cellText)), Array.from({ length: 20 }, (_, i) => `row ${i}`));
  for (const t of tables) {
    assert.equal(t.rowH.length, t.rows.length);
    assert.ok(t.rowH[0] < t.rowH[1], "the short header row stays shorter than a wrapped row");
    assert.ok(t.y + t.rowH.reduce((a, b) => a + b, 0) <= 7.5 - 0.35 + 1e-9, "table runs past the body");
  }
});
