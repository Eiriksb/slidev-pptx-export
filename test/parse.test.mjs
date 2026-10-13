import { test } from "node:test";
import assert from "node:assert/strict";
import { parseSlide, mdInlineTokens, mdInlineToRuns } from "../src/parse.mjs";
import { mergeConfig } from "../src/config.mjs";

const config = mergeConfig({});
const parse = (md) => parseSlide(md, [], config);

test("mdInlineTokens splits bold, italic and code", () => {
  assert.deepEqual(mdInlineTokens("a **b** *c* `d` e"), [
    { text: "a " }, { text: "b", bold: true }, { text: " " }, { text: "c", italic: true },
    { text: " " }, { text: "d", code: true }, { text: " e" },
  ]);
  assert.deepEqual(mdInlineTokens(""), [{ text: "" }]);
});

test("mdInlineToRuns still produces the old pptxgenjs run shape", () => {
  const runs = mdInlineToRuns("x **y**", { fontSize: 13 }, config);
  assert.deepEqual(runs.map((r) => [r.text, r.options.bold === true]), [["x ", false], ["y", true]]);
  assert.equal(runs[1].options.color, config.colors.accentDark);
});

test("the last trailing HTML comment becomes speaker notes", () => {
  const s = parse("# Title\n\nBody text\n\n<!-- An earlier comment -->\n\nMore\n\n<!--\nLine one\n\nLine two\n-->\n");
  assert.equal(s.notes, "Line one\n\nLine two");
  assert.equal(s.title, "Title");
  assert.deepEqual(s.blocks.map((b) => b.text), ["Body text", "More"]);
});

test("a comment that is followed by content is not notes", () => {
  assert.equal(parse("# T\n\n<!-- not notes -->\n\nBody").notes, null);
});

test("a comment inside a code fence is not notes", () => {
  assert.equal(parse("# T\n\n```html\n<!-- in code -->\n```\n").notes, null);
});
