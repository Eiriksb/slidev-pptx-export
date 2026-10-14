import { test } from "node:test";
import assert from "node:assert/strict";
import { parseSlide, mdInlineTokens } from "../src/parse.mjs";
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

test("F5: an ordered list starting above 1 keeps its start number", () => {
  const s = parse("# T\n\n3. third\n4. fourth\n");
  assert.deepEqual(s.blocks[0], { type: "bullets", ordered: true, start: 3, items: ["third", "fourth"] });
});

test("F5: a default (start-1) ordered list and an unordered list carry no surprising start", () => {
  const ordered = parse("# T\n\n1. first\n2. second\n");
  assert.equal(ordered.blocks[0].start, 1);
  const unordered = parse("# T\n\n- a\n- b\n");
  assert.equal(unordered.blocks[0].start, undefined);
});

test("F8: notes containing a heading don't get mistaken for the slide title", () => {
  const s = parse("Body text\n\n<!--\n# Not a title\nactual notes\n-->\n");
  assert.equal(s.title, null);
  assert.equal(s.notes, "# Not a title\nactual notes");
});
