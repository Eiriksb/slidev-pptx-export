import { test } from "node:test";
import assert from "node:assert/strict";
import { notesParas } from "../src/ooxml/notes.mjs";
import { writePptx } from "../src/ooxml/index.mjs";
import { DECK } from "./deck.mjs";
import { lintPackage } from "./lint.mjs";
import { unzip } from "./helpers.mjs";

test("notesParas: one paragraph per line, blank runs collapse", () => {
  assert.deepEqual(notesParas("\n  first  \nsecond\n\n\n- third\n\n"), [
    { runs: [{ text: "first" }] },
    { runs: [{ text: "second" }] },
    { runs: [] },
    { runs: [{ text: "- third" }] },
  ]);
});

test("slides with notes get notes slides wired to the notes master", async () => {
  const slides = [
    { layout: "blank", shapes: [], notes: "Say hello <warmly> & smile" },
    { layout: "blank", shapes: [] },
  ];
  const buffer = await writePptx(slides, DECK);
  assert.deepEqual(await lintPackage(buffer), []);
  const z = await unzip(buffer);
  assert.ok(z.names.includes("ppt/notesSlides/notesSlide1.xml"));
  assert.ok(!z.names.some((n) => n.startsWith("ppt/notesSlides/notesSlide2")));
  assert.match(z.text("ppt/notesSlides/notesSlide1.xml"), /<p:ph type="body" idx="1"\/>.*<a:t>Say hello &lt;warmly&gt; &amp; smile<\/a:t>/);
  assert.match(z.text("ppt/notesSlides/_rels/notesSlide1.xml.rels"), /\/notesMaster" Target="..\/notesMasters\/notesMaster1.xml"/);
  assert.match(z.text("ppt/notesSlides/_rels/notesSlide1.xml.rels"), /\/slide" Target="..\/slides\/slide1.xml"/);
  assert.match(z.text("ppt/slides/_rels/slide1.xml.rels"), /\/notesSlide" Target="..\/notesSlides\/notesSlide1.xml"/);
  assert.match(z.text("ppt/notesMasters/_rels/notesMaster1.xml.rels"), /\/theme" Target="..\/theme\/theme2.xml"/);
  assert.match(z.text("ppt/presentation.xml"), /<\/p:sldMasterIdLst><p:notesMasterIdLst><p:notesMasterId r:id="rId\d+"\/><\/p:notesMasterIdLst><p:sldIdLst>/);
  assert.match(z.text("docProps/app.xml"), /<Notes>1<\/Notes>/);
});

test("decks without notes have no notes master", async () => {
  const z = await unzip(await writePptx([{ layout: "blank", shapes: [] }], DECK));
  assert.ok(!z.names.some((n) => n.includes("notesMaster")));
  assert.doesNotMatch(z.text("ppt/presentation.xml"), /notesMasterIdLst/);
});
