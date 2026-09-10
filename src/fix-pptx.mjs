// pptxgenjs@4.0.1's makeXmlContTypes() writes one "<Override .../ppt/slideMasters/
// slideMasterN.xml" entry per *slide* (idx + 1) instead of per actual slide master, even
// though only slideMaster1.xml is ever written to the zip (see gitbrent/PptxGenJS#1449).
// A multi-slide deck therefore ships a [Content_Types].xml that declares parts
// (slideMaster2.xml, slideMaster3.xml, ...) which don't exist in the package. Desktop
// PowerPoint silently tolerates/self-heals this; PowerPoint Online and other stricter
// OOXML readers don't and show a "needs repair" prompt on every open.
//
// Since this is upstream's bug, not something a config or layout change here can avoid,
// patch the zip after pptxgenjs writes it: drop any slideMaster Override whose target file
// isn't actually present in the package.

import JSZip from "jszip";
import { readFile, writeFile } from "node:fs/promises";

const CONTENT_TYPES_PATH = "[Content_Types].xml";
const PHANTOM_OVERRIDE_RE = /<Override PartName="\/ppt\/slideMasters\/(slideMaster\d+\.xml)"[^>]*\/>/g;

export async function stripPhantomSlideMasterEntries(pptxPath) {
  const zip = await JSZip.loadAsync(await readFile(pptxPath));

  const contentTypesFile = zip.file(CONTENT_TYPES_PATH);
  if (!contentTypesFile) return;
  let contentTypes = await contentTypesFile.async("string");

  const before = contentTypes;
  contentTypes = contentTypes.replace(PHANTOM_OVERRIDE_RE, (match, fileName) =>
    zip.file(`ppt/slideMasters/${fileName}`) ? match : ""
  );
  if (contentTypes === before) return;

  zip.file(CONTENT_TYPES_PATH, contentTypes);
  const buffer = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  await writeFile(pptxPath, buffer);
}
