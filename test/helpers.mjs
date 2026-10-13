// Small helpers for asserting on generated packages: unzip into memory and parse parts.
import JSZip from "jszip";
import { DOMParser } from "@xmldom/xmldom";

export function parseXml(text, name = "part") {
  const errors = [];
  let doc;
  try {
    doc = new DOMParser({
      onError: (level, msg) => {
        if (level !== "warning") errors.push(msg);
      },
    }).parseFromString(text, "text/xml");
  } catch (err) {
    errors.push(err.message);
  }
  if (errors.length) throw new Error(`${name} is not well-formed XML: ${errors[0]}`);
  return doc;
}

export async function unzip(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
  const texts = new Map();
  const buffers = new Map();
  for (const n of names) {
    buffers.set(n, await zip.file(n).async("nodebuffer"));
    if (/\.(xml|rels)$/.test(n)) texts.set(n, buffers.get(n).toString("utf8"));
  }
  return {
    names,
    text: (n) => {
      if (!texts.has(n)) throw new Error(`No XML part ${n}`);
      return texts.get(n);
    },
    xml: (n) => parseXml(texts.get(n) ?? (() => { throw new Error(`No XML part ${n}`); })(), n),
    buffer: (n) => buffers.get(n),
  };
}

export function* elements(node) {
  for (let c = node.firstChild; c; c = c.nextSibling) {
    if (c.nodeType === 1) {
      yield c;
      yield* elements(c);
    }
  }
}

export function byLocal(node, localName) {
  return [...elements(node)].filter((e) => e.localName === localName);
}
