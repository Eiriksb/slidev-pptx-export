// Package lint: structural checks for the things that make PowerPoint (especially Online)
// show a repair prompt. Not a schema validator - it checks the invariants we've been bitten
// by or that DeckCraft's writer guarantees by construction.

import JSZip from "jszip";
import { parseXml, elements } from "./helpers.mjs";

const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

function sourceOfRels(relsName) {
  // "ppt/slides/_rels/slide1.xml.rels" -> "ppt/slides/slide1.xml"; "_rels/.rels" -> ""
  const m = /^(.*?)_rels\/(.*)\.rels$/.exec(relsName);
  return m ? `${m[1]}${m[2]}` : null;
}

function directChild(parent, localName) {
  for (let c = parent.firstChild; c; c = c.nextSibling) {
    if (c.nodeType === 1 && c.localName === localName) return c;
  }
  return null;
}

function resolveTarget(sourcePart, target) {
  if (target.startsWith("/")) return target.slice(1);
  const parts = sourcePart.split("/").slice(0, -1);
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== ".") parts.push(seg);
  }
  return parts.join("/");
}

export async function lintPackage(buffer, label = "") {
  const problems = [];
  const report = (msg) => problems.push(label ? `${label}: ${msg}` : msg);
  const zip = await JSZip.loadAsync(buffer);
  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
  const nameSet = new Set(names);

  const docs = new Map();
  for (const n of names.filter((n) => /\.(xml|rels)$/.test(n))) {
    try {
      docs.set(n, parseXml(await zip.file(n).async("string"), n));
    } catch (err) {
      report(err.message);
    }
  }

  // Content types.
  const ct = docs.get("[Content_Types].xml");
  if (!ct) {
    report("missing or unparseable [Content_Types].xml");
  } else {
    const defaults = new Set();
    const overrides = new Set();
    for (const el of elements(ct)) {
      if (el.localName === "Default") defaults.add(el.getAttribute("Extension").toLowerCase());
      if (el.localName === "Override") {
        const part = el.getAttribute("PartName").slice(1);
        overrides.add(part);
        if (!nameSet.has(part)) report(`Override for missing part /${part}`);
      }
    }
    for (const n of names) {
      if (n === "[Content_Types].xml") continue;
      const ext = n.split(".").pop().toLowerCase();
      if (!overrides.has(n) && !defaults.has(ext)) report(`no content type for /${n}`);
    }
  }

  // Relationships.
  const relIds = new Map(); // source part -> Set(ids)
  for (const [n, doc] of docs) {
    if (!n.endsWith(".rels")) continue;
    const source = sourceOfRels(n);
    const ids = new Set();
    for (const el of elements(doc)) {
      if (el.localName !== "Relationship") continue;
      const id = el.getAttribute("Id");
      if (ids.has(id)) report(`${n}: duplicate relationship id ${id}`);
      ids.add(id);
      if (el.getAttribute("TargetMode") === "External") continue;
      const target = resolveTarget(source, el.getAttribute("Target"));
      if (!nameSet.has(target)) report(`${n}: ${id} targets missing part /${target}`);
    }
    relIds.set(source, ids);
  }

  // Per-part checks.
  const masterLayoutIds = new Set();
  const sldIds = new Set();
  for (const [n, doc] of docs) {
    if (n.endsWith(".rels") || n === "[Content_Types].xml") continue;
    const cNvPrIds = new Set();
    for (const el of elements(doc)) {
      for (let i = 0; i < el.attributes.length; i++) {
        const attr = el.attributes[i];
        if (attr.namespaceURI === R_NS && attr.value !== "") {
          if (!(relIds.get(n) || new Set()).has(attr.value)) report(`${n}: ${attr.name}="${attr.value}" not defined in its rels`);
        }
      }
      if (el.localName === "cNvPr") {
        const id = el.getAttribute("id");
        if (cNvPrIds.has(id)) report(`${n}: duplicate cNvPr id ${id}`);
        cNvPrIds.add(id);
      }
      if (el.localName === "p" && el.namespaceURI === A_NS) {
        const kids = [...elements(el)].filter((k) => k.parentNode === el);
        const pPrs = kids.filter((k) => k.localName === "pPr");
        if (pPrs.length > 1) report(`${n}: a:p with more than one a:pPr`);
        if (pPrs.length === 1 && kids[0] !== pPrs[0]) report(`${n}: a:pPr is not the first child of a:p`);
      }
      if (["txBody", "txPr", "rich"].includes(el.localName)) {
        const hasP = [...elements(el)].some((k) => k.parentNode === el && k.localName === "p");
        if (!hasP) report(`${n}: ${el.localName === "txBody" ? "txBody" : el.localName} without a:p`);
      }
      if (el.localName === "barChart") {
        const axIds = [...elements(el)].filter((k) => k.parentNode === el && k.localName === "axId").length;
        if (axIds !== 2) report(`${n}: barChart has ${axIds} c:axId (expected 2)`);
      }
      if (el.localName === "pt" && (el.parentNode.localName === "numCache" || el.parentNode.localName === "strCache")) {
        const ptCountEl = directChild(el.parentNode, "ptCount");
        const ptCount = ptCountEl ? Number(ptCountEl.getAttribute("val")) : NaN;
        const idx = Number(el.getAttribute("idx"));
        if (!(idx < ptCount)) report(`${n}: c:pt idx=${idx} >= ptCount ${ptCount}`);
      }
      if (el.localName === "sldId") {
        const id = Number(el.getAttribute("id"));
        if (!(id >= 256)) report(`${n}: sldId ${id} < 256`);
        if (sldIds.has(id)) report(`${n}: sldId ${id} used twice`);
        sldIds.add(id);
      }
      if (el.localName === "sldMasterId" || el.localName === "sldLayoutId") {
        const id = Number(el.getAttribute("id"));
        if (!(id >= 2147483648)) report(`${n}: ${el.localName} ${id} < 2147483648`);
        if (masterLayoutIds.has(id)) report(`${n}: master/layout id ${id} used twice`);
        masterLayoutIds.add(id);
      }
    }
  }

  // Embedded workbooks are packages too.
  for (const n of names.filter((n) => n.endsWith(".xlsx"))) {
    problems.push(...(await lintPackage(await zip.file(n).async("nodebuffer"), `${label ? `${label} > ` : ""}${n}`)));
  }

  return problems;
}
