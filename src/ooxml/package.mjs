// OPC package writer. The only way to add a part is through addXml/addMedia, which also
// registers its content type, and [Content_Types].xml and every _rels/*.rels are generated
// from what was actually added - so a phantom Override (pptxgenjs#1449) or a relationship to a
// part that doesn't exist can't be written. Same calls in the same order give the same bytes.

import JSZip from "jszip";
import { XML_DECL, NS, esc } from "./xml.mjs";

const OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export const REL = {
  officeDocument: `${OFFICE_REL}/officeDocument`,
  coreProps: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
  extProps: `${OFFICE_REL}/extended-properties`,
  slideMaster: `${OFFICE_REL}/slideMaster`,
  slideLayout: `${OFFICE_REL}/slideLayout`,
  slide: `${OFFICE_REL}/slide`,
  theme: `${OFFICE_REL}/theme`,
  notesMaster: `${OFFICE_REL}/notesMaster`,
  notesSlide: `${OFFICE_REL}/notesSlide`,
  presProps: `${OFFICE_REL}/presProps`,
  viewProps: `${OFFICE_REL}/viewProps`,
  tableStyles: `${OFFICE_REL}/tableStyles`,
  image: `${OFFICE_REL}/image`,
  video: `${OFFICE_REL}/video`,
  media: "http://schemas.microsoft.com/office/2007/relationships/media",
  chart: `${OFFICE_REL}/chart`,
  package: `${OFFICE_REL}/package`,
  hyperlink: `${OFFICE_REL}/hyperlink`,
  worksheet: `${OFFICE_REL}/worksheet`,
  styles: `${OFFICE_REL}/styles`,
  sharedStrings: `${OFFICE_REL}/sharedStrings`,
};

const PML = "application/vnd.openxmlformats-officedocument.presentationml";
const SML = "application/vnd.openxmlformats-officedocument.spreadsheetml";

export const CT = {
  presentation: `${PML}.presentation.main+xml`,
  slideMaster: `${PML}.slideMaster+xml`,
  slideLayout: `${PML}.slideLayout+xml`,
  slide: `${PML}.slide+xml`,
  notesMaster: `${PML}.notesMaster+xml`,
  notesSlide: `${PML}.notesSlide+xml`,
  presProps: `${PML}.presProps+xml`,
  viewProps: `${PML}.viewProps+xml`,
  tableStyles: `${PML}.tableStyles+xml`,
  theme: "application/vnd.openxmlformats-officedocument.theme+xml",
  chart: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
  coreProps: "application/vnd.openxmlformats-package.core-properties+xml",
  extProps: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
  workbook: `${SML}.sheet.main+xml`,
  worksheet: `${SML}.worksheet+xml`,
  styles: `${SML}.styles+xml`,
  sharedStrings: `${SML}.sharedStrings+xml`,
};

const MEDIA_TYPES = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  svg: "image/svg+xml",
  mp4: "video/mp4",
  mov: "video/quicktime",
  xlsx: `${SML}.sheet`,
};

// DOS epoch: the earliest date a zip entry can carry. Fixed so output is reproducible.
const FIXED_DATE = new Date(Date.UTC(1980, 0, 1, 0, 0, 0));

function relsPartFor(sourcePart) {
  if (sourcePart === "") return "/_rels/.rels";
  const i = sourcePart.lastIndexOf("/");
  return `${sourcePart.slice(0, i)}/_rels/${sourcePart.slice(i + 1)}.rels`;
}

function relativeTarget(fromPart, toPart) {
  const fromDir = fromPart.split("/").slice(1, -1);
  const to = toPart.split("/").slice(1);
  let i = 0;
  while (i < fromDir.length && i < to.length - 1 && fromDir[i] === to[i]) i++;
  return "../".repeat(fromDir.length - i) + to.slice(i).join("/");
}

function assertPartName(name) {
  if (!/^\/[^\s]+[^/]$/.test(name)) throw new Error(`Invalid part name ${JSON.stringify(name)}`);
}

export class Package {
  constructor() {
    this.parts = new Map(); // partName -> { data, compress }
    this.overrides = new Map(); // partName -> content type
    this.defaults = new Map([
      ["rels", "application/vnd.openxmlformats-package.relationships+xml"],
      ["xml", "application/xml"],
    ]);
    this.rels = new Map(); // sourcePart ("" = package root) -> [{ id, type, target, external }]
  }

  addXml(partName, contentType, xml) {
    assertPartName(partName);
    if (this.parts.has(partName)) throw new Error(`Duplicate part ${partName}`);
    this.parts.set(partName, { data: xml, compress: true });
    this.overrides.set(partName, contentType);
  }

  addMedia(partName, buffer) {
    assertPartName(partName);
    if (this.parts.has(partName)) throw new Error(`Duplicate part ${partName}`);
    const ext = partName.split(".").pop().toLowerCase();
    const type = MEDIA_TYPES[ext];
    if (!type) throw new Error(`Unsupported media type for ${partName}`);
    // Already-compressed binaries are stored, not deflated again.
    this.parts.set(partName, { data: buffer, compress: false });
    this.defaults.set(ext, type);
  }

  has(partName) {
    return this.parts.has(partName);
  }

  // Returns the rId of a relationship from `fromPart` to `target`, reusing an identical one.
  rel(fromPart, target, type, { external = false } = {}) {
    if (!this.rels.has(fromPart)) this.rels.set(fromPart, []);
    const list = this.rels.get(fromPart);
    const found = list.find((r) => r.type === type && r.target === target && r.external === external);
    if (found) return found.id;
    const id = `rId${list.length + 1}`;
    list.push({ id, type, target, external });
    return id;
  }

  contentTypesXml() {
    const defaults = [...this.defaults]
      .map(([ext, type]) => `<Default Extension="${esc(ext)}" ContentType="${esc(type)}"/>`)
      .join("");
    const overrides = [...this.overrides]
      .map(([name, type]) => `<Override PartName="${esc(name)}" ContentType="${esc(type)}"/>`)
      .join("");
    return `${XML_DECL}<Types xmlns="${NS.ct}">${defaults}${overrides}</Types>`;
  }

  relsXml(sourcePart, list) {
    const items = list
      .map((r) => {
        const target = r.external ? r.target : relativeTarget(sourcePart, r.target);
        const mode = r.external ? ' TargetMode="External"' : "";
        return `<Relationship Id="${r.id}" Type="${esc(r.type)}" Target="${esc(target)}"${mode}/>`;
      })
      .join("");
    return `${XML_DECL}<Relationships xmlns="${NS.rels}">${items}</Relationships>`;
  }

  async toBuffer() {
    for (const [source, list] of this.rels) {
      // A rels source that was never added via addXml/addMedia would otherwise be silently
      // dropped below (the part loop only visits parts that exist) instead of writing the
      // .rels file a relationship to it needs.
      if (source !== "" && !this.parts.has(source)) {
        throw new Error(`Relationships source part ${source} was never added`);
      }
      for (const r of list) {
        if (!r.external && !this.parts.has(r.target)) {
          throw new Error(`Relationship ${r.id} from ${source || "/"} points at missing part ${r.target}`);
        }
      }
    }
    const zip = new JSZip();
    const add = (name, data, compress) =>
      zip.file(name.slice(1), data, { date: FIXED_DATE, createFolders: false, compression: compress ? "DEFLATE" : "STORE" });

    add("/[Content_Types].xml", this.contentTypesXml(), true);
    if (this.rels.has("")) add(relsPartFor(""), this.relsXml("", this.rels.get("")), true);
    for (const [name, part] of this.parts) {
      add(name, part.data, part.compress);
      const list = this.rels.get(name);
      if (list && list.length) add(relsPartFor(name), this.relsXml(name, list), true);
    }
    return zip.generateAsync({ type: "nodebuffer", compressionOptions: { level: 6 } });
  }
}
