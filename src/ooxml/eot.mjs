// Fonts embedded in a .pptx (/ppt/fonts/*.fntdata) are Embedded OpenType: an EOT header (the
// font's names, PANOSE, weight, Unicode and code page ranges, embedding rights) in front of the
// TrueType file, uncompressed (version 0x00020001, no MTX compression, no XOR). This reads what the
// header needs from the font's own tables.

const EOT_VERSION = 0x00020001;
const MAGIC = 0x504c;

// sfnt table directory: tag -> { offset, length }.
export function sfntTables(buf) {
  const tag = buf.toString("latin1", 0, 4);
  if (tag !== "\0\x01\0\0" && tag !== "true") {
    throw new Error(tag === "OTTO" ? "CFF (PostScript) outlines: PowerPoint embeds TrueType fonts only" : "not a TrueType font");
  }
  const n = buf.readUInt16BE(4);
  const tables = {};
  for (let i = 0; i < n; i++) {
    const o = 12 + 16 * i;
    tables[buf.toString("latin1", o, o + 4)] = { offset: buf.readUInt32BE(o + 8), length: buf.readUInt32BE(o + 12) };
  }
  for (const t of ["name", "OS/2", "head", "glyf"]) if (!tables[t]) throw new Error(`no ${t} table`);
  return tables;
}

// A name (nameID) from the name table: the Windows Unicode English record, else any Windows
// Unicode one, as UTF-16LE bytes (what the EOT header holds).
function nameUtf16le(buf, table, nameId) {
  const base = table.offset;
  const count = buf.readUInt16BE(base + 2);
  const strings = base + buf.readUInt16BE(base + 4);
  let found = null;
  for (let i = 0; i < count; i++) {
    const r = base + 6 + 12 * i;
    const [platform, encoding, language, id] = [0, 2, 4, 6].map((k) => buf.readUInt16BE(r + k));
    if (id !== nameId || platform !== 3 || (encoding !== 1 && encoding !== 10)) continue;
    const rec = { length: buf.readUInt16BE(r + 8), offset: buf.readUInt16BE(r + 10) };
    if (!found || language === 0x409) found = rec;
    if (language === 0x409) break;
  }
  if (!found) return Buffer.alloc(0);
  const be = buf.subarray(strings + found.offset, strings + found.offset + found.length);
  const le = Buffer.alloc(be.length);
  for (let i = 0; i + 1 < be.length; i += 2) {
    le[i] = be[i + 1];
    le[i + 1] = be[i];
  }
  return le;
}

// What the font says about itself: family and style names, weight, italic, embedding rights.
export function fontInfo(buf) {
  const t = sfntTables(buf);
  const os2 = t["OS/2"].offset;
  const decode = (id) => nameUtf16le(buf, t.name, id).toString("utf16le");
  return {
    family: decode(16) || decode(1),
    style: decode(17) || decode(2),
    weight: buf.readUInt16BE(os2 + 4),
    fsType: buf.readUInt16BE(os2 + 8),
    italic: (buf.readUInt16BE(os2 + 62) & 1) === 1,
  };
}

// fsType bits: 0x0002 restricted license embedding (may not be embedded at all); 0x0200 bitmap
// embedding only. Everything else (installable, preview & print, editable) may be embedded.
export function embeddable(fsType) {
  return (fsType & 0x000f) !== 0x0002 && !(fsType & 0x0200);
}

export function eotFromTtf(buf) {
  const t = sfntTables(buf);
  const os2 = t["OS/2"].offset;
  const os2Version = buf.readUInt16BE(os2);
  const fsType = buf.readUInt16BE(os2 + 8);
  if (!embeddable(fsType)) throw new Error(`the font's license does not allow embedding (fsType ${fsType})`);
  const names = [1, 2, 5, 4].map((id) => nameUtf16le(buf, t.name, id)); // family, style, version, full name
  const fixed = Buffer.alloc(80); // up to Reserved4; each name below brings its padding word
  let o = 8; // EOTSize and FontDataSize are written last
  fixed.writeUInt32LE(EOT_VERSION, o); o += 4;
  fixed.writeUInt32LE(0, o); o += 4; // Flags: not subsetted, not compressed, not XORed
  buf.copy(fixed, o, os2 + 32, os2 + 42); o += 10; // PANOSE
  fixed.writeUInt8(1, o); o += 1; // Charset: DEFAULT_CHARSET
  fixed.writeUInt8(buf.readUInt16BE(os2 + 62) & 1, o); o += 1; // Italic: fsSelection bit 0
  fixed.writeUInt32LE(buf.readUInt16BE(os2 + 4), o); o += 4; // Weight: usWeightClass
  fixed.writeUInt16LE(fsType, o); o += 2;
  fixed.writeUInt16LE(MAGIC, o); o += 2;
  for (let i = 0; i < 4; i++, o += 4) fixed.writeUInt32LE(buf.readUInt32BE(os2 + 42 + 4 * i), o); // ulUnicodeRange1-4
  for (let i = 0; i < 2; i++, o += 4) fixed.writeUInt32LE(os2Version >= 1 ? buf.readUInt32BE(os2 + 78 + 4 * i) : 0, o); // ulCodePageRange1-2
  fixed.writeUInt32LE(buf.readUInt32BE(t.head.offset + 8), o); o += 4; // head.checkSumAdjustment
  o += 16; // Reserved1-4
  const parts = [fixed];
  for (const name of names) {
    const head = Buffer.alloc(4); // Padding (0), NameSize
    head.writeUInt16LE(name.length, 2);
    parts.push(head, name);
  }
  parts.push(Buffer.alloc(4)); // Padding5, RootStringSize 0 (no root string: usable in any document)
  const header = Buffer.concat(parts);
  header.writeUInt32LE(header.length + buf.length, 0);
  header.writeUInt32LE(buf.length, 4);
  return Buffer.concat([header, buf]);
}

// The EOT header read back (for tests and checks): sizes, version, flags and the four names.
export function readEot(eot) {
  const u32 = (o) => eot.readUInt32LE(o);
  let o = 80;
  const names = [];
  for (let i = 0; i < 4; i++) {
    o += 2;
    const n = eot.readUInt16LE(o);
    o += 2;
    names.push(eot.toString("utf16le", o, o + n));
    o += n;
  }
  o += 2;
  const rootSize = eot.readUInt16LE(o);
  o += 2 + rootSize;
  return {
    eotSize: u32(0), fontDataSize: u32(4), version: u32(8), flags: u32(12), italic: eot[27], weight: u32(28),
    fsType: eot.readUInt16LE(32), magic: eot.readUInt16LE(34), names, fontData: eot.subarray(o),
  };
}
