// Fonts to embed (--embed-fonts): for each font name the slides use, the installed family's regular,
// bold, italic and bold italic files (fontconfig's fc-match, or the config's `embedFonts.files`), as
// Embedded OpenType for the writer. A web font's name ("Inter Variable", how @fontsource-variable
// calls Inter) that no installed font has falls back to the family without " Variable"; the slides
// are then written with that name, the one PowerPoint finds in the embedded font.

import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { eotFromTtf, fontInfo } from "./ooxml/eot.mjs";

export const STYLES = {
  regular: { pattern: ":weight=regular:slant=roman", bold: false, italic: false },
  bold: { pattern: ":weight=bold:slant=roman", bold: true, italic: false },
  italic: { pattern: ":weight=regular:slant=italic", bold: false, italic: true },
  boldItalic: { pattern: ":weight=bold:slant=italic", bold: true, italic: true },
};

function fcMatch(pattern) {
  try {
    return execFileSync("fc-match", ["-f", "%{file}", pattern], { encoding: "utf8", stdio: ["ignore", "pipe", "ignore"] }).trim() || null;
  } catch {
    return null;
  }
}

// The file for one style of `family`, if it is that family in that style: fc-match answers with the
// nearest font it has, which may be another family, or the regular face for a bold it lacks.
function styleFile(family, style, files, baseDir) {
  const given = files?.[family]?.[style];
  const path = given ? resolve(baseDir, given) : fcMatch(`${family}${STYLES[style].pattern}`);
  if (!path) return null;
  try {
    const data = readFileSync(path);
    const info = fontInfo(data);
    if (!given && info.family.toLowerCase() !== family.toLowerCase()) return null;
    if (!given && ((info.weight >= 600) !== STYLES[style].bold || info.italic !== STYLES[style].italic)) return null;
    return { path, data, info };
  } catch (err) {
    return { path, error: String(err.message || err) };
  }
}

/**
 * (name) => { family, fonts: { regular?, bold?, italic?, boldItalic? }: EOT Buffers, notes: [] } | null
 * @param {{ files?: Record<string, Record<string, string>>, aliases?: Record<string, string> }} [options]
 *   files: family -> style -> font file (paths relative to baseDir); aliases: font name -> family
 */
export function fontResolver(options = {}, baseDir = process.cwd()) {
  const cache = new Map();
  return (name) => {
    if (cache.has(name)) return cache.get(name);
    const candidates = [options.aliases?.[name], name, name.replace(/ Variable$/, "")].filter(Boolean);
    let found = null;
    for (const family of [...new Set(candidates)]) {
      const regular = styleFile(family, "regular", options.files, baseDir);
      if (!regular || regular.error) continue;
      const fonts = {};
      const notes = [];
      for (const style of Object.keys(STYLES)) {
        const f = style === "regular" ? regular : styleFile(family, style, options.files, baseDir);
        if (!f) continue;
        if (f.error) {
          notes.push(`${family} ${style} (${f.path}): ${f.error}`);
          continue;
        }
        try {
          fonts[style] = eotFromTtf(f.data);
        } catch (err) {
          notes.push(`${family} ${style} (${f.path}): ${String(err.message || err)}`);
        }
      }
      if (fonts.regular) {
        found = { family, fonts, notes };
        break;
      }
      if (notes.length) found = { family, fonts: {}, notes };
    }
    cache.set(name, found);
    return found;
  };
}
