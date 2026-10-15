// Render mode's in-page extractor. capture.mjs loads this file into a slide's page with
// page.addScriptTag and calls window.__pptxExtract(pageEl, canvas). It walks the rendered DOM
// under one `.slidev-page` and returns a Scene: plain JSON describing every visible element -
// its measured box in canvas px (relative to the slide's top-left) and its computed style.
// A classic script with no imports, so it runs as-is in Chromium. Elements it can't describe
// natively come back as `raster` nodes; their elements are tagged data-pptx-node="<id>" so
// capture.mjs can screenshot exactly that element.
(() => {
  const SKIP_TAGS = new Set(["script", "style", "noscript", "template", "link", "meta", "head", "title"]);
  // Drawn as pictures.
  const RASTER_TAGS = {
    svg: "inline SVG", iframe: "iframe", video: "video", audio: "audio",
    object: "embedded object", embed: "embedded object", math: "math",
  };
  const REPLACED = new Set(["img", "svg", "canvas", "video", "iframe", "object", "embed", "math", "audio"]);
  const IMAGE_FORMATS = { png: "png", jpg: "jpg", jpeg: "jpg", gif: "gif", svg: "svg", "svg+xml": "svg" };
  const GENERIC_FONTS = new Set(["serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui", "ui-serif", "ui-sans-serif", "ui-monospace", "ui-rounded", "emoji", "math", "fangsong"]);
  // Stands between two paragraphs (e.g. loose-list <p>s) joined into one run list.
  const BREAK = { br: true };
  const r2 = (v) => Math.round(v * 100) / 100;

  // Splits a comma-separated CSS value at top level only (commas inside rgba(...) etc. stay).
  function splitTop(value) {
    const out = [];
    let depth = 0;
    let cur = "";
    for (const ch of String(value)) {
      if (ch === "(") depth++;
      else if (ch === ")") depth--;
      if (ch === "," && depth === 0) {
        out.push(cur.trim());
        cur = "";
      } else cur += ch;
    }
    if (cur.trim()) out.push(cur.trim());
    return out;
  }

  const probe = document.createElement("canvas").getContext("2d", { willReadFrequently: true });

  // Any CSS color -> { hex: "RRGGBB", alpha }. Fully transparent -> null, unless keepTransparent
  // (gradient stops need their position even when invisible). Computed styles are usually
  // rgb()/rgba(); anything else (oklch(), color()) goes through a 1px canvas.
  function parseColor(css, keepTransparent = false) {
    if (!css || css === "none") return null;
    if (css === "transparent") return keepTransparent ? { hex: "000000", alpha: 0 } : null;
    let r, g, b, a;
    const m = /^rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*(?:,\s*([\d.]+)\s*)?\)$/.exec(String(css).trim());
    if (m) {
      r = +m[1]; g = +m[2]; b = +m[3]; a = m[4] === undefined ? 1 : +m[4];
    } else {
      probe.clearRect(0, 0, 1, 1);
      probe.fillStyle = "rgba(0, 0, 0, 0)";
      probe.fillStyle = css;
      probe.fillRect(0, 0, 1, 1);
      const d = probe.getImageData(0, 0, 1, 1).data;
      r = d[0]; g = d[1]; b = d[2]; a = d[3] / 255;
    }
    if (!(a > 0.001)) return keepTransparent ? { hex: "000000", alpha: 0 } : null;
    const hex = [r, g, b].map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, "0")).join("").toUpperCase();
    return { hex, alpha: Math.round(a * 1000) / 1000 };
  }

  function firstFamily(list) {
    const fams = String(list || "").split(",").map((f) => f.trim().replace(/^["']|["']$/g, "")).filter(Boolean);
    return fams.find((f) => !GENERIC_FONTS.has(f.toLowerCase())) || fams[0] || "";
  }

  function toDeg(a) {
    const v = parseFloat(a);
    if (a.endsWith("turn")) return v * 360;
    if (a.endsWith("grad")) return v * 0.9;
    if (a.endsWith("rad")) return (v * 180) / Math.PI;
    return v;
  }

  // "to top", "to bottom right" -> a CSS angle (0 = up, clockwise), or null if it isn't a side or
  // corner. A corner aims so the 50% line runs through the box's other two corners.
  function sideAngle(target, rc) {
    const w = target.split(/\s+/);
    const sy = w.includes("top") ? -1 : w.includes("bottom") ? 1 : 0;
    const sx = w.includes("left") ? -1 : w.includes("right") ? 1 : 0;
    if (w.length > 2 || w.some((x) => !/^(top|bottom|left|right)$/.test(x)) || (!sx && !sy) || (w.length === 2 && !(sx && sy))) return null;
    return (Math.atan2(sx * rc.h, -sy * rc.w) * 180) / Math.PI;
  }

  // "rgba(26, 131, 59, 0.12) 0%" -> { color, pos: "0%" | null }; null if it isn't a plain color stop
  // (a transition hint, a two-position stop, an interpolation space).
  function parseColorStop(p) {
    const m = /^((?:rgba?|hsla?|hwb|lab|lch|oklab|oklch|color)\([^)]*\)|#[0-9a-f]{3,8}|[a-z]+)(?:\s+(-?[\d.]+(?:%|px)))?$/i.exec(String(p).trim());
    if (!m || !CSS.supports("color", m[1])) return null;
    return { color: parseColor(m[1], true), pos: m[2] ?? null };
  }

  const NOT_NATIVE = { unsupported: "gradient" };

  // One CSS gradient layer, measured against the element's box (rc) ->
  //   { gradient: { type: "linear", angle } | { type: "radial", cx, cy, rx, ry }, + stops [{ pos, color }] }
  // or { unsupported: reason } for repeating, unparseable or unfaithful gradients. A radial
  // gradient is native only when its ellipse has about the box's aspect (DrawingML stretches the
  // gradient to the box) and its center is inside the box.
  function parseGradient(layer, rc) {
    const m = /^(repeating-)?(linear|radial)-gradient\((.*)\)$/s.exec(layer.trim());
    if (!m || m[1]) return NOT_NATIVE;
    const parts = splitTop(m[3]);
    if (parts.length < 2) return NOT_NATIVE;
    const g = { type: m[2] };
    let length;
    if (g.type === "linear") {
      let angle = 180;
      if (/^-?[\d.]+(deg|turn|grad|rad)$/.test(parts[0])) angle = toDeg(parts.shift());
      else if (/^to\s/.test(parts[0])) {
        angle = sideAngle(parts.shift().slice(3).trim(), rc);
        if (angle == null) return NOT_NATIVE;
      }
      g.angle = r2(((angle % 360) + 360) % 360);
      const rad = (g.angle * Math.PI) / 180;
      length = Math.abs(rc.w * Math.sin(rad)) + Math.abs(rc.h * Math.cos(rad));
    } else {
      const shape = parseColorStop(parts[0]) ? "" : parts.shift();
      const [sizePart, atPart] = shape.split(/(?:^|\s+)at\s+/);
      const at = (atPart || "center").trim().split(/\s+/);
      if (at.length > 2) return NOT_NATIVE;
      let [px, py] = at;
      if (px === "top" || px === "bottom") [px, py] = [py || "center", px];
      const frac = (v, size) => {
        if (v === undefined || v === "center") return 0.5;
        if (v === "left" || v === "top") return 0;
        if (v === "right" || v === "bottom") return 1;
        return v.endsWith("%") ? parseFloat(v) / 100 : parseFloat(v) / size;
      };
      const fx = frac(px, rc.w);
      const fy = frac(py, rc.h);
      if (Number.isNaN(fx) || Number.isNaN(fy)) return NOT_NATIVE;
      if (fx < 0 || fx > 1 || fy < 0 || fy > 1) return { unsupported: "radial gradient focus" };
      const cx = fx * rc.w;
      const cy = fy * rc.h;
      const [dl, dr, dt, db] = [cx, rc.w - cx, cy, rc.h - cy];
      const tokens = (sizePart || "").trim().split(/\s+/).filter(Boolean);
      const keywords = tokens.filter((t) => t === "circle" || t === "ellipse");
      const sizes = tokens.filter((t) => t !== "circle" && t !== "ellipse");
      const lengths = sizes.filter((t) => /^[\d.]+(?:px|%)$/.test(t));
      const extents = sizes.filter((t) => /^(?:closest|farthest)-(?:side|corner)$/.test(t));
      if (keywords.length > 1 || lengths.length + extents.length !== sizes.length || extents.length > 1 || (extents.length && lengths.length)) return NOT_NATIVE;
      // Without a keyword, a single length makes a circle; anything else an ellipse.
      const circle = keywords.length ? keywords[0] === "circle" : lengths.length === 1;
      if (lengths.length && (circle ? lengths.length !== 1 || lengths[0].endsWith("%") : lengths.length !== 2)) return NOT_NATIVE;
      let rx;
      let ry;
      if (lengths.length) {
        const len = (t, full) => (t.endsWith("%") ? (parseFloat(t) / 100) * full : parseFloat(t));
        rx = len(lengths[0], rc.w);
        ry = circle ? rx : len(lengths[1], rc.h);
      } else {
        const extent = extents[0] || "farthest-corner";
        const pick = extent.startsWith("closest") ? Math.min : Math.max;
        if (circle) {
          rx = ry = extent.endsWith("side") ? pick(dl, dr, dt, db) : pick(Math.hypot(dl, dt), Math.hypot(dr, dt), Math.hypot(dl, db), Math.hypot(dr, db));
        } else {
          // An ellipse keeps the aspect of its side distances; reaching a corner scales both by sqrt(2).
          const k = extent.endsWith("corner") ? Math.SQRT2 : 1;
          rx = pick(dl, dr) * k;
          ry = pick(dt, db) * k;
        }
      }
      if (!(rx > 0 && ry > 0)) return NOT_NATIVE;
      const aspect = rx / ry / (rc.w / rc.h);
      if (aspect < 0.7 || aspect > 1.4) return { unsupported: "radial gradient shape" };
      g.cx = r2(fx);
      g.cy = r2(fy);
      g.rx = r2(rx);
      g.ry = r2(ry);
      length = rx;
    }
    if (!(length > 0)) return NOT_NATIVE;
    const stops = [];
    for (const p of parts) {
      const s = parseColorStop(p);
      if (!s) return NOT_NATIVE;
      stops.push({ color: s.color, pos: s.pos == null ? null : s.pos.endsWith("%") ? parseFloat(s.pos) / 100 : parseFloat(s.pos) / length });
    }
    if (stops.length < 2) return NOT_NATIVE;
    // Missing end positions are 0 and 1; a position behind an earlier one is pulled up to it;
    // the rest are spread evenly between their positioned neighbours.
    if (stops[0].pos == null) stops[0].pos = 0;
    if (stops[stops.length - 1].pos == null) stops[stops.length - 1].pos = 1;
    let high = -Infinity;
    for (const s of stops) {
      if (s.pos == null) continue;
      s.pos = Math.max(s.pos, high);
      high = s.pos;
    }
    for (let i = 1; i < stops.length - 1; i++) {
      if (stops[i].pos != null) continue;
      let j = i;
      while (stops[j].pos == null) j++;
      const a = stops[i - 1].pos;
      const b = stops[j].pos;
      for (let k = i; k < j; k++) stops[k].pos = a + ((b - a) * (k - i + 1)) / (j - i + 1);
    }
    g.stops = stops.map((s) => ({ pos: Math.round(s.pos * 10000) / 10000, color: s.color }));
    return { gradient: g };
  }

  function formatOf(src) {
    if (!src) return null;
    const data = /^data:image\/([a-z0-9+.-]+)/i.exec(src);
    let ext = data ? data[1] : null;
    if (!ext) {
      try {
        ext = (/\.([a-z0-9]+)$/i.exec(new URL(src, location.href).pathname) || [])[1];
      } catch {
        ext = null;
      }
    }
    return IMAGE_FORMATS[String(ext || "").toLowerCase()] || null;
  }

  // A transform that only rotates (and translates) -> { rotateOnly, rotate }; anything else
  // (scale, skew, mirror, 3D) -> { rotateOnly: false }; none -> null.
  function transformOf(cs) {
    if (!cs.transform || cs.transform === "none") return null;
    const m = new DOMMatrixReadOnly(cs.transform);
    if (!m.is2D) return { rotateOnly: false };
    const sx = Math.hypot(m.a, m.b);
    const sy = Math.hypot(m.c, m.d);
    const skewed = Math.abs(m.a * m.c + m.b * m.d) > 1e-6;
    const mirrored = m.a * m.d - m.b * m.c < 0;
    if (skewed || mirrored || Math.abs(sx - 1) > 0.01 || Math.abs(sy - 1) > 0.01) return { rotateOnly: false };
    const rotate = (Math.atan2(m.b, m.a) * 180) / Math.PI;
    return Math.abs(rotate) < 0.01 ? null : { rotateOnly: true, rotate: r2(rotate) };
  }

  function decoOf(cs) {
    const line = cs.textDecorationLine || "";
    return { underline: line.includes("underline"), strike: line.includes("line-through") };
  }

  function mergeDeco(d, cs) {
    const own = decoOf(cs);
    return { underline: d.underline || own.underline, strike: d.strike || own.strike };
  }

  function transformText(text, cs) {
    const ws = cs.whiteSpace;
    let t = /^(pre|pre-wrap|break-spaces)$/.test(ws) ? text : ws === "pre-line" ? text.replace(/[ \t]+/g, " ") : text.replace(/\s+/g, " ");
    if (cs.textTransform === "uppercase") t = t.toUpperCase();
    else if (cs.textTransform === "lowercase") t = t.toLowerCase();
    else if (cs.textTransform === "capitalize") t = t.replace(/(^|\s)(\S)/g, (_, s, c) => s + c.toUpperCase());
    return t;
  }

  function textAlign(cs) {
    const a = cs.textAlign;
    if (a.includes("center")) return "center";
    if (a.includes("right")) return "right";
    return a;
  }

  function sameColor(a, b) {
    if (!a || !b) return !a && !b;
    return a.hex === b.hex && a.alpha === b.alpha;
  }

  function sameStyle(a, b) {
    return a.font === b.font && a.size === b.size && a.bold === b.bold && a.italic === b.italic &&
      a.underline === b.underline && a.strike === b.strike && a.spacing === b.spacing &&
      sameColor(a.color, b.color) && sameColor(a.highlight, b.highlight);
  }

  function runStyle(cs, deco, highlight) {
    return {
      font: firstFamily(cs.fontFamily),
      size: r2(parseFloat(cs.fontSize)),
      bold: (parseInt(cs.fontWeight, 10) || 400) >= 600,
      italic: cs.fontStyle !== "normal",
      color: parseColor(cs.color),
      spacing: cs.letterSpacing === "normal" ? 0 : r2(parseFloat(cs.letterSpacing) || 0),
      underline: deco.underline,
      strike: deco.strike,
      highlight: highlight || null,
    };
  }

  // Collapses whitespace across run boundaries like the browser does, trims each line, and
  // merges neighbouring runs that look the same.
  function tidyRuns(runs, cs) {
    if (!/^(pre|pre-wrap|break-spaces)$/.test(cs.whiteSpace)) {
      let skipSpace = true;
      for (const r of runs) {
        if (r.br) {
          skipSpace = true;
          continue;
        }
        if (skipSpace) r.text = r.text.replace(/^ +/, "");
        if (r.text) skipSpace = r.text.endsWith(" ");
      }
      for (let i = 0; i < runs.length; i++) {
        if (runs[i].br) continue;
        const next = runs.slice(i + 1).find((x) => x.br || x.text);
        if (!next || next.br) runs[i].text = runs[i].text.replace(/ +$/, "");
      }
    }
    const out = [];
    for (const r of runs) {
      if (!r.text) continue;
      const last = out[out.length - 1];
      if (last && sameStyle(last, r)) last.text += r.text;
      else out.push({ ...r });
    }
    return out.map(({ br, ...r }) => r);
  }

  window.__pptxExtract = function extract(pageEl, canvas) {
    const origin = pageEl.getBoundingClientRect();
    const scale = origin.width / canvas.w || 1;
    const layoutRoot = pageEl.querySelector(".slidev-layout") || pageEl.firstElementChild || pageEl;
    const titleEl = pageEl.querySelector("h1");
    const nodes = [];
    let nextId = 0;

    const rectOf = (r) => ({ x: r2((r.left - origin.left) / scale), y: r2((r.top - origin.top) / scale), w: r2(r.width / scale), h: r2(r.height / scale) });
    const onCanvas = (rc) => rc.w > 0.5 && rc.h > 0.5 && rc.x < canvas.w && rc.y < canvas.h && rc.x + rc.w > 0 && rc.y + rc.h > 0;
    const labelOf = (el) => `${el.localName}${[...el.classList].slice(0, 2).map((c) => `.${c}`).join("")}`;
    const mark = (el) => {
      const id = nextId++;
      el.setAttribute("data-pptx-node", String(id));
      return id;
    };

    function contentBox(el, cs, rc) {
      const n = (p) => parseFloat(cs[p]) || 0;
      const l = n("borderLeftWidth") + n("paddingLeft");
      const t = n("borderTopWidth") + n("paddingTop");
      const r = n("borderRightWidth") + n("paddingRight");
      const b = n("borderBottomWidth") + n("paddingBottom");
      return { x: r2(rc.x + l), y: r2(rc.y + t), w: r2(Math.max(0, rc.w - l - r)), h: r2(Math.max(0, rc.h - t - b)) };
    }

    // The element's own box with the rotation taken out (getBoundingClientRect is the rotated hull).
    function unrotatedRect(el, r) {
      const c = rectOf(r);
      const w = el.offsetWidth;
      const h = el.offsetHeight;
      return { x: r2(c.x + c.w / 2 - w / 2), y: r2(c.y + c.h / 2 - h / 2), w, h };
    }

    function parseShadow(css) {
      if (!css || css === "none") return null;
      for (const part of splitTop(css)) {
        if (/\binset\b/.test(part)) continue;
        const colorMatch = /(rgba?\([^)]*\)|#[0-9a-f]{3,8}\b)/i.exec(part);
        const color = parseColor(colorMatch ? colorMatch[1] : "rgb(0, 0, 0)");
        if (!color) continue;
        const lengths = (part.replace(colorMatch ? colorMatch[1] : "", "").match(/-?[\d.]+px/g) || []).map(parseFloat);
        return { x: lengths[0] || 0, y: lengths[1] || 0, blur: lengths[2] || 0, color };
      }
      return null;
    }

    function decoration(cs, rc) {
      const sides = ["Top", "Right", "Bottom", "Left"]
        .map((s, i) => {
          const width = parseFloat(cs[`border${s}Width`]) || 0;
          const style = cs[`border${s}Style`];
          const color = width && style !== "none" && style !== "hidden" ? parseColor(cs[`border${s}Color`]) : null;
          return { side: "trbl"[i], width, color };
        })
        .filter((s) => s.color);
      const uniform = sides.length === 4 && sides.every((s) => s.width === sides[0].width && sameColor(s.color, sides[0].color));
      const rad = cs.borderTopLeftRadius || "0px";
      const radius = rad.endsWith("%") ? (parseFloat(rad) / 100) * Math.min(rc.w, rc.h) : parseFloat(rad) || 0;
      return {
        fill: parseColor(cs.backgroundColor),
        border: uniform ? { width: sides[0].width, color: sides[0].color } : null,
        sides: !uniform && sides.length ? sides : null,
        radius: r2(radius),
        shadow: parseShadow(cs.boxShadow),
      };
    }

    // A single background layer -> { gradient } | { src, format, fit } | { unsupported }.
    function backgroundImageOf(cs, rc) {
      const bg = cs.backgroundImage;
      if (!bg || bg === "none") return null;
      // A `background` shorthand ending in a color reports an extra "none" image layer.
      const all = splitTop(bg);
      const layers = all.filter((l) => l !== "none");
      if (layers.length > 1) return { unsupported: "background layers" };
      if (!layers.length) return null;
      // The longhands list one value per layer ("cover, auto"); this layer's is the one we want.
      const index = all.findIndex((l) => l !== "none");
      const layerValue = (v) => {
        const list = splitTop(v);
        return list.length ? list[index % list.length] : "";
      };
      const size = layerValue(cs.backgroundSize);
      if (/gradient\(/.test(layers[0])) {
        // A gradient sized or tiled by background-size is a pattern, not a fill.
        return /^(auto|auto auto|100% 100%|cover|contain)$/.test(size) ? parseGradient(layers[0], rc) : NOT_NATIVE;
      }
      const m = /^url\("?(.*?)"?\)$/.exec(layers[0]);
      if (!m) return { unsupported: "background image" };
      const src = new URL(m[1], location.href).href;
      const format = formatOf(src);
      if (!format || format === "svg") return { unsupported: "background image" };
      // The picture is stretched, or cropped / letterboxed around its center, to fill the box;
      // anything that depends on the image's own size, tiling or an off-center position is a picture.
      const noRepeat = /^no-repeat( no-repeat)?$/.test(layerValue(cs.backgroundRepeat));
      const fit = size === "cover" ? "cover" : size === "contain" && noRepeat ? "contain" : size === "100% 100%" ? "fill" : null;
      if (!fit) return { unsupported: "background image size" };
      if (fit !== "fill" && layerValue(cs.backgroundPosition) !== "50% 50%") return { unsupported: "background image position" };
      return { src, format, fit };
    }

    // An SVG <img> is exported as the SVG itself, stretched over its box. The page draws the same
    // only when the box has the SVG's aspect (then fill, contain and cover agree); otherwise the
    // browser letterboxes, crops or leaves it at natural size, so it stays a picture.
    function svgFits(el, cs, rc) {
      if (!/^(fill|contain|cover)$/.test(cs.objectFit)) return false;
      // A transform turns rc into the rotated hull; the element's own layout size is its box.
      const box = contentBox(el, cs, transformOf(cs) ? { x: 0, y: 0, w: el.offsetWidth, h: el.offsetHeight } : rc);
      return box.w > 0 && box.h > 0 && Math.abs(box.w / box.h / (el.naturalWidth / el.naturalHeight) - 1) <= 0.01;
    }

    function rasterReason(el, cs, tag, op, rc) {
      if (RASTER_TAGS[tag]) return RASTER_TAGS[tag];
      if (tag === "img") {
        const fmt = formatOf(el.currentSrc || el.src);
        if (!fmt) return "image format";
        if (op < 0.999) return "image opacity";
        if (fmt === "svg" && el.complete && el.naturalWidth && !svgFits(el, cs, rc)) return "SVG image fit";
      }
      // Code is drawn as unwrapped text over its panel, so a block that scrolls inside the panel
      // (the browser clips it) or wraps (the browser breaks lines the text box won't) is a picture.
      if (tag === "pre") {
        if ((cs.overflowX !== "visible" || cs.overflowY !== "visible") && (el.scrollWidth > el.clientWidth + 1 || el.scrollHeight > el.clientHeight + 1)) {
          return "code overflows its panel";
        }
        if (/^(pre-wrap|break-spaces|pre-line)$/.test(cs.whiteSpace)) {
          const rows = measure([...el.childNodes]);
          if (rows && rows.lines > codeLines(el, cs, null).filter((l) => l.length).length) return "wrapped code";
        }
      }
      if (cs.filter && cs.filter !== "none") return "CSS filter";
      if (cs.backdropFilter && cs.backdropFilter !== "none") return "backdrop-filter";
      const mask = cs.maskImage || cs.webkitMaskImage;
      if (mask && mask !== "none") return "mask";
      if (cs.clipPath && cs.clipPath !== "none") return "clip-path";
      if (cs.mixBlendMode && cs.mixBlendMode !== "normal") return "blend mode";
      if ((cs.backgroundClip || cs.webkitBackgroundClip) === "text") return "background-clip: text";
      const bg = backgroundImageOf(cs, rc);
      if (bg && bg.unsupported) return bg.unsupported;
      if (bg && bg.src && op < 0.999) return "image opacity";
      const t = transformOf(cs);
      if (t && !t.rotateOnly) return "transform";
      if (t && el.children.length) return "rotated content";
      return null;
    }

    function isInlineFlow(el) {
      if (REPLACED.has(el.localName) || SKIP_TAGS.has(el.localName)) return false;
      const d = getComputedStyle(el).display;
      return d === "inline" || d === "contents";
    }

    function inheritedDeco(el) {
      let d = { underline: false, strike: false };
      for (let e = el; e && e !== pageEl.parentElement; e = e.parentElement) d = mergeDeco(d, getComputedStyle(e));
      return d;
    }

    // Text runs from inline content. Elements that aren't inline flow (images, inline-blocks)
    // are collected in `deferred` and become nodes of their own.
    function collectRuns(list, deco, highlight, runs, deferred) {
      for (const n of list) {
        if (n === BREAK) {
          if (runs.length) runs.push({ ...runs[runs.length - 1], text: "\n", br: true });
          continue;
        }
        if (n.nodeType === Node.TEXT_NODE) {
          const pcs = getComputedStyle(n.parentElement);
          runs.push({ ...runStyle(pcs, deco, highlight), text: transformText(n.data, pcs) });
          continue;
        }
        if (n.nodeType !== Node.ELEMENT_NODE) continue;
        const ccs = getComputedStyle(n);
        if (ccs.display === "none" || ccs.visibility === "hidden" || parseFloat(ccs.opacity) === 0) continue;
        if (n.localName === "br") {
          runs.push({ ...runStyle(ccs, deco, highlight), text: "\n", br: true });
          continue;
        }
        if (!isInlineFlow(n)) {
          deferred.push(n);
          continue;
        }
        collectRuns(n.childNodes, mergeDeco(deco, ccs), parseColor(ccs.backgroundColor) || highlight, runs, deferred);
      }
    }

    // The client rects of the text inside a list of inline nodes. Only text counts: an inline
    // element's own border box (padding, background) can overlap the neighbouring lines, so
    // counting it would merge two lines into one. Skips what collectRuns skips, and the
    // non-inline elements that become nodes of their own.
    function textRects(list, out) {
      for (const n of list) {
        if (n === BREAK) continue;
        if (n.nodeType === Node.TEXT_NODE) {
          const range = document.createRange();
          range.selectNodeContents(n);
          for (const x of range.getClientRects()) if (x.width > 0 && x.height > 0) out.push(x);
        } else if (n.nodeType === Node.ELEMENT_NODE) {
          const ccs = getComputedStyle(n);
          if (ccs.display === "none" || ccs.visibility === "hidden" || parseFloat(ccs.opacity) === 0) continue;
          if (isInlineFlow(n)) textRects(n.childNodes, out);
        }
      }
      return out;
    }

    // Measures a list of inline nodes: its rect, how many lines it wraps to, and where its
    // first and last lines start.
    function measure(list) {
      const dom = list.filter((n) => n !== BREAK);
      if (!dom.length) return null;
      const range = document.createRange();
      range.setStartBefore(dom[0]);
      range.setEndAfter(dom[dom.length - 1]);
      let rr = range.getBoundingClientRect();
      if (!rr.width || !rr.height) return null;
      let rects = textRects(dom, []);
      if (rects.length) {
        const left = Math.min(...rects.map((x) => x.left));
        const top = Math.min(...rects.map((x) => x.top));
        rr = new DOMRect(left, top, Math.max(...rects.map((x) => x.right)) - left, Math.max(...rects.map((x) => x.bottom)) - top);
      } else rects = [...range.getClientRects()].filter((x) => x.width > 0 && x.height > 0);
      rects.sort((a, b) => a.top - b.top);
      let lines = 0;
      let bottom = -Infinity;
      let firstTop = rr.top;
      let lastTop = rr.top;
      for (const x of rects) {
        if (x.top >= bottom - 2) {
          if (lines === 0) firstTop = x.top;
          lines++;
          lastTop = x.top;
          bottom = x.bottom;
        } else bottom = Math.max(bottom, x.bottom);
      }
      return { rect: rectOf(rr), lines: Math.max(1, lines), firstTop: r2((firstTop - origin.top) / scale), lastTop: r2((lastTop - origin.top) / scale) };
    }

    function lineHeightOf(cs, m) {
      const lh = parseFloat(cs.lineHeight);
      return r2(lh > 0 ? lh : m.rect.h / m.lines);
    }

    function textNode(el, cs, segment, op, box, rotate) {
      const raw = [];
      const deferred = [];
      collectRuns(segment, inheritedDeco(el), null, raw, deferred);
      const runs = tidyRuns(raw, cs);
      const m = runs.length ? measure(segment) : null;
      if (m) {
        const multi = m.lines > 1;
        let rect = m.rect;
        if (rotate && box) rect = contentBox(el, cs, box);
        else if (multi && box) {
          const cb = contentBox(el, cs, box);
          rect = { x: cb.x, y: m.rect.y, w: cb.w, h: m.rect.h };
        }
        if (onCanvas(rect)) {
          nodes.push({ kind: "text", rect, role: el === titleEl ? "title" : "body", align: textAlign(cs), lineHeight: lineHeightOf(cs, m), singleLine: !multi, runs, opacity: op, rotate });
        }
      }
      for (const d of deferred) walk(d, op);
    }

    // Splits an element's children into runs of inline content (each a text node) and block
    // children (walked on their own), in document order.
    function content(el, cs, op, box, rotate) {
      let segment = [];
      const flush = () => {
        if (segment.length) textNode(el, cs, segment, op, box, rotate);
        segment = [];
      };
      for (const child of el.childNodes) {
        if (child.nodeType === Node.TEXT_NODE) segment.push(child);
        else if (child.nodeType === Node.ELEMENT_NODE) {
          if (isInlineFlow(child)) segment.push(child);
          else {
            flush();
            walk(child, op);
          }
        }
      }
      flush();
    }

    function boxNode(el, cs, rc, rotate, op) {
      if (el === pageEl || !onCanvas(rc)) return;
      const d = decoration(cs, rc);
      const bg = backgroundImageOf(cs, rc);
      // A gradient fills the box, over its background color.
      if (bg && bg.gradient) d.fill = { gradient: bg.gradient, base: d.fill };
      if (d.fill || d.border || d.sides || d.shadow) nodes.push({ kind: "box", rect: rc, ...d, opacity: op, rotate });
      if (bg && bg.src) {
        nodes.push({ kind: "image", id: mark(el), rect: rc, src: bg.src, format: bg.format, fit: bg.fit, natural: null, alt: "", border: null, shadow: null, opacity: op, rotate });
      }
    }

    function imageNode(el, cs, rc, rotate, op) {
      const src = el.currentSrc || el.src;
      if (!el.complete || !el.naturalWidth) {
        nodes.push({ kind: "missing", label: labelOf(el), src });
        return;
      }
      if (!onCanvas(rc)) return;
      const d = decoration(cs, rc);
      nodes.push({
        kind: "image", id: mark(el), rect: contentBox(el, cs, rc), src, format: formatOf(src), fit: cs.objectFit || "fill",
        natural: { w: el.naturalWidth, h: el.naturalHeight }, alt: el.alt || "", border: d.border, shadow: d.shadow, opacity: op, rotate,
      });
    }

    function collectItems(listEl, box, level, items, op) {
      let first = true;
      for (const li of listEl.children) {
        if (li.localName !== "li") continue;
        const lcs = getComputedStyle(li);
        if (lcs.display === "none" || lcs.visibility === "hidden" || parseFloat(lcs.opacity) === 0) continue;
        const segment = [];
        const nested = [];
        const blocks = [];
        for (const c of li.childNodes) {
          if (c.nodeType === Node.TEXT_NODE) segment.push(c);
          else if (c.nodeType !== Node.ELEMENT_NODE) continue;
          else if (c.localName === "ul" || c.localName === "ol") nested.push(c);
          else if (isInlineFlow(c)) segment.push(c);
          else if (c.localName === "p") {
            const pcs = getComputedStyle(c);
            if (pcs.display === "none" || pcs.visibility === "hidden" || parseFloat(pcs.opacity) === 0) continue;
            if (segment.length) segment.push(BREAK);
            segment.push(...c.childNodes);
          } else blocks.push(c);
        }
        const raw = [];
        const deferred = [];
        collectRuns(segment, inheritedDeco(li), null, raw, deferred);
        const runs = tidyRuns(raw, lcs);
        const type = lcs.listStyleType;
        const marker = type === "none" ? "none" : /decimal|alpha|roman|latin|greek|armenian|georgian|cjk/.test(type) ? "number" : "bullet";
        if (runs.length) {
          const m = measure(segment);
          const lr = rectOf(li.getBoundingClientRect());
          const lineHeight = m ? lineHeightOf(lcs, m) : r2(parseFloat(lcs.fontSize) * 1.2);
          const top = m ? m.firstTop : lr.y;
          const prev = items[items.length - 1];
          const item = {
            level, marker,
            markerColor: parseColor(getComputedStyle(li, "::marker").color),
            indent: r2(lr.x + (parseFloat(lcs.paddingLeft) || 0) - box.x),
            spaceBefore: prev ? r2(Math.max(0, top - prev.lastTop - lineHeight)) : 0,
            lineHeight, align: textAlign(lcs), runs, lastTop: m ? m.lastTop : lr.y,
          };
          if (marker === "number" && first) item.start = listEl.localName === "ol" ? listEl.start || 1 : 1;
          items.push(item);
          first = false;
        }
        for (const n of nested) collectItems(n, box, level + 1, items, op);
        for (const b of [...blocks, ...deferred]) walk(b, op);
      }
    }

    function listNode(listEl, box, op) {
      const items = [];
      collectItems(listEl, box, 0, items, op);
      if (items.length && onCanvas(box)) nodes.push({ kind: "list", rect: box, items: items.map(({ lastTop, ...it }) => it), opacity: op });
    }

    // The opacity an element is drawn with inside `root`: the product of its own and its ancestors'
    // below `root` (exclusive), 0 when any of them is not drawn.
    function opacityWithin(el, root) {
      let a = 1;
      for (let e = el; e && e !== root; e = e.parentElement) {
        const s = getComputedStyle(e);
        if (s.display === "none" || s.visibility === "hidden") return 0;
        a *= parseFloat(s.opacity);
      }
      return a;
    }

    function cellBorders(cs) {
      const out = {};
      for (const [side, name] of [["t", "Top"], ["r", "Right"], ["b", "Bottom"], ["l", "Left"]]) {
        const width = parseFloat(cs[`border${name}Width`]) || 0;
        const style = cs[`border${name}Style`];
        const color = width && style !== "none" && style !== "hidden" ? parseColor(cs[`border${name}Color`]) : null;
        if (color) out[side] = { width, color };
      }
      return out;
    }

    // A cell's inline content (its <p> children included) as runs; other blocks are returned
    // to be walked as nodes of their own, on top of the table.
    function cellContent(cell) {
      const segment = [];
      const blocks = [];
      for (const c of cell.childNodes) {
        if (c.nodeType === Node.TEXT_NODE) segment.push(c);
        else if (c.nodeType !== Node.ELEMENT_NODE) continue;
        else if (isInlineFlow(c)) segment.push(c);
        else if (c.localName === "p") {
          const pcs = getComputedStyle(c);
          if (pcs.display === "none" || pcs.visibility === "hidden" || parseFloat(pcs.opacity) === 0) continue;
          if (segment.length) segment.push(BREAK);
          segment.push(...c.childNodes);
        } else blocks.push(c);
      }
      const raw = [];
      const deferred = [];
      collectRuns(segment, inheritedDeco(cell), null, raw, deferred);
      const runs = tidyRuns(raw, getComputedStyle(cell));
      return { runs, m: runs.length ? measure(segment) : null, blocks: [...blocks, ...deferred] };
    }

    // The table as one node: its grid (spans marked on the cell that starts them, the slots they
    // cover as { covered }), the measured column widths and row heights, and every cell's own
    // runs, fill, borders, padding and alignment.
    function tableNode(table, op) {
      const rows = [...table.rows].filter((tr) => getComputedStyle(tr).display !== "none");
      if (!rows.length) return;
      // The HTML table grid: which slot each cell starts in, and which slots spans cover.
      const grid = [];
      const placed = [];
      rows.forEach((tr, r) => {
        grid[r] = grid[r] || [];
        let c = 0;
        for (const td of tr.cells) {
          if (getComputedStyle(td).display === "none") continue;
          while (grid[r][c]) c++;
          const cspan = Math.max(1, td.colSpan || 1);
          const rspan = Math.max(1, Math.min(td.rowSpan || 1, rows.length - r));
          for (let i = 0; i < rspan; i++) {
            grid[r + i] = grid[r + i] || [];
            for (let j = 0; j < cspan; j++) grid[r + i][c + j] = i || j ? (i && j ? "hv" : i ? "v" : "h") : "origin";
          }
          placed.push({ r, c, cspan, rspan, td });
          c += cspan;
        }
      });
      const nCols = Math.max(...grid.map((row) => row.length));
      // Column edges from the cells that start in each column; the last edge is the widest right side.
      const left = Array(nCols).fill(null);
      let right = -Infinity;
      for (const p of placed) {
        const rc = rectOf(p.td.getBoundingClientRect());
        if (left[p.c] === null || rc.x < left[p.c]) left[p.c] = rc.x;
        right = Math.max(right, rc.x + rc.w);
      }
      // A column no cell starts in (every cell over it spans in from the left) shares the gap evenly.
      for (let c = 1; c < nCols; c++) {
        if (left[c] !== null) continue;
        let next = c + 1;
        while (next < nCols && left[next] === null) next++;
        left[c] = left[c - 1] + ((next < nCols ? left[next] : right) - left[c - 1]) / (next - c + 1);
      }
      const colW = left.map((x, c) => r2((c + 1 < nCols ? left[c + 1] : right) - x));
      const rowRects = rows.map((tr) => rectOf(tr.getBoundingClientRect()));
      const rowH = rowRects.map((rc) => rc.h);
      const collapsed = getComputedStyle(table).borderCollapse === "collapse";
      const empty = () => ({ runs: [], align: "left", valign: "top", lineHeight: 0, fill: null, borders: {}, padding: { t: 0, r: 0, b: 0, l: 0 }, colspan: 1, rowspan: 1 });
      const cells = grid.map((row) => Array.from({ length: nCols }, (_, c) => (row[c] && row[c] !== "origin" ? { covered: row[c] } : empty())));
      const later = [];
      for (const p of placed) {
        const cs = getComputedStyle(p.td);
        const { runs, m, blocks } = cellContent(p.td);
        later.push(...blocks);
        const n = (prop) => parseFloat(cs[prop]) || 0;
        const borders = cellBorders(cs);
        if (collapsed) {
          // With collapsed borders a row's rules (Slidev's `tr { border-bottom }`) are drawn along the
          // cells in it; the wider rule wins, the cell's own on a tie.
          const top = cellBorders(getComputedStyle(rows[p.r]));
          const bottom = cellBorders(getComputedStyle(rows[p.r + p.rspan - 1]));
          const edge = { t: top.t, b: bottom.b, l: p.c === 0 ? top.l : null, r: p.c + p.cspan === nCols ? top.r : null };
          for (const side of ["t", "r", "b", "l"]) if (edge[side] && !(borders[side] && borders[side].width >= edge[side].width)) borders[side] = edge[side];
        }
        // The cell's own background, else the nearest row's or row group's behind it.
        let fill = null;
        for (let e = p.td; e && e !== table && !fill; e = e.parentElement) fill = parseColor(getComputedStyle(e).backgroundColor);
        cells[p.r][p.c] = {
          runs, align: textAlign(cs), valign: cs.verticalAlign,
          lineHeight: m ? lineHeightOf(cs, m) : r2(parseFloat(cs.fontSize) * 1.2),
          fill, borders,
          padding: { t: n("paddingTop"), r: n("paddingRight"), b: n("paddingBottom"), l: n("paddingLeft") },
          colspan: p.cspan, rowspan: p.rspan,
        };
      }
      const rect = { x: left[0], y: rowRects[0].y, w: r2(colW.reduce((a, b) => a + b, 0)), h: r2(rowH.reduce((a, b) => a + b, 0)) };
      if (onCanvas(rect)) nodes.push({ kind: "table", rect, colW, rowH, rows: cells, opacity: op });
      if (table.caption) walk(table.caption, op);
      for (const b of later) walk(b, op);
    }

    // tab-size in characters: a number as it is, a length divided by the width of one character.
    function tabWidth(cs) {
      const n = parseFloat(cs.tabSize);
      if (!(n > 0)) return 8;
      if (!/px$/.test(cs.tabSize)) return Math.max(1, Math.round(n));
      probe.font = `${cs.fontSize} ${cs.fontFamily}`;
      const ch = probe.measureText("0").width;
      return ch > 0 ? Math.max(1, Math.round(n / ch)) : 8;
    }

    // A code block's text as one list of runs per source line: token colors kept (dimmed by the
    // opacity of the spans around them), tabs expanded to the next tab stop, spaces untouched.
    // `rects`, when given, receives the first client rect of each source line that has text.
    function codeLines(pre, cs, rects) {
      const lines = [[]];
      const tab = tabWidth(cs);
      const range = document.createRange();
      let col = 0;
      const walker = document.createTreeWalker(pre, NodeFilter.SHOW_TEXT);
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        const dim = opacityWithin(n.parentElement, pre);
        if (!(dim > 0.001)) continue;
        const style = runStyle(getComputedStyle(n.parentElement), inheritedDeco(n.parentElement), null);
        if (style.color && dim < 1) style.color = { ...style.color, alpha: Math.round(style.color.alpha * dim * 1000) / 1000 };
        let offset = 0;
        n.data.split("\n").forEach((part, i) => {
          if (i > 0) {
            lines.push([]);
            col = 0;
            offset++;
          }
          const start = offset;
          offset += part.length;
          if (!part) return;
          let text = "";
          for (const ch of part) {
            if (ch === "\t") {
              const w = tab - (col % tab);
              text += " ".repeat(w);
              col += w;
            } else {
              text += ch;
              col++;
            }
          }
          lines[lines.length - 1].push({ ...style, text });
          if (rects && !rects[lines.length - 1]) {
            range.setStart(n, start);
            range.setEnd(n, offset);
            const r = range.getClientRects()[0];
            if (r) rects[lines.length - 1] = r;
          }
        });
      }
      while (lines.length && !lines[lines.length - 1].length) lines.pop();
      return lines;
    }

    function codeNode(pre, cs, box, op) {
      // With `line-height: normal` the pitch between lines is measured, from the first and last
      // source line that have text (blank lines between them have no rect of their own).
      const lh = parseFloat(cs.lineHeight);
      const rects = lh > 0 ? null : {};
      const lines = codeLines(pre, cs, rects);
      if (!lines.length) return;
      let pitch = r2(lh);
      if (!(lh > 0)) {
        const rows = Object.keys(rects).map(Number);
        const first = rows[0];
        const last = rows[rows.length - 1];
        const px = rows.length ? (last > first ? (rects[last].top - rects[first].top) / (last - first) : rects[first].height) : parseFloat(cs.fontSize) * 1.2;
        pitch = r2(rows.length ? px / scale : px);
      }
      nodes.push({
        kind: "code", rect: contentBox(pre, cs, box), lines: lines.map((runs) => tidyRuns(runs, { whiteSpace: "pre" })),
        lineHeight: pitch, opacity: op,
      });
    }

    // The Chart.js chart drawn on a canvas: through the Vue component that owns it (vue-chartjs
    // exposes it as `chart`), else Chart.getChart when Chart.js is a page global.
    function findChart(canvasEl) {
      const components = [];
      let chart = null;
      let inst = canvasEl.__vueParentComponent || (canvasEl.parentElement && canvasEl.parentElement.__vueParentComponent);
      for (let i = 0; inst && i < 25; i++, inst = inst.parent) {
        const name = inst.type && (inst.type.__name || inst.type.name);
        if (name) components.push(name);
        for (const src of [inst.exposed, inst.setupState]) {
          if (chart || !src) continue;
          let c = null;
          try {
            c = src.chart;
          } catch {
            c = null;
          }
          const v = c && (c.value !== undefined ? c.value : c);
          if (v && v.config && v.data && v.canvas === canvasEl) chart = v;
        }
      }
      if (!chart && window.Chart && typeof window.Chart.getChart === "function") chart = window.Chart.getChart(canvasEl) || null;
      return chart ? { chart, components } : null;
    }

    // How the value axis's tick labels write their number - {pre: "", suf: "%", decimals: 0} for 0%,
    // 20%, 40% - read from the labels Chart.js drew, so a tick callback is never run. It is
    // written as the Excel format 0"%". `decimals` is the fixed number of decimals every label
    // has (0.0, 10.0); labels that differ (Chart.js's own 0, 0.5, 1.0) are plain numbers, and so
    // is 1,000: the native axis shows the same values. null when the labels aren't one number
    // each with the same prefix and suffix, or the affix can't be quoted in a format code.
    function tickAffix(scale) {
      const ticks = scale.ticks || [];
      if (!ticks.length) return { pre: "", suf: "", decimals: 0 };
      const split = (label) => {
        const m = /[-+]?(?:\d[\d,]*(?:\.\d+)?|\.\d+)/.exec(label);
        return m ? { pre: label.slice(0, m.index), core: m[0], suf: label.slice(m.index + m[0].length) } : null;
      };
      const first = split(String(ticks[0].label));
      if (!first || /["\\]/.test(first.pre + first.suf)) return null;
      const plain = ticks.every((t) => {
        const part = split(String(t.label));
        const n = part && Number(part.core.replace(/,/g, ""));
        return part && part.pre === first.pre && part.suf === first.suf && Number.isFinite(n) && Math.abs(n - t.value) <= 1e-6 * Math.max(1, Math.abs(t.value));
      });
      if (!plain) return null;
      const digits = new Set(ticks.map((t) => (split(String(t.label)).core.split(".")[1] || "").length));
      return { pre: first.pre, suf: first.suf, decimals: digits.size === 1 ? [...digits][0] : 0 };
    }

    // The visible datasets and the scales a bar chart is drawn on, found through the datasets' own
    // scales (else by axis letter) - Chart.js scale ids are whatever the options named them. Every
    // value scale that is drawn counts, used by a dataset or not.
    function chartParts(chart) {
      const horizontal = (chart.options || {}).indexAxis === "y";
      const datasets = chart.data.datasets || [];
      const visible = datasets.map((d, i) => i).filter((i) => chart.isDatasetVisible(i));
      const metas = visible.map((i) => chart.getDatasetMeta(i));
      const scales = Object.values(chart.scales || {});
      const byAxis = (axis) => scales.find((s) => s.axis === axis);
      const used = metas.map((m) => m.vScale).filter(Boolean);
      const valueAxis = horizontal ? "x" : "y";
      const valueScales = new Set([...used, ...scales.filter((s) => s.axis === valueAxis && s.options.display !== false)]);
      return {
        horizontal,
        datasets,
        visible,
        valueScales,
        valScale: used[0] || byAxis(valueAxis),
        catScale: (metas[0] && metas[0].iScale) || byAxis(horizontal ? "y" : "x"),
      };
    }

    // Why this chart can't be a native one without changing what it shows (null when it can).
    // The canvas is then exported as a picture, with `chart <reason>` as the notice's reason.
    function unsupportedChart(chart, parts) {
      const { datasets, visible, valueScales, valScale, catScale } = parts;
      const type = chart.config && chart.config.type;
      if (type !== "bar") return `type ${type}`;
      if (datasets.some((d) => d.type && d.type !== "bar")) return "mixed chart types";
      if (!visible.length) return "no visible datasets";
      const shown = visible.map((i) => datasets[i]);
      const custom = (p) => p !== undefined && p !== true;
      if (custom((chart.options || {}).parsing) || shown.some((d) => custom(d.parsing))) return "custom parsing";
      const plain = (v) => v == null || typeof v === "number" || (typeof v === "string" && v.trim() !== "" && Number.isFinite(Number(v)));
      if (shown.some((d) => !Array.isArray(d.data) || !d.data.every(plain))) return "data format";
      if (!valScale || !catScale) return "axes";
      if (valScale.type !== "linear") return "value axis type";
      if (catScale.type !== "category") return "category axis type";
      if (valueScales.size > 1) return "second value axis";
      const sv = valScale.options.stacked;
      const sc = catScale.options.stacked;
      // Stacked bars are one stack: both axes stacked, no `stack` names splitting them. A dataset
      // that names a stack without `stacked` set is stacked by Chart.js too.
      if (typeof sv === "string" || typeof sc === "string" || !!sv !== !!sc) return "stack groups";
      if (sv && new Set(shown.map((d) => d.stack)).size > 1) return "stack groups";
      if (sv === undefined && sc === undefined && shown.some((d) => d.stack !== undefined)) return "stack groups";
      if (catScale.options.display === false) return "hidden category axis";
      if (valScale.options.reverse || catScale.options.reverse) return "reversed axis";
      const plugins = (chart.options || {}).plugins || {};
      if ((plugins.title && plugins.title.display) || (plugins.subtitle && plugins.subtitle.display)) return "title";
      for (const d of shown) {
        const c = d.backgroundColor;
        if (typeof c === "function" || (c && typeof c === "object" && !Array.isArray(c))) return "scripted colors";
        if (Array.isArray(c)) {
          if (c.some((x) => typeof x !== "string")) return "scripted colors";
          if (new Set(c.map((x) => JSON.stringify(parseColor(x)))).size > 1) return "per-bar colors";
        }
      }
      if (valScale.options.display !== false && !tickAffix(valScale)) return "tick labels";
      return null;
    }

    function chartSpec(chart, parts) {
      const { horizontal, datasets, visible, valScale, catScale } = parts;
      const vo = valScale.options;
      const o = chart.options || {};
      const legend = (o.plugins && o.plugins.legend) || {};
      const num = (v) => (typeof v === "number" && Number.isFinite(v) ? v : null);
      const css = (v, fallback) => parseColor(typeof v === "string" && v ? v : fallback);
      // A bar chart's value axis begins at zero unless the options say otherwise; a suggested
      // bound is where the drawn axis actually ended up.
      const bound = (key, suggested, drawn) => {
        const explicit = num(vo[key]);
        if (explicit !== null) return explicit;
        return num(vo[suggested]) !== null ? num(drawn) : null;
      };
      const min = bound("min", "suggestedMin", valScale.min);
      const shown = vo.display !== false;
      const affix = shown ? tickAffix(valScale) : null;
      const grid = vo.grid || {};
      const seriesColor = (i) => {
        const c = datasets[i].backgroundColor;
        const given = Array.isArray(c) ? c[0] : c;
        if (typeof given === "string") return parseColor(given);
        // Nothing set: whatever Chart.js drew with (its default, or the Colors plugin's pick).
        try {
          const drawn = chart.getDatasetMeta(i).controller.getStyle(0).backgroundColor;
          return typeof drawn === "string" ? parseColor(drawn) : null;
        } catch {
          return null;
        }
      };
      const tick = (s) => s.options.ticks && s.options.ticks.color;
      const q = (s) => (s ? `"${s}"` : "");
      return {
        type: horizontal ? "bar" : "column",
        categories: Array.from(catScale.getLabels(), (l) => String(l)),
        series: visible.map((i) => ({ name: String(datasets[i].label ?? ""), values: Array.from(datasets[i].data, (v) => (v == null ? null : num(Number(v)))), color: seriesColor(i) })),
        min: min !== null ? min : valScale.min === 0 ? 0 : null,
        max: bound("max", "suggestedMax", valScale.max),
        stacked: !!vo.stacked,
        valueAxis: shown,
        legend: legend.display !== false && visible.length > 0,
        legendPos: { top: "t", bottom: "b", left: "l", right: "r" }[legend.position] || "t",
        numFmt: affix && (affix.pre || affix.suf || affix.decimals) ? `${q(affix.pre)}0${affix.decimals ? `.${"0".repeat(affix.decimals)}` : ""}${q(affix.suf)}` : null,
        gridColor: !shown || grid.display === false || grid.drawOnChartArea === false ? null : css(grid.color, "rgba(0, 0, 0, 0.1)"),
        labelColor: css(tick(valScale) || tick(catScale), "#666666"),
        legendColor: css(legend.labels && legend.labels.color, "#666666"),
      };
    }

    // A canvas is a native chart when it holds a Chart.js bar chart that can be drawn faithfully,
    // else a picture (with the reason).
    function chartNode(canvasEl, cs, rc, box, rotate, op) {
      if (!onCanvas(rc)) return;
      const found = findChart(canvasEl);
      const picture = (reason) =>
        nodes.push({ kind: "raster", id: mark(canvasEl), rect: rc, reason, label: labelOf(canvasEl), alt: canvasEl.getAttribute("aria-label") || "" });
      if (!found) return picture("canvas");
      let spec = null;
      try {
        const parts = chartParts(found.chart);
        const why = unsupportedChart(found.chart, parts);
        if (why) return picture(`chart ${why}`);
        if (rotate) return picture("rotated chart");
        if (op < 0.999) return picture("chart opacity");
        spec = chartSpec(found.chart, parts);
      } catch {
        return picture("chart unreadable");
      }
      // The chart is drawn inside the canvas's content box; a background, border or shadow on the
      // canvas itself is its own box.
      boxNode(canvasEl, cs, box, rotate, op);
      nodes.push({ kind: "chart", id: nextId++, rect: contentBox(canvasEl, cs, box), components: found.components, chart: spec, opacity: op });
    }

    function walk(el, opacity) {
      const tag = el.localName;
      if (SKIP_TAGS.has(tag)) return;
      const cs = getComputedStyle(el);
      if (cs.display === "none" || cs.visibility === "hidden" || cs.visibility === "collapse") return;
      const op = opacity * (parseFloat(cs.opacity) || 0);
      if (!(op > 0.001)) return;
      if (cs.display === "contents") {
        content(el, cs, op, null, 0);
        return;
      }
      if (el === layoutRoot) {
        content(el, cs, op, rectOf(el.getBoundingClientRect()), 0);
        return;
      }
      const r = el.getBoundingClientRect();
      const rc = rectOf(r);
      const reason = rasterReason(el, cs, tag, op, rc);
      if (reason) {
        if (onCanvas(rc)) nodes.push({ kind: "raster", id: mark(el), rect: rc, reason, label: labelOf(el), alt: el.getAttribute("alt") || el.getAttribute("aria-label") || "" });
        return;
      }
      const t = transformOf(cs);
      const rotate = t ? t.rotate : 0;
      const box = rotate ? unrotatedRect(el, r) : rc;
      if (tag === "img") {
        imageNode(el, cs, box, rotate, op);
        return;
      }
      if (tag === "canvas") {
        chartNode(el, cs, rc, box, rotate, op);
        return;
      }
      if (tag === "table") {
        boxNode(el, cs, box, rotate, op);
        tableNode(el, op);
        return;
      }
      if (tag === "pre") {
        boxNode(el, cs, box, rotate, op);
        codeNode(el, cs, box, op);
        return;
      }
      boxNode(el, cs, box, rotate, op);
      if (tag === "ul" || tag === "ol") {
        listNode(el, box, op);
        return;
      }
      content(el, cs, op, box, rotate);
    }

    function backgroundOf() {
      for (let e = layoutRoot; e && e !== pageEl.parentElement; e = e.parentElement) {
        const cs = getComputedStyle(e);
        if (cs.backgroundImage && cs.backgroundImage !== "none") {
          const bg = backgroundImageOf(cs, rectOf(e.getBoundingClientRect()));
          if (bg && bg.gradient) return { gradient: bg.gradient, base: parseColor(cs.backgroundColor) };
          const why = bg && bg.unsupported;
          return { raster: true, reason: why === "gradient" ? "gradient background" : /gradient/.test(why || "") ? why : "background image" };
        }
        const c = parseColor(cs.backgroundColor);
        if (c) return { color: c };
      }
      return { color: { hex: "FFFFFF", alpha: 1 } };
    }

    content(pageEl, getComputedStyle(pageEl), 1, rectOf(origin), 0);
    const layout = layoutRoot.classList && layoutRoot.classList.contains("slidev-layout")
      ? [...layoutRoot.classList].find((c) => c !== "slidev-layout") || "default"
      : "default";
    return { canvas: { w: canvas.w, h: canvas.h }, layout, background: backgroundOf(), nodes };
  };
})();
