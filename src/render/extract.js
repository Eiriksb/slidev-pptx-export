// Render mode's in-page extractor. capture.mjs loads this file into a slide's page with
// page.addScriptTag and calls window.__pptxExtract(pageEl, canvas). It walks the rendered DOM
// under one `.slidev-page` and returns a Scene: plain JSON describing every visible element -
// its measured box in canvas px (relative to the slide's top-left) and its computed style.
// A classic script with no imports, so it runs as-is in Chromium. Elements it can't describe
// natively come back as `raster` nodes; their elements are tagged data-pptx-node="<id>" so
// capture.mjs can screenshot exactly that element.
(() => {
  const SKIP_TAGS = new Set(["script", "style", "noscript", "template", "link", "meta", "head", "title"]);
  // Drawn as pictures; later tasks turn tables, code blocks and charts native.
  const RASTER_TAGS = {
    canvas: "canvas", svg: "inline SVG", iframe: "iframe", video: "video", audio: "audio",
    object: "embedded object", embed: "embedded object", math: "math", table: "table", pre: "code block",
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

  // Any CSS color -> { hex: "RRGGBB", alpha }, or null when fully transparent. Computed styles
  // are usually rgb()/rgba(); anything else (oklch(), color()) goes through a 1px canvas.
  function parseColor(css) {
    if (!css || css === "transparent" || css === "none") return null;
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
    if (!(a > 0.001)) return null;
    const hex = [r, g, b].map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, "0")).join("").toUpperCase();
    return { hex, alpha: Math.round(a * 1000) / 1000 };
  }

  function firstFamily(list) {
    const fams = String(list || "").split(",").map((f) => f.trim().replace(/^["']|["']$/g, "")).filter(Boolean);
    return fams.find((f) => !GENERIC_FONTS.has(f.toLowerCase())) || fams[0] || "";
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

    // A single url() background image -> its source and fit; anything else is unsupported.
    function backgroundImageOf(cs) {
      const bg = cs.backgroundImage;
      if (!bg || bg === "none") return null;
      // A `background` shorthand ending in a color reports an extra "none" image layer.
      if (splitTop(bg).filter((l) => l !== "none").length > 1) return { unsupported: "background layers" };
      const m = /^url\("?(.*?)"?\)$/.exec(bg);
      if (!m) return { unsupported: /gradient\(/.test(bg) ? "gradient" : "background image" };
      const src = new URL(m[1], location.href).href;
      const format = formatOf(src);
      if (!format || format === "svg") return { unsupported: "background image" };
      const size = cs.backgroundSize;
      return { src, format, fit: size === "cover" ? "cover" : size === "contain" ? "contain" : "fill" };
    }

    function rasterReason(el, cs, tag, op) {
      if (RASTER_TAGS[tag]) return RASTER_TAGS[tag];
      if (tag === "img") {
        const fmt = formatOf(el.currentSrc || el.src);
        if (fmt === "svg") return "SVG image";
        if (!fmt) return "image format";
        if (op < 0.999) return "image opacity";
      }
      if (cs.filter && cs.filter !== "none") return "CSS filter";
      if (cs.backdropFilter && cs.backdropFilter !== "none") return "backdrop-filter";
      const mask = cs.maskImage || cs.webkitMaskImage;
      if (mask && mask !== "none") return "mask";
      if (cs.clipPath && cs.clipPath !== "none") return "clip-path";
      if (cs.mixBlendMode && cs.mixBlendMode !== "normal") return "blend mode";
      if ((cs.backgroundClip || cs.webkitBackgroundClip) === "text") return "background-clip: text";
      const bg = backgroundImageOf(cs);
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
        if (ccs.display === "none" || ccs.visibility === "hidden") continue;
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

    // Measures a list of inline nodes: its rect, how many lines it wraps to, and where its
    // first and last lines start.
    function measure(list) {
      const dom = list.filter((n) => n !== BREAK);
      if (!dom.length) return null;
      const range = document.createRange();
      range.setStartBefore(dom[0]);
      range.setEndAfter(dom[dom.length - 1]);
      const rr = range.getBoundingClientRect();
      if (!rr.width || !rr.height) return null;
      const rects = [...range.getClientRects()].filter((x) => x.width > 0 && x.height > 0).sort((a, b) => a.top - b.top);
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
      if (d.fill || d.border || d.sides || d.shadow) nodes.push({ kind: "box", rect: rc, ...d, opacity: op, rotate });
      const bg = backgroundImageOf(cs);
      if (bg && !bg.unsupported) {
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
        if (lcs.display === "none" || lcs.visibility === "hidden") continue;
        const segment = [];
        const nested = [];
        const blocks = [];
        for (const c of li.childNodes) {
          if (c.nodeType === Node.TEXT_NODE) segment.push(c);
          else if (c.nodeType !== Node.ELEMENT_NODE) continue;
          else if (c.localName === "ul" || c.localName === "ol") nested.push(c);
          else if (isInlineFlow(c)) segment.push(c);
          else if (c.localName === "p") {
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
      const reason = rasterReason(el, cs, tag, op);
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
          return { raster: true, reason: /gradient\(/.test(cs.backgroundImage) ? "gradient background" : "background image" };
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
