// Theme color mapping. The config's named colors become the deck's theme (so PowerPoint's
// Design > Variants can recolor everything), and any hex color layout uses is resolved back to
// a theme slot when it matches one - otherwise it stays a fixed hex color.

const SLOT_ORDER = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];

export function normalizeHex(hex) {
  let h = String(hex ?? "").trim().replace(/^#/, "").toUpperCase();
  if (/^[0-9A-F]{3}$/.test(h)) h = h.split("").map((c) => c + c).join("");
  if (!/^[0-9A-F]{6}$/.test(h)) throw new Error(`Invalid color ${JSON.stringify(hex)} (expected RRGGBB)`);
  return h;
}

export function themeColors(colors) {
  const n = normalizeHex;
  return {
    dk1: n(colors.text),
    lt1: "FFFFFF",
    dk2: n(colors.accentDark),
    lt2: n(colors.panel),
    accent1: n(colors.accent),
    accent2: n(colors.accent2),
    accent3: n(colors.muted),
    accent4: n(colors.border),
    // The config has no fifth/sixth accent; neutral gray and amber keep charts with extra
    // series readable.
    accent5: "9CA3AF",
    accent6: "F59E0B",
    hlink: n(colors.accent),
    folHlink: n(colors.accentDark),
  };
}

export function toColor(hex, theme) {
  const h = normalizeHex(hex);
  return SLOT_ORDER.find((slot) => theme[slot] === h) || `#${h}`;
}

export function hexOf(color, theme) {
  return color.startsWith("#") ? normalizeHex(color) : theme[color];
}

// CSS opacity composited against the white slide background - the faded shade a Tailwind
// `opacity-NN` wrapper actually produces on screen.
export function blendWithWhite(hex, alpha) {
  const n = parseInt(normalizeHex(hex), 16);
  const mix = (c) => Math.round(c * alpha + 255 * (1 - alpha));
  return [mix((n >> 16) & 0xff), mix((n >> 8) & 0xff), mix(n & 0xff)]
    .map((v) => v.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();
}
