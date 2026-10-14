// Slide geometry in inches, shared by the layout engine (where free shapes go) and the slide
// master (where placeholders sit), so the two always line up. Layouts always reserve the
// eyebrow band, like a real template, so titles sit at the same height on every slide.
// `logo` ({ height?, aspect? }, or null for none) is the footer logo drawn bottom-left on the
// master: when there is one, every body area ends 0.1in above it instead of overlapping it.

export function slideGeometry({ width, height, marginX }, logo = null) {
  const contentW = width - marginX * 2;
  const box = (y, h) => ({ x: marginX, y, w: contentW, h });
  let logoBox = null;
  let bottomMargin = 0.35;
  if (logo) {
    const h = logo.height || 0.27;
    logoBox = { x: 0.5, y: height - 0.29 - h, w: h * (logo.aspect || 1), h };
    bottomMargin = Math.max(bottomMargin, height - logoBox.y + 0.1);
  }
  return {
    width,
    height,
    marginX,
    contentW,
    bottomMargin,
    logo: logoBox,
    eyebrow: box(0.45, 0.3),
    title: box(0.8, 0.6),
    body: box(1.6, height - bottomMargin - 1.6),
    coverTitle: box(height * 0.36, 1.3),
    coverSubtitle: box(height * 0.36 + 1.4, 0.8),
    sectionTitle: box(height * 0.36, 1.0),
    sectionBody: box(height * 0.36 + 1.1, 1.0),
    centeredTitle: box(height / 2 - 1.0, 0.6),
    centeredBody: box(height / 2 - 0.2, 2.0),
  };
}
