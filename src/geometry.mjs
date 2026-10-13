// Slide geometry in inches, shared by the layout engine (where free shapes go) and the slide
// master (where placeholders sit), so the two always line up. Layouts always reserve the
// eyebrow band, like a real template, so titles sit at the same height on every slide.

export function slideGeometry({ width, height, marginX }) {
  const contentW = width - marginX * 2;
  const box = (y, h) => ({ x: marginX, y, w: contentW, h });
  const bottomMargin = 0.35;
  return {
    width,
    height,
    marginX,
    contentW,
    bottomMargin,
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
