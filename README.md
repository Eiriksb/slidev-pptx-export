# slidev-pptx-export

Converts a [Slidev](https://sli.dev) markdown deck into a fully editable native `.pptx` -
real PowerPoint text boxes, tables, shapes and charts, not the screenshot-per-slide export
that `slidev export` produces.

It understands plain markdown (headings, paragraphs, lists, tables, blockquotes, images,
fenced code blocks, `**bold**`/`*italic*`/`` `code` ``) plus a handful of common Slidev
authoring patterns: frontmatter-separated slides, Tailwind/UnoCSS grid layouts (`grid
grid-cols-N` -> a row of cards), `text-2xl`+ stat numbers, `opacity-NN` fades, and `<Vue
Component />` tags resolved into native charts via a `chartResolvers` config.

## What you get in PowerPoint

- **Real layouts and placeholders.** Slide titles sit in the layout's title placeholder, so
  outline view, the Accessibility Checker, Reset Slide and layout switching work. Text-only
  slides (paragraphs, bullets, sub-headings) put their body in the content placeholder, with
  native bullets and numbering and "shrink text on overflow". Slides with cards, tables, code,
  charts or images use the *Title Only* layout plus positioned shapes.
- **A theme built from your config.** `colors` and `fonts` become real theme colors and
  heading/body fonts in `theme1.xml`, and the slide master and layouts reference them the way
  a hand-built PowerPoint template would.
- **Slidev layouts mapped:** first slide / `cover` / `intro` → *Title Slide*, `center` / `end`
  → *Centered*, `section` → *Section Header*, everything else → *Title and Content* or *Title
  Only*.
- **Speaker notes** from the slide's last `<!-- ... -->` comment (Slidev's convention).
- **Native tables** styled by a deck table style, and **native charts** backed by an embedded
  Excel workbook that holds the chart's data, the way a chart built directly in PowerPoint is.
  A table too long for its slide continues on extra slides - same title plus "(cont.)", header
  row repeated - so the export can have more slides than the Slidev deck.
- **Overflow warnings:** content that still doesn't fit after shrinking to 55% prints a
  warning naming the slide, so you know which slides to split by hand.
- **Clean files:** content types and relationships are generated from the parts actually
  written, so nothing is orphaned or missing; `npm run render` round-trips every export
  through LibreOffice as a render check, and the same input always produces a byte-identical
  `.pptx`.

## Usage

```bash
npx slidev-pptx-export slides.md slides-editable.pptx
```

By default it looks for a `pptx.config.mjs` next to `slides.md` and merges it over the
built-in generic defaults. Pass `--config path/to/config.mjs` to use a different one.

Programmatic API:

```js
import { convert } from "slidev-pptx-export";

await convert({
  inputPath: "slides.md",
  outputPath: "slides-editable.pptx",
  config: { colors: { accent: "1A833B" } },
});
```

## Render mode (`--render`)

By default the exporter reads only `slides.md`, so custom CSS, HTML and Vue components reach
the `.pptx` only through what `pptx.config.mjs` restates. With `--render` it instead runs the
deck with its own Slidev in headless Chromium and converts what is actually drawn:

```bash
npm i -D playwright-chromium   # once, in the deck folder (the package `slidev export` uses)
npx playwright install chromium
npx slidev-pptx-export slides.md slides-editable.pptx --render
```

- Text keeps its rendered font, size, weight, color, letter spacing and position; slide
  titles still go into title placeholders.
- Boxes with a background, border, rounded corners or shadow (cards, callouts) become native
  shapes; lists get native bullets in their rendered colors.
- Images become native pictures. Anything PowerPoint can't express natively is drawn as a
  picture of just that element, with a notice naming the slide.
- `colors` and `fonts` in `pptx.config.mjs` still build the theme (rendered colors that match
  become theme colors); `logo`, `eyebrowClass` and `cssVarColors` aren't needed.
  `render: { timeout }` sets the per-slide wait (default 15000 ms). The first slide also waits
  for Slidev's cold start (Vite dependency optimisation), so slow machines or heavy decks may
  need a larger timeout.
- A slide that can't be rendered falls back to the markdown export, with a warning.

Render mode needs the deck's dependencies installed (`npm install` in the deck folder) and
takes roughly 10–30 seconds per deck.

## Config

A config file's default export is deep-merged over the built-in defaults
(`src/default-config.mjs`). Nothing is required - with no config at all you get a plain,
readable deck in a neutral color scheme.

```js
// pptx.config.mjs
export default {
  slide: { width: 13.333, height: 7.5, marginX: 0.6 }, // inches; 13.333x7.5 = 16:9

  colors: {
    accent: "2563EB",     // eyebrow label, cover subtitle color, chart accents
    accentDark: "1E3A8A", // headings, **bold**, `code`
    accent2: "7C3AED",    // cover title
    panel: "F1F5F9",      // card/table-header backgrounds
    border: "E2E8F0",
    text: "1F2937",
    muted: "6B7280",      // faded/opacity-NN text
  },

  fonts: { heading: "Arial", body: "Arial", mono: "Courier New" },

  lang: "en-US",     // language tag on every text run (spell-check)
  timestamps: false, // true = write the export time into the file (breaks byte-identical output)

  // Inline style="color: var(--x)" -> one of the color keys above, for decks that color
  // text via CSS variables rather than Tailwind classes.
  cssVarColors: { "--brand-blue": "accent2" },

  // <div class="..."> that renders as the small uppercase label above a slide's title.
  eyebrowClass: "eyebrow",

  // Footer logo drawn bottom-left on every slide, or omit/null for none. Slide bodies end
  // 0.1in above it, so content never runs into the logo.
  logo: { path: "./assets/logo.png", aspect: 598 / 165, height: 0.27 },

  // Directory (relative to the input .md) chart resolvers may read data from.
  dataDir: "public/data",

  // <ComponentName /> tag -> (dataDirPathOrNull) => ResolvedChart | null
  chartResolvers: {
    MyChart(dataDir) {
      /* read dataDir/whatever.json yourself, shape it into one of: */
      return { kind: "statcards", cards: [{ title: "...", stats: [{ value: "42%", label: "..." }] }] };
      // or: { kind: "barchart-panels", panels: [{ title, color, rows: [{ model, success }] }], min: 0, max: 100 }
      // or: { kind: "barchart", labels: [...], datasets: [{ name, color, data: [...] }], min: 0, max: 100 }
    },
  },
};
```

A component tag with no matching resolver (or whose resolver throws / returns `null`)
renders as a labeled placeholder box instead of failing the whole export.

## Development

```bash
npm test         # unit tests, package lint, kitchen-sink fixture export
npm run render   # LibreOffice render of the fixture (and ../innotek-SlideDevBase if present) to out/render/
```

## What it doesn't do

This is a markdown -> pptx layout engine tuned for the patterns Slidev decks commonly use,
not a full HTML/CSS renderer. Arbitrary nested HTML, custom Vue components without a
resolver, and complex CSS layouts fall back to a flattened, best-effort text rendering.

Only PNG/JPEG/GIF images and logos embed natively; an SVG image exports as a labeled
placeholder and an SVG logo is skipped. A `layout: section` slide whose body is more than
plain text/bullets (a table, chart or card row) uses the *Title Only* layout instead of
*Section Header*, since that content doesn't fit the section body placeholder.
