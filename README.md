# slidev-pptx-export

Converts a [Slidev](https://sli.dev) markdown deck into a fully editable native `.pptx` -
real PowerPoint text boxes, tables, shapes and charts, not the screenshot-per-slide export
that `slidev export` produces.

It understands plain markdown (headings, paragraphs, lists, tables, blockquotes, images,
fenced code blocks, `**bold**`/`*italic*`/`` `code` ``) plus a handful of common Slidev
authoring patterns: frontmatter-separated slides, Tailwind/UnoCSS grid layouts (`grid
grid-cols-N` -> a row of cards), `text-2xl`+ stat numbers, `opacity-NN` fades, and `<Vue
Component />` tags resolved into native charts via a `chartResolvers` config.

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

  // Inline style="color: var(--x)" -> one of the color keys above, for decks that color
  // text via CSS variables rather than Tailwind classes.
  cssVarColors: { "--brand-blue": "accent2" },

  // <div class="..."> that renders as the small uppercase label above a slide's title.
  eyebrowClass: "eyebrow",

  // Footer logo drawn bottom-left on every slide, or omit/null for none.
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

## What it doesn't do

This is a markdown -> pptx layout engine tuned for the patterns Slidev decks commonly use,
not a full HTML/CSS renderer. Arbitrary nested HTML, custom Vue components without a
resolver, and complex CSS layouts fall back to a flattened, best-effort text rendering.
