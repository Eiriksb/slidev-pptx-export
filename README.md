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

// Render mode (see below): runs the deck's own Slidev in headless Chromium.
await convert({ inputPath: "slides.md", outputPath: "slides-editable.pptx", render: true });
```

`render: { screenshots: true }` also returns Slidev's own screenshot of each exported slide as
`result.screenshots` (`null` for a slide exported from markdown), which the visual check uses.

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
- Boxes with a background, border (solid, dashed or dotted; a single dashed side is a native line),
  rounded corners, shadow or gradient (cards, callouts) become native shapes; lists get native bullets in their rendered colors; tables become
  native tables with each cell's own fill and borders; code blocks keep their syntax colors
  (a block that scrolls or wraps in the page is a picture instead).
- Chart.js bar charts (e.g. through vue-chartjs) become native, editable charts with the
  page's data, colors, axis range and number format; horizontal bars keep their order. Only
  charts PowerPoint can draw faithfully are native: for example, other chart types, mixed
  types, a log or second value axis, stack groups, a chart title, per-bar or scripted colors,
  and rotated or translucent charts are exported as pictures. A `chartResolvers` entry
  returning `kind: "barchart"` for the component still wins, also over a chart that would
  have been a picture.
- Videos become native PowerPoint videos: the file (MP4 or MOV) is embedded, cropped or letterboxed as
  the page draws it (`object-fit`), with its `poster` (else the frame shown) as the picture before it plays.
  `autoplay` (the attribute, or `<SlidevVideo autoplay>`) starts it with the slide, `loop` repeats it until
  the slide is left and `muted` mutes it, the way PowerPoint writes *Start: Automatically* and *Loop until
  stopped*; without autoplay it plays when clicked. A video that is transformed, filtered, masked, clipped or
  in another format (WebM) stays a picture of its current frame, with a notice.
- `::before` / `::after` boxes and their text become native shapes and text boxes: a coloured square
  before a label, a numbered circle before a step (`counter()`, `counters()`, `list-item`, `attr()` and
  quoted strings; decimal, alpha and roman numbering), placed by `position: absolute` or as the flow puts
  them. An unsized pseudo-element with only text in the flow is left to the text around it.
- Images become native pictures; one that didn't load is skipped, with a notice. SVG images
  stay vector in PowerPoint 2016+ (with a PNG fallback elsewhere) when their box has the
  SVG's aspect, and are pictures otherwise. An image rounded with one radius all round (its own, or that of a
  box it fills which clips it) is a picture with a rounded-rectangle shape; other rounding (some corners,
  an image in one corner of a rounded box) is a picture of the page's drawing.
- A `url()` background is native only when it is centred cover/contain or 100% 100%. Any
  other background PowerPoint can't draw - an SVG, tiled or off-centre image, a translucent
  one, layered backgrounds, a repeating gradient, a radial gradient whose shape differs a lot
  from its box - becomes a picture behind the element's contents, which stay native text
  and shapes.
- `global-top.vue` and `global-bottom.vue` are exported with every slide (the bottom layer
  only where the slide lets it show through), so a shared footer usually lands on the slide
  master. A slide's `zoom:` is honoured: its text and shapes keep their zoomed size.
- Elements drawn identically on every slide, like a footer logo, go on the slide master once,
  unless slide content covers them. A picture there is reported as "slide master".
- Anything PowerPoint can't express natively (blur and other CSS filters, masks, inline SVG
  drawings, other canvases, and the cases above) is drawn as a picture of just that element,
  with a notice naming the slide.
- `colors` and `fonts` in `pptx.config.mjs` still build the theme (rendered colors that match
  become theme colors); `logo`, `eyebrowClass` and `cssVarColors` aren't needed, and a notice
  says they are ignored when the config sets them. `render: { timeout }` sets the per-slide
  wait (default 15000 ms). The first slide also waits for Slidev's cold start (Vite
  dependency optimisation), so slow machines or heavy decks may need a larger timeout.
- A slide Slidev can't load or render falls back to the markdown export, with a warning. A
  slide imported from another file (`src:`) has no markdown fallback (the default mode
  doesn't read imported files), so it is left blank, with a warning. When Slidev can't start
  (for example its theme isn't installed), Chromium can't launch or the Slidev version isn't
  supported, the export stops with an error instead; Chromium is launched before Slidev
  starts, so a missing browser fails fast.
- Limits: each slide is captured in its final `v-click` state, with every fragment shown.
  Animations and transitions (videos excepted, see above), hyperlinks (link text keeps its styling) and Slidev
  drawings aren't carried over; fonts only with `--embed-fonts` (see Fonts).

Render mode needs the deck's dependencies installed (`npm install` in the deck folder), Node
20.12 or newer (what Slidev 52 requires) and Playwright's Chromium. It looks for
`playwright-chromium`, then `playwright`, each first in the deck folder and then next to this
package. It takes roughly 10–30 seconds per deck.

## Fonts (`--embed-fonts`)

```bash
npx slidev-pptx-export slides.md slides.pptx --render --embed-fonts
```

The fonts the slides use are carried in the `.pptx`, so it looks the same on a machine without them:
regular, bold, italic and bold italic of each family, as Embedded OpenType (`/ppt/fonts/*.fntdata`, what
PowerPoint itself writes), listed in the presentation's `embeddedFontLst`.

- Each font name is looked up with fontconfig (`fc-match`). A web font's name that no installed font has
  ("Inter Variable", as `@fontsource-variable` calls Inter) falls back to the family without " Variable",
  and the slides are written with that family's name, the one inside the embedded font.
- Only TrueType fonts whose licence allows embedding (OS/2 `fsType` not restricted) are embedded;
  PostScript-outline (CFF) fonts and restricted ones are left out with a notice, as are names no installed
  font has (like Arial on Linux: a theme default every PowerPoint machine has).
- Fonts that aren't installed can be given in the config:

```js
export default {
  embedFonts: {
    aliases: { "JetBrains Mono Variable": "JetBrains Mono" },
    files: { "JetBrains Mono": { regular: "fonts/JetBrainsMono-Regular.ttf", bold: "fonts/JetBrainsMono-Bold.ttf" } },
  },
};
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
npm test                    # unit tests, package lint, kitchen-sink fixture export
npm run render              # LibreOffice render of the fixture (and ../innotek-SlideDevBase if present) to out/render/
npm run test:render         # browser tests (needs playwright-chromium; skips otherwise)
npm run render -- --render  # render-mode export + visual check against Slidev's own screenshots
```

The browser tests (`test/render/`) use `playwright-chromium` and `chart.js`, both
devDependencies; `chart.js` is only there to draw test charts. `npm test` needs no
browser. The render checks skip when LibreOffice, poppler or Chromium isn't available.

## What it doesn't do

The default mode is a markdown -> pptx layout engine tuned for the patterns Slidev decks
commonly use, not a full HTML/CSS renderer. Arbitrary nested HTML, custom Vue components
without a resolver, and complex CSS layouts fall back to a flattened, best-effort text
rendering. `--render` converts what the page draws instead, so custom HTML, CSS and Vue
components come through (see Render mode above for its limits).

In the default mode only PNG/JPEG/GIF images and logos embed natively; an SVG image exports
as a labeled placeholder and an SVG logo is skipped (`--render` embeds SVG images, see
above). A `layout: section` slide whose body is more than plain text/bullets (a table, chart
or card row) uses the *Title Only* layout instead of *Section Header*, since that content
doesn't fit the section body placeholder.
