// Generic, deck-agnostic defaults. A consuming deck overrides whichever of these it needs
// (usually colors/fonts/logo and chartResolvers for its own Vue components) via a
// `pptx.config.mjs` next to its slides.md - see mergeConfig() in ./config.mjs and the README.

export const defaultConfig = {
  // PowerPoint document metadata author field.
  author: "",

  slide: {
    width: 13.333,
    height: 7.5,
    marginX: 0.6,
  },

  colors: {
    accent: "2563EB",
    accentDark: "1E3A8A",
    accent2: "7C3AED",
    panel: "F1F5F9",
    border: "E2E8F0",
    text: "1F2937",
    muted: "6B7280",
  },

  fonts: {
    heading: "Arial",
    body: "Arial",
    mono: "Courier New",
  },

  // Maps a CSS custom-property name - as it appears in an inline style="color: var(--x)"
  // attribute in the deck's markdown - to one of the keys in `colors` above. Only decks that
  // color text via CSS variables (rather than Tailwind/UnoCSS classes) need to populate this.
  cssVarColors: {},

  // Class name a deck can put on a <div> to render as the small uppercase eyebrow label
  // above a slide's title. Not a Slidev built-in - just a convention this tool recognizes.
  eyebrowClass: "eyebrow",

  // Optional footer logo image drawn bottom-left on every slide (skipped if null or the
  // file doesn't exist). `path` is resolved relative to the config file itself.
  logo: null, // { path: string, aspect: number, height?: number }

  // Directory chart resolvers may read data files from, relative to the input .md file.
  dataDir: "public/data",

  // Maps a `<ComponentName />` tag as it appears in the markdown to a resolver function
  // `(dataDirAbsolutePathOrNull) => ResolvedChart | null`. A component with no matching key
  // (or whose resolver throws / returns null) renders as a labeled placeholder box.
  //
  // ResolvedChart is one of:
  //   { kind: "statcards", cards: [{ title?, accent?, big?, subtitle?, stats?, footnote? }] }
  //   { kind: "barchart-panels", panels: [{ title, color, rows: [{ model, success }] }], min, max }
  //   { kind: "barchart", labels: string[], datasets: [{ name, color, data: number[] }], min, max }
  chartResolvers: {},
};
