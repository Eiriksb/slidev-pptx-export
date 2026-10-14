// Config for the kitchen-sink fixture deck: one resolver per chart kind the exporter supports.
import { readFileSync } from "node:fs";
import { join } from "node:path";

export default {
  author: "Fixture",
  colors: { accent: "1A833B", accentDark: "14532D", accent2: "1A589F", panel: "EEF4F0", border: "D9E3DB", text: "1F2937", muted: "6B7280" },
  fonts: { heading: "Plus Jakarta Sans", body: "Inter", mono: "JetBrains Mono" },
  logo: { path: "./public/pic.png", aspect: 2, height: 0.27 },
  dataDir: "public/data",
  chartResolvers: {
    FixtureChart(dataDir) {
      const d = JSON.parse(readFileSync(join(dataDir, "chart.json"), "utf8"));
      return { kind: "barchart", labels: d.labels, datasets: d.datasets, min: d.min, max: d.max };
    },
    FixtureStats: () => ({
      kind: "statcards",
      cards: [
        { title: "Speed", stats: [{ value: "3.4s", label: "p50" }, { value: "9.1s", label: "p95" }] },
        { title: "Quality", big: { value: "72%", label: "pass rate" }, footnote: "n = 120" },
      ],
    }),
    FixturePanels: () => ({
      kind: "barchart-panels",
      panels: [
        { title: "Model A", color: "1A833B", rows: [{ model: "alpha", success: 80 }, { model: "beta", success: 60 }] },
        { title: "Model B", color: "1A589F", rows: [{ model: "alpha", success: 70 }] },
      ],
      min: 0,
      max: 100,
    }),
  },
};
