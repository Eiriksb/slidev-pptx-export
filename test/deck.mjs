// DeckOptions shared by the writer tests (Tasks 8-11).
import { fileURLToPath } from "node:url";
import { slideGeometry } from "../src/geometry.mjs";
import { themeColors } from "../src/colors.mjs";

export const PIC = fileURLToPath(new URL("./fixtures/kitchen-sink/public/pic.png", import.meta.url));
export const geometry = slideGeometry({ width: 13.333, height: 7.5, marginX: 0.6 });

export const DECK = {
  geometry,
  theme: themeColors({ accent: "1A833B", accentDark: "14532D", accent2: "1A589F", panel: "EEF4F0", border: "D9E3DB", text: "1F2937", muted: "6B7280" }),
  fonts: { heading: "Plus Jakarta Sans", body: "Inter", mono: "JetBrains Mono" },
  lang: "en-US",
  title: "Writer test",
  author: "Tester",
  timestamps: false,
  logo: { path: PIC, box: { x: 0.5, y: 6.94, w: 0.54, h: 0.27 } },
};
