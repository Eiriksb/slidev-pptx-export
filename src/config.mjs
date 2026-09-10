import { defaultConfig } from "./default-config.mjs";

// Shallow-merges each top-level section (colors/fonts/cssVarColors/chartResolvers get their
// keys unioned with the defaults; slide/logo/eyebrowClass/dataDir get replaced wholesale if
// the user config sets them) so a deck only has to specify what it wants to change.
export function mergeConfig(userConfig = {}) {
  return {
    ...defaultConfig,
    ...userConfig,
    slide: { ...defaultConfig.slide, ...userConfig.slide },
    colors: { ...defaultConfig.colors, ...userConfig.colors },
    fonts: { ...defaultConfig.fonts, ...userConfig.fonts },
    cssVarColors: { ...defaultConfig.cssVarColors, ...userConfig.cssVarColors },
    chartResolvers: { ...defaultConfig.chartResolvers, ...userConfig.chartResolvers },
    logo: userConfig.logo !== undefined ? userConfig.logo : defaultConfig.logo,
  };
}
