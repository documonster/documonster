/**
 * `Font` namespace surface — font embedding / subsetting, text shaping,
 * hyphenation.
 *
 * `import { Font } from "documonster/word"` →
 *   `Font.embed(opts)`, `Font.subset(data, chars)`, `Font.shapeText(...)`,
 *   `Font.hyphenateWord(...)`, … — tree-shaken via `export * as Font`.
 */
export {
  embedFont as embed,
  embedFontFamily as embedFamily,
  addEmbeddedFonts as addEmbedded,
  subsetFont as subset
} from "@word/font/font-embed";
export { shapeText, detectScript, detectDirection } from "@word/font/text-shaping";
export {
  createHyphenator,
  hyphenateWord,
  hyphenateText,
  ENGLISH_US_PATTERNS
} from "@word/font/hyphenation";
