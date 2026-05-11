/** @module Sub-path entry point for direct imports: `import { ... } from "excelts/word/font"` */

export { subsetFont, embedFont, embedFontFamily, addEmbeddedFonts } from "./font-embed";
export type { FontEmbedStyle, EmbedFontOptions, EmbedFontResult } from "./font-embed";
export { deobfuscateFont, obfuscateFont, generateFontKey } from "./font-obfuscation";
export { shapeText, detectScript, detectDirection } from "./text-shaping";
export type { ShapedCluster, ShapingOptions, ScriptType, BiDiDirection } from "./text-shaping";
export { createHyphenator, hyphenateWord, hyphenateText, ENGLISH_US_PATTERNS } from "./hyphenation";
export type { HyphenationOptions, HyphenationPatterns } from "./hyphenation";
