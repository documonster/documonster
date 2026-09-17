/**
 * Text shaping for the Word module — a re-export of the shared engine.
 *
 * The engine itself is `@utils/text-shaping`, at Layer 0, because the rasteriser in
 * `draw` needs it and `draw` cannot import `word`. See that module's header for why it
 * moved.
 *
 * This file exists so that `documonster/word`'s published surface is unchanged: it has
 * exported `shapeText`, `detectScript` and `detectDirection` from here, and a consumer's
 * import should not have to follow an internal reorganisation.
 *
 * @module
 */

export { detectDirection, detectScript, shapeText } from "@utils/text-shaping";
export type { BiDiDirection, ScriptType, ShapedCluster, ShapingOptions } from "@utils/text-shaping";
