/**
 * Text shaping for the Word module — the shared engine behind `Font.*`.
 *
 * The engine itself is `@utils/text-shaping`, at Layer 0, because the rasteriser in
 * `draw` needs it and `draw` cannot import `word`. See that module's header for why it
 * moved.
 *
 * This file exists so that `documonster/word`'s published surface is unchanged: it has
 * exported `shapeText`, `detectScript` and `detectDirection` from here, and a consumer's
 * import should not have to follow an internal reorganisation. It checks their arguments
 * as Word's other public functions do, so a wrong one is a `DocxArgumentError` rather
 * than a `TypeError` from inside the engine.
 *
 * @module
 */

import type { SchemaMatches } from "@utils/schema";
import type { BiDiDirection, ScriptType, ShapedCluster, ShapingOptions } from "@utils/text-shaping";
import {
  detectDirection as detectDirectionImpl,
  detectScript as detectScriptImpl,
  shapeText as shapeTextImpl
} from "@utils/text-shaping";
import { checkString } from "@word/core/argument-checks";
import { checkSchema, w } from "@word/core/option-schema";

export type { BiDiDirection, ScriptType, ShapedCluster, ShapingOptions };

const shapingOptionsSchema = w.record({
  direction: w.oneOf(["ltr", "rtl"]).optional(),
  arabicJoining: w.boolean().optional(),
  bidiReorder: w.boolean().optional()
});
const _shapingOptionsMatch: SchemaMatches<typeof shapingOptionsSchema, ShapingOptions> = true;
void _shapingOptionsMatch;

/**
 * Shape a text string for complex script rendering.
 * Returns an array of shaped clusters in visual order.
 *
 * @param text - The input text in logical order.
 * @param options - Shaping options.
 * @returns Array of shaped clusters.
 */
export function shapeText(text: string, options?: ShapingOptions): ShapedCluster[] {
  const where = "Font.shapeText";
  checkString(text, "text", where);
  checkSchema(shapingOptionsSchema.optional(), options, where);
  return shapeTextImpl(text, options);
}

/** Detect the dominant script of a text string. */
export function detectScript(text: string): ScriptType {
  checkString(text, "text", "Font.detectScript");
  return detectScriptImpl(text);
}

/** Determine the BiDi direction of a text based on its first strong character. */
export function detectDirection(text: string): BiDiDirection {
  checkString(text, "text", "Font.detectDirection");
  return detectDirectionImpl(text);
}
