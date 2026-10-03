/**
 * Heading detection — the one rule every consumer of a paragraph uses.
 *
 * A paragraph's heading level is its *effective* outline level: a direct
 * `w:outlineLvl`, otherwise the nearest one in the paragraph style's `basedOn`
 * chain. Level 9 means "body text" and overrides a heading style — including a
 * built-in one, wherever in the chain the 9 is set. Only when nothing in the
 * whole chain sets an outline level does a built-in heading style
 * (`Heading1`…`Heading9` by id, `heading 1`…`heading 9` by name) or `Title`
 * stand in, the nearest one in the chain winning.
 *
 * `Title` is reported as `kind: "title"`, `level: 1`. Clamping (HTML and
 * Markdown to six levels) is left to the caller.
 */

import { indexStyles } from "@word/query/style-resolve";
import type { StyleIndex } from "@word/query/style-resolve";
import type { DocxDocument, Paragraph } from "@word/types";

/** The heading a paragraph resolves to. */
export interface HeadingInfo {
  /** Heading level, 1–9 (outline level + 1). */
  readonly level: number;
  /** `"title"` when it comes from the built-in Title style rather than an outline level. */
  readonly kind: "heading" | "title";
}

/** `w:outlineLvl` value meaning "body text — not part of the outline". */
const BODY_TEXT_OUTLINE_LEVEL = 9;

const BUILT_IN_HEADING = /^heading\s*([1-9])$/i;

/**
 * Resolve whether a paragraph is a heading, and at which level.
 *
 * @param styles - Optional prebuilt `indexStyles` index, reused across calls.
 * @returns The heading, or `undefined` for body text.
 */
export function resolveHeadingLevel(
  doc: DocxDocument,
  para: Paragraph,
  styles?: StyleIndex
): HeadingInfo | undefined {
  const direct = para.properties?.outlineLevel;
  if (direct !== undefined) {
    return fromOutlineLevel(direct);
  }
  let current = para.properties?.style;
  if (!current) {
    return undefined;
  }
  const styleMap = styles ?? indexStyles(doc);
  const chain: string[] = [];
  const visited = new Set<string>();
  while (current && !visited.has(current)) {
    visited.add(current);
    chain.push(current);
    current = styleMap.get(current)?.basedOn;
  }
  // The effective outline level is an ordinary inherited property: settle it
  // over the whole chain before consulting names, so a `Heading2` whose base
  // marks it body text (level 9) is body text.
  for (const id of chain) {
    const def = styleMap.get(id);
    const outline = def?.paragraphProperties?.outlineLevel ?? def?.outlineLevel;
    if (outline !== undefined) {
      return fromOutlineLevel(outline);
    }
  }
  for (const id of chain) {
    const builtIn = builtInHeading(id) ?? builtInHeading(styleMap.get(id)?.name);
    if (builtIn) {
      return builtIn;
    }
  }
  return undefined;
}

function fromOutlineLevel(outlineLevel: number): HeadingInfo | undefined {
  if (
    !Number.isInteger(outlineLevel) ||
    outlineLevel < 0 ||
    outlineLevel >= BODY_TEXT_OUTLINE_LEVEL
  ) {
    return undefined;
  }
  return { level: outlineLevel + 1, kind: "heading" };
}

function builtInHeading(name: string | undefined): HeadingInfo | undefined {
  // A hand-built StyleDef may omit its name despite the type.
  if (typeof name !== "string") {
    return undefined;
  }
  const match = BUILT_IN_HEADING.exec(name.trim());
  if (match) {
    return { level: Number(match[1]), kind: "heading" };
  }
  if (name.trim().toLowerCase() === "title") {
    return { level: 1, kind: "title" };
  }
  return undefined;
}
