/**
 * Document Split API
 *
 * Split a single DocxDocument into multiple documents based on
 * sections, page breaks, or heading levels.
 */

import type { DocxDocument, BodyContent, Paragraph } from "../types";

// =============================================================================
// Public API
// =============================================================================

/** Options for splitting a document. */
export interface SplitOptions {
  /**
   * Split criteria.
   *
   * - `"section"` — split at every section break (paragraph with `sectionProperties`).
   * - `"pageBreak"` — split at every explicit page break (run with `breakType: "page"`).
   * - `"heading"` — split at every Heading 1 paragraph (or `headingLevel`).
   *
   * Default: `"section"`.
   */
  readonly by?: "section" | "pageBreak" | "heading";
  /**
   * Heading level to split on (only used when `by: "heading"`).
   * Default: 1 (Heading 1).
   */
  readonly headingLevel?: number;
  /**
   * Whether each split document keeps the original document's
   * styles, numbering, settings, fonts, etc.
   * Default: true.
   */
  readonly preserveSharedParts?: boolean;
}

/**
 * Split a DocxDocument into multiple documents.
 *
 * Each resulting document is a complete, valid DocxDocument with body content
 * from one segment of the original. Headers, footers, footnotes, endnotes,
 * comments, styles, numbering, and settings are preserved in each split unless
 * `preserveSharedParts: false` is specified.
 *
 * @param doc - The document to split.
 * @param options - Split criteria.
 * @returns Array of split documents (at least 1).
 */
export function splitDocument(doc: DocxDocument, options?: SplitOptions): DocxDocument[] {
  const opts: Required<SplitOptions> = {
    by: options?.by ?? "section",
    headingLevel: options?.headingLevel ?? 1,
    preserveSharedParts: options?.preserveSharedParts ?? true
  };

  const segments = splitBody(doc.body as readonly BodyContent[], opts);

  if (segments.length === 0) {
    return [doc];
  }

  return segments.map(segment => buildSplitDoc(doc, segment, opts));
}

// =============================================================================
// Internal
// =============================================================================

/**
 * Split the body into segments based on the split criteria.
 * Returns an array of segments; each segment is an array of body content blocks.
 */
function splitBody(body: readonly BodyContent[], opts: Required<SplitOptions>): BodyContent[][] {
  const segments: BodyContent[][] = [];
  let current: BodyContent[] = [];

  for (const block of body) {
    if (shouldSplitBefore(block, opts)) {
      // Heading-based split or pageBreakBefore: start a new segment BEFORE this block
      if (current.length > 0) {
        segments.push(current);
      }
      current = [block];
    } else {
      current.push(block);
      if (shouldSplitAfter(block, opts)) {
        segments.push(current);
        current = [];
      }
    }
  }

  if (current.length > 0) {
    segments.push(current);
  }

  return segments;
}

function shouldSplitBefore(block: BodyContent, opts: Required<SplitOptions>): boolean {
  if (block.type !== "paragraph") {
    return false;
  }
  if (opts.by === "heading") {
    return isHeadingLevel(block, opts.headingLevel);
  }
  if (opts.by === "pageBreak") {
    // pageBreakBefore semantically means "this paragraph starts on a new page".
    // Only the explicit run page break stays as an "after" split.
    return block.properties?.pageBreakBefore === true;
  }
  return false;
}

function shouldSplitAfter(block: BodyContent, opts: Required<SplitOptions>): boolean {
  if (block.type !== "paragraph") {
    return false;
  }
  if (opts.by === "section") {
    return block.properties?.sectionProperties !== undefined;
  }
  if (opts.by === "pageBreak") {
    return paragraphHasExplicitPageBreakRun(block);
  }
  return false;
}

function isHeadingLevel(para: Paragraph, level: number): boolean {
  const style = para.properties?.style;
  if (!style) {
    return false;
  }
  // Common heading style IDs: "Heading1", "heading 1", "Heading%i", etc.
  const target = String(level);
  const normalized = style.replace(/\s+/g, "").toLowerCase();
  return (
    normalized === `heading${target}` ||
    normalized === `h${target}` ||
    normalized === `title${target}`
  );
}

function paragraphHasExplicitPageBreakRun(para: Paragraph): boolean {
  for (const child of para.children) {
    if (!("type" in child) && "content" in child && Array.isArray(child.content)) {
      for (const c of child.content) {
        if (c.type === "break" && c.breakType === "page") {
          return true;
        }
      }
    }
  }
  return false;
}

function buildSplitDoc(
  source: DocxDocument,
  segment: BodyContent[],
  opts: Required<SplitOptions>
): DocxDocument {
  if (!opts.preserveSharedParts) {
    // Minimal split: just body + docType
    return {
      docType: source.docType,
      body: segment
    };
  }

  // Full split: preserve all shared parts
  return {
    ...source,
    body: segment
  };
}
