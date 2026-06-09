/**
 * Document Split API
 *
 * Split a single DocxDocument into multiple documents based on
 * sections, page breaks, or heading levels.
 */

import { isRun } from "../core/text-utils";
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
  const props = para.properties;
  // outlineLevel mirrors getHeadings()'s detection: outlineLevel 0 == H1,
  // outlineLevel 1 == H2, etc. Levels >= 9 mean "body text" and should not
  // qualify as headings.
  if (props?.outlineLevel !== undefined && props.outlineLevel < 9) {
    if (props.outlineLevel + 1 === level) {
      return true;
    }
  }
  const style = props?.style;
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
    if (isRun(child)) {
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
  // When splitting by section, the segment's last paragraph carries the
  // section break (often `nextPage`) that originally separated it from the
  // following section. In a standalone split document that break has nothing
  // after it, so Word renders a trailing blank page. Strip the paragraph-level
  // sectPr and promote its page setup to the document's own section
  // properties instead.
  let body = segment;
  let promotedSectPr: DocxDocument["sectionProperties"] | undefined;
  if (opts.by === "section" && segment.length > 0) {
    const last = segment[segment.length - 1];
    if (last.type === "paragraph" && last.properties?.sectionProperties) {
      const para = last as Paragraph;
      const props = para.properties!;
      const sectionProperties = props.sectionProperties!;
      const restProps = { ...props };
      delete (restProps as { sectionProperties?: unknown }).sectionProperties;
      // Drop the inter-section break type: in a standalone document this is
      // the (only/first) section, so a "nextPage" break would just push all
      // content onto page 2, leaving page 1 blank.
      const { breakType: _drop, ...sectWithoutBreak } = sectionProperties;
      promotedSectPr = sectWithoutBreak;
      const cleaned: Paragraph = { ...para, properties: restProps };
      body = [...segment.slice(0, -1), cleaned];
    }
  }

  if (!opts.preserveSharedParts) {
    // Minimal split: just body + docType
    return {
      docType: source.docType,
      body,
      ...(promotedSectPr ? { sectionProperties: promotedSectPr } : {})
    };
  }

  // Full split: preserve all shared parts
  return {
    ...source,
    body,
    ...(promotedSectPr ? { sectionProperties: promotedSectPr } : {})
  };
}
