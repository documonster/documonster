/**
 * DOCX Module - Reference Scanners
 *
 * Shared helpers for scanning the document model for hyperlink / chart
 * references (images: see `part-images.ts`) that need to be registered against a part-level .rels
 * file. Used by both the bulk packager and the streaming writer to avoid
 * behaviour drift in how header/footer/notes/comments parts collect their
 * own relationships.
 */

import { RelType } from "@word/constants";
import { sanitizeUrl } from "@word/core/internal-utils";
import { walkBlocks } from "@word/core/walker";
import type {
  BodyContent,
  ChartContent,
  HeaderFooterContent,
  Hyperlink,
  Paragraph,
  ParagraphChild,
  PreservedLink
} from "@word/types";
import { addRelationship } from "@word/writer/relationships";
import type { RelationshipsState } from "@word/writer/relationships";
import type { PreservedLinkRId } from "@word/writer/render-context";

/**
 * Recursively collect external-URL hyperlinks from a list of paragraph
 * children. Descends into track-change wrappers so an inserted hyperlink
 * isn't dropped.
 *
 * Always emits hyperlinks that carry a `url`, regardless of whether the
 * model already has an `rId`. Reader-supplied `rId`s are stale once the
 * package is repackaged because the relationship table is rebuilt from
 * scratch — the packager assigns its own canonical `rId<N>` and exposes
 * it via the `hyperlinkRIds` WeakMap, which the renderer consults.
 */
export function scanChildrenForHyperlinks(
  children: readonly ParagraphChild[],
  out: Hyperlink[]
): void {
  walkBlocks([{ type: "paragraph", children } as Paragraph], {
    enterHyperlink(h) {
      if (h.url) {
        out.push(h);
      }
    }
  });
}

/**
 * Walk paragraphs in a block list (recursing through tables, SDTs, TOC
 * cached paragraphs, text boxes) and visit each paragraph.
 */
function walkParagraphs(blocks: readonly BodyContent[], onParagraph: (p: Paragraph) => void): void {
  walkBlocks(blocks, {
    enterParagraph(p) {
      onParagraph(p);
    }
  });
}

/** Collect hyperlinks from header/footer content. */
export function collectHyperlinksFromHeaderFooter(content: HeaderFooterContent): Hyperlink[] {
  const links: Hyperlink[] = [];
  walkParagraphs(content.children as readonly BodyContent[], p => {
    scanChildrenForHyperlinks(p.children, links);
  });
  return links;
}

/** Collect all chart contents inside header/footer content. */
export function collectChartsFromHeaderFooter(
  content: HeaderFooterContent,
  out: ChartContent[]
): void {
  walkBlocks(content.children as readonly BodyContent[], {
    visitChart(chart) {
      out.push(chart);
    }
  });
}

export function collectHyperlinksFromNotes(
  notes: readonly { content: readonly Paragraph[] }[] | undefined
): Hyperlink[] {
  const links: Hyperlink[] = [];
  if (!notes) {
    return links;
  }
  for (const note of notes) {
    for (const p of note.content) {
      scanChildrenForHyperlinks(p.children, links);
    }
  }
  return links;
}

/**
 * Every hyperlink that markup kept verbatim under `root` refers to — an
 * `a:hlinkClick` on a drawing's `docPr` or a shape's `cNvPr`. Each must be
 * registered in the `.rels` of the part `root` is written into, like a
 * modelled hyperlink, and its markup rewritten to the id it gets.
 */
export function collectPreservedLinks(root: unknown): PreservedLink[] {
  const out: PreservedLink[] = [];
  const stack: unknown[] = [root];
  const seen = new Set<object>();
  while (stack.length > 0) {
    const node = stack.pop();
    if (node === null || typeof node !== "object" || ArrayBuffer.isView(node) || seen.has(node)) {
      continue;
    }
    // A model may share or even cycle through objects; visit each once.
    seen.add(node);
    if (Array.isArray(node)) {
      for (const item of node) {
        stack.push(item);
      }
      continue;
    }
    if (node instanceof Map) {
      for (const value of node.values()) {
        stack.push(value);
      }
      continue;
    }
    const record = node as Record<string, unknown>;
    for (const key of ["_frame", "_preserved"]) {
      const kept = record[key] as { links?: readonly PreservedLink[] } | undefined;
      for (const link of kept?.links ?? []) {
        if (link.url) {
          out.push(link);
        }
      }
    }
    for (const key in record) {
      const value = record[key];
      if (value !== null && typeof value === "object" && key !== "_frame" && key !== "_preserved") {
        stack.push(value);
      }
    }
  }
  return out;
}

/**
 * A {@link PreservedLinkRId} for the part whose relationships are `rels`: each
 * hyperlink that markup kept verbatim refers to is registered there, once, as
 * the markup is written — the part's `.rels` must therefore be serialised
 * after the part. A target `sanitizeUrl` rejects gets no relationship, and
 * the writer drops its id.
 */
export function preservedLinkRegistrar(rels: RelationshipsState): PreservedLinkRId {
  const ids = new WeakMap<PreservedLink, string | undefined>();
  return link => {
    if (ids.has(link)) {
      return ids.get(link);
    }
    const safe = link.url ? sanitizeUrl(link.url) : undefined;
    const id = safe ? addRelationship(rels, RelType.Hyperlink, safe, "External") : undefined;
    ids.set(link, id);
    return id;
  };
}
