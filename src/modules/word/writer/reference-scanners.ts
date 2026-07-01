/**
 * DOCX Module - Reference Scanners
 *
 * Shared helpers for scanning the document model for image / hyperlink /
 * chart references that need to be registered against a part-level .rels
 * file. Used by both the bulk packager and the streaming writer to avoid
 * behaviour drift in how header/footer/notes/comments parts collect their
 * own relationships.
 */

import { walkBlocks } from "@word/core/walker";
import type {
  BodyContent,
  ChartContent,
  HeaderFooterContent,
  Hyperlink,
  Paragraph,
  ParagraphChild,
  Run
} from "@word/types";

/** Scan a Run for image rIds. */
function scanRunForImages(run: Run, out: Set<string>): void {
  for (const rc of run.content) {
    if (rc.type === "image" && rc.rId) {
      out.add(rc.rId);
    }
  }
}

/**
 * Recursively collect image rIds referenced inside a list of paragraph
 * children. Descends into track-change wrappers (`InsertedRun`,
 * `MovedToRun`, etc.) so an image embedded inside `<w:ins>...</w:ins>` is
 * not silently lost.
 */
export function scanChildrenForImages(children: readonly ParagraphChild[], out: Set<string>): void {
  walkBlocks([{ type: "paragraph", children } as Paragraph], {
    enterRun(run) {
      scanRunForImages(run, out);
    }
  });
}

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

/** Collect all image rIds referenced in header/footer content. */
export function collectImageRidsFromContent(content: HeaderFooterContent): Set<string> {
  const rIds = new Set<string>();
  walkParagraphs(content.children as readonly BodyContent[], p => {
    scanChildrenForImages(p.children, rIds);
  });
  return rIds;
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

/** Collect image rIds and hyperlinks from a notes/comments collection. */
export function collectImageRidsFromNotes(
  notes: readonly { content: readonly Paragraph[] }[] | undefined
): Set<string> {
  const rIds = new Set<string>();
  if (!notes) {
    return rIds;
  }
  for (const note of notes) {
    for (const p of note.content) {
      scanChildrenForImages(p.children, rIds);
    }
  }
  return rIds;
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
