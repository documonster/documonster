/**
 * DOCX Module - Patcher
 *
 * Internal logic for replacing placeholders (e.g. `{{name}}`) in a DocxDocument
 * with text, paragraphs, tables, or images. Used by `patchDocument` and
 * `patchTemplate` in document-io.ts.
 *
 * Public API is exposed through document-io.ts; this file contains the
 * pure patch logic (no IO).
 *
 * Container coverage: the patcher walks every container that may carry
 * paragraphs — body, headers/footers, footnotes/endnotes, table cells
 * (including nested tables), structured-document tags (block & inline),
 * text boxes, drawing shape text bodies, and the cached paragraphs of a
 * Table of Contents. Anything paragraph-shaped is run through
 * `patchParagraph`, so a placeholder is matched the same way regardless
 * of which container holds it.
 */

import { type Mutable } from "./core/internal-utils";
import { isHyperlink, isRun } from "./core/text-utils";
import type {
  BodyContent,
  DocxDocument,
  DrawingShape,
  HeaderFooterContent,
  ImageDef,
  InlineImageContent,
  Paragraph,
  ParagraphChild,
  Run,
  StructuredDocumentTag,
  Table,
  TableCell,
  TableOfContents,
  TextBox,
  TextContent
} from "./types";

// =============================================================================
// Public types
// =============================================================================

/** Type of content to patch into a placeholder. */
export type PatchContent =
  | { readonly type: "text"; readonly text: string }
  | { readonly type: "paragraph"; readonly children: readonly Paragraph[] }
  | { readonly type: "table"; readonly table: Table }
  | {
      readonly type: "image";
      readonly image: ImageDef;
      readonly width: number;
      readonly height: number;
    };

/** A single patch operation mapping a placeholder to replacement content. */
export interface PatchOperation {
  /** Placeholder string to find (e.g. "{{name}}"). */
  readonly placeholder: string;
  /** Content to replace the placeholder with. */
  readonly content: PatchContent;
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Apply patches to a parsed document model.
 *
 * Returns a NEW {@link DocxDocument} with a fresh `body` array and a new
 * `images` array. However, the inner content reachable from
 * `headers` / `footers` / `footnotes` / `endnotes` and from non-replaced
 * paragraphs/tables is mutated in place: text replacements rewrite the
 * existing `Paragraph.children[].content[].text`, table rows/cells are
 * patched in place, etc.
 *
 * **Practical consequence**: do not keep using the input `doc` after calling
 * this function — its inner state has been modified. Always use the return
 * value. Concurrent use of the input `doc` from another thread/path is not
 * supported.
 *
 * @param doc - The document model. Internal arrays/objects will be mutated.
 * @param patches - Patches to apply.
 * @returns A new {@link DocxDocument} reference with patches applied.
 */
export function applyPatchesToDocument(
  doc: DocxDocument,
  patches: readonly PatchOperation[]
): DocxDocument {
  // Build a canonical fileName → rId map so two image patches that point to
  // the same fileName always emit the same r:embed. Without this, the first
  // patch wins doc.images de-duplication below but a later patch with a
  // different rId would still write its own rId into the body, producing a
  // dangling reference and a blank image in Word.
  const imageFileNameToRId = new Map<string, string>();
  if (doc.images) {
    for (const img of doc.images) {
      if (img.rId) {
        imageFileNameToRId.set(img.fileName, img.rId);
      }
    }
  }
  // Normalize image patches: assign each unique fileName a single rId
  // (preferring an existing one in doc.images, otherwise the first patch's
  // rId, otherwise a stable generated id). Rewrite all subsequent patches
  // that share the fileName to use the same rId. The body-render path below
  // reads `image.rId` directly from these normalized patches, so generated
  // ids and existing ones flow through the same code path.
  const normalizedPatches: PatchOperation[] = patches.map(patch => {
    if (patch.content.type !== "image") {
      return patch;
    }
    const img = patch.content.image;
    let rId = imageFileNameToRId.get(img.fileName);
    if (!rId) {
      rId = img.rId ?? `rId_img_${img.fileName}`;
      imageFileNameToRId.set(img.fileName, rId);
    }
    if (img.rId === rId) {
      return patch;
    }
    return {
      ...patch,
      content: {
        ...patch.content,
        image: { ...img, rId }
      }
    } satisfies PatchOperation;
  });

  // Build lookup map for quick placeholder matching
  const patchMap = new Map<string, PatchOperation>();
  for (const patch of normalizedPatches) {
    patchMap.set(patch.placeholder, patch);
  }

  // Process body content (top-level)
  const newBody = patchBlockList(doc.body, patchMap);

  // Patch headers
  if (doc.headers) {
    for (const [, headerDef] of doc.headers) {
      patchHeaderFooterContent(headerDef.content, patchMap);
    }
  }

  // Patch footers
  if (doc.footers) {
    for (const [, footerDef] of doc.footers) {
      patchHeaderFooterContent(footerDef.content, patchMap);
    }
  }

  // Patch footnotes (text-only — structural patches don't make sense in notes)
  if (doc.footnotes) {
    for (const note of doc.footnotes) {
      if (note.id <= 0) {
        continue;
      }
      for (const para of note.content) {
        patchTextInParagraph(para, patchMap);
      }
    }
  }

  // Patch endnotes
  if (doc.endnotes) {
    for (const note of doc.endnotes) {
      if (note.id <= 0) {
        continue;
      }
      for (const para of note.content) {
        patchTextInParagraph(para, patchMap);
      }
    }
  }

  // Add any new images from patches. Patches were already normalized so
  // that two patches sharing a fileName carry the same rId; this loop just
  // unions the unique images.
  const images = doc.images ? [...doc.images] : [];
  for (const patch of normalizedPatches) {
    if (patch.content.type === "image") {
      const imgContent = patch.content;
      const existing = images.find(i => i.fileName === imgContent.image.fileName);
      if (!existing) {
        images.push(imgContent.image);
      }
    }
  }

  return {
    ...doc,
    body: newBody,
    images: images.length > 0 ? images : undefined
  };
}

// =============================================================================
// Internal helpers
// =============================================================================

/**
 * Iterate every visible Run in a paragraph (in document order), descending
 * into hyperlinks and into the wrapper around `insertedRun` / `movedToRun`
 * track-change markers. `deletedRun` / `movedFromRun` are intentionally
 * skipped: by the OOXML conventions used elsewhere in this codebase
 * (search.ts, replace.ts, extractParagraphText) those wrappers do not
 * contribute to the document's visible text, so a `{{name}}` placeholder
 * sitting inside a pending deletion must not be replaced.
 */
function forEachVisibleRun(para: Paragraph, fn: (run: Run) => void): void {
  const visit = (children: readonly ParagraphChild[]): void => {
    for (const child of children) {
      if (isHyperlink(child)) {
        visit(child.children as readonly ParagraphChild[]);
        continue;
      }
      if ("type" in child) {
        // Track-change wrappers: descend into the wrapped run only when
        // the wrapper represents visible text (insert / move-to). Pending
        // deletions and move-from sources, bookmarks, and comment range
        // markers contribute no visible text and are skipped.
        const t = (child as { type?: string }).type;
        if (t === "insertedRun" || t === "movedToRun") {
          const inner = (child as { run?: Run }).run;
          if (inner) {
            fn(inner);
          }
        }
        continue;
      }
      // No `type` discriminator → the only paragraph child shape that
      // matches is a Run.
      if (isRun(child)) {
        fn(child);
      }
    }
  };
  visit(para.children);
}

/**
 * Extract concatenated plain text from a paragraph's runs, ignoring tabs,
 * breaks, hyphens, fields, and any non-text run content.
 *
 * Intentionally **not** the same as `extractParagraphText` from
 * `core/text-utils`: that helper expands `tab` → "\t", `break` → "\n",
 * `noBreakHyphen` → "-", etc., which would corrupt placeholder matching.
 * A placeholder like `{{name}}` should only match against the literal text
 * the author wrote, not against synthetic characters injected by formatting
 * elements.
 *
 * Visits hyperlinks and tracked-insert/move-to wrappers so that a
 * placeholder embedded in a hyperlink display text or a pending revision
 * is still found.
 */
function paragraphText(para: Paragraph): string {
  let t = "";
  forEachVisibleRun(para, run => {
    for (const c of run.content) {
      if (c.type === "text") {
        t += (c as TextContent).text;
      }
    }
  });
  return t;
}

/** Replace text within a single paragraph (mutates run-text in place). */
function replaceInParagraph(para: Paragraph, search: string, replacement: string): void {
  // Pass 1: in-segment replacement (the placeholder lives wholly inside
  // one text node).
  forEachVisibleRun(para, run => {
    for (const c of run.content) {
      if (c.type !== "text") {
        continue;
      }
      const tc = c as Mutable<TextContent>;
      if (tc.text.includes(search)) {
        tc.text = tc.text.replaceAll(search, replacement);
      }
    }
  });

  // Pass 2: cross-run / cross-segment fallback. If after pass 1 the
  // concatenated paragraph text still contains the search string, the
  // placeholder must have straddled multiple text nodes. Concatenate the
  // whole paragraph, do the replacement, then write the new text into the
  // FIRST visible text node and clear the rest. This is the same strategy
  // the original implementation used; we just delegate text-node iteration
  // to forEachVisibleRun so hyperlinked / tracked-insert runs participate.
  const fullText = paragraphText(para);
  if (!fullText.includes(search)) {
    return;
  }
  const newText = fullText.replaceAll(search, replacement);
  let placed = false;
  forEachVisibleRun(para, run => {
    for (const c of run.content) {
      if (c.type !== "text") {
        continue;
      }
      const tc = c as Mutable<TextContent>;
      if (!placed) {
        tc.text = newText;
        placed = true;
      } else {
        tc.text = "";
      }
    }
  });
}

/** Patch a paragraph — returns replacement content or null to remove. */
function patchParagraph(
  para: Paragraph,
  patchMap: Map<string, PatchOperation>
): BodyContent | BodyContent[] | null {
  let text = paragraphText(para);

  // First pass: apply all text replacements (can have multiple in same paragraph)
  for (const [placeholder, patch] of patchMap) {
    if (patch.content.type !== "text") {
      continue;
    }
    if (!text.includes(placeholder)) {
      continue;
    }
    replaceInParagraph(para, placeholder, patch.content.text);
    text = paragraphText(para); // Re-read after replacement
  }

  // Second pass: structural replacements (paragraph/table/image replace entire paragraph)
  text = paragraphText(para);
  for (const [placeholder, patch] of patchMap) {
    if (patch.content.type === "text") {
      continue;
    }
    if (!text.includes(placeholder)) {
      continue;
    }

    switch (patch.content.type) {
      case "paragraph": {
        return patch.content.children as BodyContent[];
      }
      case "table": {
        return patch.content.table;
      }
      case "image": {
        // The patch was normalized in applyPatchesToDocument so image.rId
        // is guaranteed populated and consistent with other patches that
        // share the same fileName. Read it directly — never re-derive it
        // here, or rId generation could drift between the body reference
        // and the registered relationship.
        const img = patch.content.image;
        const rId = img.rId!;
        const imgContent: InlineImageContent = {
          type: "image",
          rId,
          width: patch.content.width,
          height: patch.content.height,
          altText: img.fileName,
          name: img.fileName
        };
        const newPara: Paragraph = {
          type: "paragraph",
          properties: para.properties,
          children: [{ content: [imgContent] } as Run]
        };
        return newPara;
      }
    }
  }

  return para;
}

/**
 * Patch a list of body content blocks, recursing into every container that
 * may hold paragraphs. Returns a new array (paragraphs that get replaced
 * by `paragraph` patches expand into multiple blocks) but mutates inner
 * tables/SDTs/text-boxes/etc. in place.
 */
function patchBlockList(
  blocks: readonly BodyContent[],
  patchMap: Map<string, PatchOperation>
): BodyContent[] {
  const out: BodyContent[] = [];
  for (const block of blocks) {
    const replaced = patchBlock(block, patchMap);
    if (replaced === null) {
      continue;
    }
    if (Array.isArray(replaced)) {
      out.push(...replaced);
    } else {
      out.push(replaced);
    }
  }
  return out;
}

/**
 * Patch a single block. Paragraphs may resolve to a single block, an array
 * of blocks (when the placeholder maps to `{ type: "paragraph", ... }`),
 * or null (currently unused — paragraphs always resolve to at least
 * themselves). Tables / SDTs / text-boxes / drawing shapes / TOCs are
 * mutated in place and the same reference is returned.
 */
function patchBlock(
  block: BodyContent,
  patchMap: Map<string, PatchOperation>
): BodyContent | BodyContent[] | null {
  switch (block.type) {
    case "paragraph":
      return patchParagraph(block, patchMap);
    case "table":
      patchTable(block as Table, patchMap);
      return block;
    case "sdt":
      patchSdt(block as StructuredDocumentTag, patchMap);
      return block;
    case "textBox":
      patchTextBox(block as TextBox, patchMap);
      return block;
    case "drawingShape":
      patchDrawingShape(block as DrawingShape, patchMap);
      return block;
    case "tableOfContents":
      patchTableOfContents(block as TableOfContents, patchMap);
      return block;
    default:
      return block;
  }
}

/** Patch text and structural placeholders inside table cells recursively. */
function patchTable(table: Table, patchMap: Map<string, PatchOperation>): void {
  for (const row of table.rows) {
    for (const cell of row.cells) {
      // TableCell.content is `(Paragraph | Table)[]`. patchBlockList may
      // expand a single paragraph into multiple paragraphs (for paragraph
      // patches) but never into types outside the cell-content union, so
      // the result is structurally compatible.
      const newContent = patchBlockList(cell.content as readonly BodyContent[], patchMap);
      (cell as Mutable<TableCell>).content = newContent as (Paragraph | Table)[];
    }
  }
}

/**
 * Patch the children of a structured document tag.
 *
 * SDT.content is `(Paragraph | Run | Table)[]`. Two cases:
 *   - Block SDTs hold paragraphs/tables → patch each through patchBlock.
 *   - Inline SDTs hold raw runs (a content control wrapping one or more
 *     runs inside a paragraph). Wrap them in a synthetic paragraph so
 *     placeholder matching works the same as in any other paragraph,
 *     then write the (possibly mutated) runs back. We never structurally
 *     replace a synthetic paragraph — paragraph/table/image patches on
 *     inline SDTs would corrupt the surrounding paragraph, so they are
 *     ignored for the inline-run case (text patches still apply).
 */
function patchSdt(sdt: StructuredDocumentTag, patchMap: Map<string, PatchOperation>): void {
  type SdtChild = Paragraph | Run | Table;
  const newChildren: SdtChild[] = [];

  // Buffer of consecutive inline runs so a placeholder split across runs
  // inside an inline SDT still gets stitched correctly.
  let runBuffer: Run[] = [];
  const flushRunBuffer = (): void => {
    if (runBuffer.length === 0) {
      return;
    }
    const synthetic: Paragraph = {
      type: "paragraph",
      children: runBuffer as Run[]
    };
    // Apply text patches only — structural patches on inline runs would
    // require splitting the enclosing paragraph, which we can't do here.
    const text = paragraphText(synthetic);
    for (const [placeholder, patch] of patchMap) {
      if (patch.content.type !== "text") {
        continue;
      }
      if (!text.includes(placeholder)) {
        continue;
      }
      replaceInParagraph(synthetic, placeholder, patch.content.text);
    }
    // Push the (possibly-mutated) runs back into the SDT content stream.
    for (const r of runBuffer) {
      newChildren.push(r);
    }
    runBuffer = [];
  };

  for (const child of sdt.content) {
    if (child && typeof child === "object" && "type" in child) {
      const typed = child as { type: string };
      if (typed.type === "paragraph") {
        flushRunBuffer();
        const result = patchParagraph(child as Paragraph, patchMap);
        if (result === null) {
          continue;
        }
        if (Array.isArray(result)) {
          // paragraph[] → keep all paragraph items, drop any non-paragraph
          // (defensive: PatchContent.paragraph carries Paragraph[], so this
          // branch only ever yields paragraphs in practice).
          for (const item of result) {
            if (item.type === "paragraph") {
              newChildren.push(item);
            }
          }
        } else if (result.type === "paragraph") {
          newChildren.push(result);
        } else if (result.type === "table") {
          newChildren.push(result);
        }
        // Other replacement types (image-as-paragraph is already a
        // paragraph) cannot appear here.
        continue;
      }
      if (typed.type === "table") {
        flushRunBuffer();
        patchTable(child as Table, patchMap);
        newChildren.push(child as Table);
        continue;
      }
    }
    // Inline run (no `type` discriminator on Run — it's the default shape).
    if (
      child &&
      typeof child === "object" &&
      !("type" in child) &&
      "content" in child &&
      Array.isArray((child as { content?: unknown }).content)
    ) {
      runBuffer.push(child as Run);
      continue;
    }
    // Anything else: pass through unchanged.
    newChildren.push(child as SdtChild);
  }
  flushRunBuffer();

  (sdt as Mutable<StructuredDocumentTag>).content = newChildren as readonly (
    | Paragraph
    | Run
    | Table
  )[];
}

/** Patch text and structural placeholders inside a text box's paragraphs. */
function patchTextBox(textBox: TextBox, patchMap: Map<string, PatchOperation>): void {
  const newContent = patchBlockList(textBox.content as readonly BodyContent[], patchMap);
  // TextBox.content is `Paragraph[]`. Discard any non-paragraph items
  // produced by structural patches (defensive — paragraph patches yield
  // paragraphs, table/image patches don't make sense inside a textBox).
  const paragraphs: Paragraph[] = [];
  for (const block of newContent) {
    if (block.type === "paragraph") {
      paragraphs.push(block);
    }
  }
  (textBox as Mutable<TextBox>).content = paragraphs;
}

/** Patch text and structural placeholders inside a drawing shape's text body. */
function patchDrawingShape(shape: DrawingShape, patchMap: Map<string, PatchOperation>): void {
  if (!shape.textContent || shape.textContent.length === 0) {
    return;
  }
  const newContent = patchBlockList(shape.textContent as readonly BodyContent[], patchMap);
  const paragraphs: Paragraph[] = [];
  for (const block of newContent) {
    if (block.type === "paragraph") {
      paragraphs.push(block);
    }
  }
  (shape as Mutable<DrawingShape>).textContent = paragraphs;
}

/**
 * Patch placeholders inside a TOC's cached paragraphs. The TOC field
 * caches the rendered text Word displays before the field is updated;
 * users authoring TOCs as templates may legitimately put placeholders in
 * the cached entries (e.g. document title shown as the TOC heading).
 */
function patchTableOfContents(toc: TableOfContents, patchMap: Map<string, PatchOperation>): void {
  if (!toc.cachedParagraphs || toc.cachedParagraphs.length === 0) {
    return;
  }
  const newContent = patchBlockList(toc.cachedParagraphs as readonly BodyContent[], patchMap);
  const paragraphs: Paragraph[] = [];
  for (const block of newContent) {
    if (block.type === "paragraph") {
      paragraphs.push(block);
    }
  }
  (toc as Mutable<TableOfContents>).cachedParagraphs = paragraphs;
}

/** Patch content in header/footer — supports all patch types like body paragraphs. */
function patchHeaderFooterContent(
  content: HeaderFooterContent,
  patchMap: Map<string, PatchOperation>
): void {
  const newChildren = patchBlockList(content.children as readonly BodyContent[], patchMap);
  // HeaderFooterContent.children is `(Paragraph | Table)[]`.
  (content as Mutable<HeaderFooterContent>).children = newChildren as (Paragraph | Table)[];
}

/** Patch text-only placeholders inside a single paragraph (footnotes/endnotes). */
function patchTextInParagraph(para: Paragraph, patchMap: Map<string, PatchOperation>): void {
  const text = paragraphText(para);
  for (const [placeholder, patch] of patchMap) {
    if (!text.includes(placeholder)) {
      continue;
    }
    if (patch.content.type === "text") {
      replaceInParagraph(para, placeholder, patch.content.text);
    }
  }
}
