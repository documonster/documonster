/**
 * DOCX Module - Patcher
 *
 * Internal logic for replacing placeholders (e.g. `{{name}}`) in a DocxDocument
 * with text, paragraphs, tables, or images. Used by `patchDocument` and
 * `patchTemplate` in document-io.ts.
 *
 * Public API is exposed through document-io.ts; this file contains the
 * pure patch logic (no IO).
 */

import { type Mutable } from "./core/internal-utils";
import { isRun } from "./core/text-utils";
import type {
  BodyContent,
  DocxDocument,
  HeaderFooterContent,
  ImageDef,
  InlineImageContent,
  Paragraph,
  Run,
  StructuredDocumentTag,
  Table,
  TableCell,
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
 * Apply patches to a parsed document model in place.
 *
 * Mutates the document's body, headers, footers, footnotes, endnotes, and
 * appends new images for image-typed patches.
 *
 * @param doc - The document model (mutated in place).
 * @param patches - Patches to apply.
 * @returns The same document with patches applied (and updated images list).
 */
export function applyPatchesToDocument(
  doc: DocxDocument,
  patches: readonly PatchOperation[]
): DocxDocument {
  // Build lookup map for quick placeholder matching
  const patchMap = new Map<string, PatchOperation>();
  for (const patch of patches) {
    patchMap.set(patch.placeholder, patch);
  }

  // Process body content
  const newBody: BodyContent[] = [];
  for (const block of doc.body) {
    if (block.type === "paragraph") {
      const result = patchParagraph(block, patchMap);
      if (result) {
        if (Array.isArray(result)) {
          newBody.push(...result);
        } else {
          newBody.push(result);
        }
      }
    } else if (block.type === "table") {
      patchTable(block as Table, patchMap);
      newBody.push(block);
    } else {
      newBody.push(block);
    }
  }

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

  // Patch footnotes
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

  // Add any new images from patches
  const images = doc.images ? [...doc.images] : [];
  for (const patch of patches) {
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

/** Extract concatenated plain text from a paragraph's runs. */
function paragraphText(para: Paragraph): string {
  let t = "";
  for (const child of para.children) {
    if (isRun(child)) {
      for (const c of child.content) {
        if (c.type === "text") {
          t += (c as TextContent).text;
        }
      }
    }
  }
  return t;
}

/** Replace text within a single paragraph. */
function replaceInParagraph(para: Paragraph, search: string, replacement: string): void {
  for (const child of para.children) {
    if (!isRun(child)) {
      continue;
    }
    for (const c of child.content) {
      if (c.type !== "text") {
        continue;
      }
      const tc = c as Mutable<TextContent>;
      if (tc.text.includes(search)) {
        tc.text = tc.text.replaceAll(search, replacement);
      }
    }
  }

  // Cross-run replacement fallback
  const fullText = paragraphText(para);
  if (fullText.includes(search)) {
    const newText = fullText.replaceAll(search, replacement);
    let placed = false;
    for (const child of para.children) {
      if (!isRun(child)) {
        continue;
      }
      for (const c of child.content) {
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
    }
  }
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
        const img = patch.content.image;
        const rId = img.rId ?? `rId_img_${img.fileName}`;
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

/** Patch text inside table cells recursively. */
function patchTable(table: Table, patchMap: Map<string, PatchOperation>): void {
  for (const row of table.rows) {
    for (const cell of row.cells) {
      const newContent: BodyContent[] = [];
      for (const block of cell.content) {
        if (block.type === "paragraph") {
          const result = patchParagraph(block, patchMap);
          if (result) {
            if (Array.isArray(result)) {
              newContent.push(...result);
            } else {
              newContent.push(result);
            }
          }
        } else if (block.type === "table") {
          patchTable(block as Table, patchMap);
          newContent.push(block);
        } else if ((block as BodyContent & { type: string }).type === "sdt") {
          // Patch inside SDT content
          const sdt = block as unknown as StructuredDocumentTag;
          for (const sdtChild of sdt.content) {
            const child = sdtChild as { type?: string };
            if (child.type === "paragraph") {
              const result = patchParagraph(sdtChild as Paragraph, patchMap);
              if (result && !Array.isArray(result)) {
                Object.assign(sdtChild, result);
              }
            } else if (child.type === "table") {
              patchTable(sdtChild as Table, patchMap);
            }
          }
          newContent.push(block);
        } else {
          newContent.push(block);
        }
      }
      (cell as Mutable<TableCell>).content = newContent as (Paragraph | Table)[];
    }
  }
}

/** Patch content in header/footer — supports all patch types like body paragraphs. */
function patchHeaderFooterContent(
  content: HeaderFooterContent,
  patchMap: Map<string, PatchOperation>
): void {
  const children = content.children as BodyContent[];
  const newChildren: BodyContent[] = [];
  for (const child of children) {
    if (child.type === "paragraph") {
      const result = patchParagraph(child, patchMap);
      if (result) {
        if (Array.isArray(result)) {
          newChildren.push(...result);
        } else {
          newChildren.push(result);
        }
      }
    } else if (child.type === "table") {
      patchTable(child as Table, patchMap);
      newChildren.push(child);
    } else {
      newChildren.push(child);
    }
  }
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
