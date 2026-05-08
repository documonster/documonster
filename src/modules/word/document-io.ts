/**
 * DOCX Module - Document IO
 *
 * IO operations that depend on docx-packager and docx-reader.
 * Separated from document.ts so that builder helpers can be imported
 * without pulling in archive/xml/writer code.
 */

import { packageDocx } from "./docx-packager";
import { readDocx } from "./docx-reader";
import { bytesToBase64 } from "./internal-utils";
import type {
  DocxDocument,
  BodyContent,
  Paragraph,
  Run,
  Table,
  HeaderFooterContent,
  ImageDef,
  InlineImageContent
} from "./types";

// =============================================================================
// Document IO (toBuffer / toBase64)
// =============================================================================

/** Package a DocxDocument model to DOCX bytes. */
export async function toBuffer(doc: DocxDocument, compressionLevel?: number): Promise<Uint8Array> {
  return packageDocx(doc, compressionLevel);
}

/** Package a DocxDocument model to base64 string. */
export async function toBase64(doc: DocxDocument, compressionLevel?: number): Promise<string> {
  const bytes = await toBuffer(doc, compressionLevel);
  return bytesToBase64(bytes);
}

// =============================================================================
// Patcher / Template Fill API
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

/** Options for patchDocument. */
export interface PatchOptions {
  /** Compression level (0-9). Default: 6. */
  readonly compressionLevel?: number;
}

/**
 * Read an existing DOCX file, replace placeholders with content, and produce a new DOCX.
 *
 * Placeholders are strings like `{{name}}` embedded in the document text.
 * They may span across multiple runs — the patcher handles cross-run matching.
 *
 * Supported patch content types:
 * - `text` — simple text replacement (preserves formatting of the first run)
 * - `paragraph` — replaces the entire paragraph containing the placeholder
 * - `table` — replaces the entire paragraph with a table
 * - `image` — replaces the placeholder with an inline image
 *
 * @param buffer - The source DOCX file as a Uint8Array.
 * @param patches - Array of patch operations to apply.
 * @param options - Optional compression settings.
 * @returns New DOCX file as a Uint8Array.
 */
export async function patchDocument(
  buffer: Uint8Array,
  patches: readonly PatchOperation[],
  options?: PatchOptions
): Promise<Uint8Array> {
  const doc = await readDocx(buffer);

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

  const patched: DocxDocument = {
    ...doc,
    body: newBody,
    images: images.length > 0 ? images : undefined
  };

  return packageDocx(patched, options?.compressionLevel);
}

// =============================================================================
// Internal helpers
// =============================================================================

/** Extract concatenated plain text from a paragraph's runs. */
function paragraphText(para: Paragraph): string {
  let t = "";
  for (const child of para.children) {
    if ("content" in child && Array.isArray(child.content)) {
      for (const c of child.content) {
        if ("type" in c && c.type === "text" && "text" in c) {
          t += (c as { text: string }).text;
        }
      }
    }
  }
  return t;
}

/** Replace text within a single paragraph. */
function replaceInParagraph(para: Paragraph, search: string, replacement: string): void {
  for (const child of para.children) {
    if (!("content" in child) || !Array.isArray(child.content)) {
      continue;
    }
    for (const c of child.content) {
      if (!("type" in c) || c.type !== "text" || !("text" in c)) {
        continue;
      }
      const before = (c as { text: string }).text;
      if (before.includes(search)) {
        (c as { text: string }).text = before.replaceAll(search, replacement);
      }
    }
  }

  // Cross-run replacement fallback
  const fullText = paragraphText(para);
  if (fullText.includes(search)) {
    const newText = fullText.replaceAll(search, replacement);
    let placed = false;
    for (const child of para.children) {
      if (!("content" in child) || !Array.isArray(child.content)) {
        continue;
      }
      for (const c of child.content) {
        if (!("type" in c) || c.type !== "text" || !("text" in c)) {
          continue;
        }
        if (!placed) {
          (c as { text: string }).text = newText;
          placed = true;
        } else {
          (c as { text: string }).text = "";
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
  const text = paragraphText(para);

  for (const [placeholder, patch] of patchMap) {
    if (!text.includes(placeholder)) {
      continue;
    }

    switch (patch.content.type) {
      case "text": {
        replaceInParagraph(para, placeholder, patch.content.text);
        return para;
      }
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
        } else {
          newContent.push(block);
        }
      }
      (cell as any).content = newContent;
    }
  }
}

/** Patch text in header/footer content. */
function patchHeaderFooterContent(
  content: HeaderFooterContent,
  patchMap: Map<string, PatchOperation>
): void {
  for (const child of content.children) {
    if (child.type === "paragraph") {
      for (const [placeholder, patch] of patchMap) {
        if (patch.content.type === "text") {
          const text = paragraphText(child);
          if (text.includes(placeholder)) {
            replaceInParagraph(child, placeholder, patch.content.text);
          }
        }
      }
    }
  }
}
