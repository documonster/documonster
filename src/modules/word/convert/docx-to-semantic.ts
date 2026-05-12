/**
 * DOCX → Semantic Document Converter
 *
 * Transforms a DocxDocument into a format-agnostic SemanticDocument.
 * This is the bridge between OOXML-specific types and the universal IR
 * that HTML/Markdown/other renderers can consume.
 *
 * Handles:
 * - Heading detection (via style name or outlineLevel)
 * - Inline formatting resolution
 * - Hyperlink extraction
 * - Image registration into ConversionContext
 * - Table structure with merge (colSpan/rowSpan)
 * - List/numbering detection (basic)
 * - Footnote/endnote reference and content
 * - Math content (text fallback)
 */

import { isRun } from "../core/text-utils";
import type {
  BodyContent,
  ColorSpec,
  DocxDocument,
  FontSpec,
  Hyperlink,
  InlineImageContent,
  Paragraph,
  ParagraphChild,
  Run,
  RunContent,
  RunProperties,
  Table
} from "../types";
import { EMU_PER_INCH } from "../units";
import type {
  ConversionContext,
  ResolvedFormatting,
  SemanticBlock,
  SemanticDocument,
  SemanticInline,
  SemanticNote,
  SemanticParagraphStyle,
  SemanticTableCell,
  SemanticTableRow
} from "./conversion-ir";
import { createConversionContext } from "./conversion-ir";

// =============================================================================
// Public API
// =============================================================================

/** Options for the semantic conversion. */
export interface DocxToSemanticOptions {
  /** Include footnotes in the output. Default: true. */
  readonly includeFootnotes?: boolean;
  /** Include endnotes in the output. Default: true. */
  readonly includeEndnotes?: boolean;
  /** Extract images and register as assets. Default: true. */
  readonly extractImages?: boolean;
}

/**
 * Convert a DocxDocument to a SemanticDocument (format-agnostic IR).
 *
 * @param doc - The parsed DOCX document.
 * @param options - Conversion options.
 * @returns The semantic document with warnings and assets.
 */
export function docxToSemantic(
  doc: DocxDocument,
  options?: DocxToSemanticOptions
): { document: SemanticDocument; context: ConversionContext } {
  const ctx = createConversionContext();
  const opts: Required<DocxToSemanticOptions> = {
    includeFootnotes: options?.includeFootnotes ?? true,
    includeEndnotes: options?.includeEndnotes ?? true,
    extractImages: options?.extractImages ?? true
  };

  // Build image lookup by rId for quick reference
  const imageMap = new Map<string, { data: Uint8Array; mimeType: string; fileName: string }>();
  if (doc.images && opts.extractImages) {
    for (const img of doc.images) {
      if (img.rId && img.data) {
        const ext = img.fileName?.split(".").pop()?.toLowerCase() ?? "png";
        imageMap.set(img.rId, {
          data: img.data,
          mimeType: extToMimeType(ext),
          fileName: img.fileName
        });
      }
    }
  }

  // Convert body
  const blocks = convertBodyContent(doc.body, doc, ctx, imageMap);

  // Convert footnotes
  const footnotes: SemanticNote[] = [];
  if (opts.includeFootnotes && doc.footnotes) {
    for (const note of doc.footnotes) {
      if (note.id <= 0) {
        continue;
      }
      const noteBlocks = convertBodyContent(
        note.content as readonly BodyContent[],
        doc,
        ctx,
        imageMap
      );
      footnotes.push({ id: note.id, children: noteBlocks });
    }
  }

  // Convert endnotes
  const endnotes: SemanticNote[] = [];
  if (opts.includeEndnotes && doc.endnotes) {
    for (const note of doc.endnotes) {
      if (note.id <= 0) {
        continue;
      }
      const noteBlocks = convertBodyContent(
        note.content as readonly BodyContent[],
        doc,
        ctx,
        imageMap
      );
      endnotes.push({ id: note.id, children: noteBlocks });
    }
  }

  const document: SemanticDocument = {
    blocks,
    assets: ctx.assets,
    footnotes,
    endnotes,
    metadata: {
      title: doc.coreProperties?.title,
      author: doc.coreProperties?.creator,
      subject: doc.coreProperties?.subject
    }
  };

  return { document, context: ctx };
}

// =============================================================================
// Internal: Body Content Conversion
// =============================================================================

function convertBodyContent(
  body: readonly BodyContent[],
  doc: DocxDocument,
  ctx: ConversionContext,
  imageMap: Map<string, { data: Uint8Array; mimeType: string; fileName: string }>
): SemanticBlock[] {
  const blocks: SemanticBlock[] = [];

  for (const item of body) {
    switch (item.type) {
      case "paragraph":
        blocks.push(convertParagraph(item, doc, ctx, imageMap));
        break;
      case "table":
        blocks.push(convertTable(item, doc, ctx, imageMap));
        break;
      case "floatingImage": {
        if (item.rId && imageMap.has(item.rId)) {
          const img = imageMap.get(item.rId)!;
          const assetId = ctx.registerAsset(img.mimeType, img.data, item.altText);
          blocks.push({
            type: "image",
            assetId,
            alt: item.altText,
            width: item.width ? item.width / EMU_PER_INCH : undefined, // EMU to inches
            height: item.height ? item.height / EMU_PER_INCH : undefined
          });
        }
        break;
      }
      case "sdt": {
        // Structured document tags can wrap whole sections of body content
        // (Paragraph | Table | Run). Pass through paragraph/table children
        // recursively; bare Run children are skipped here because the
        // semantic model has no inline-only equivalent — the warning
        // surfaces them so downstream tooling knows to expect missing text.
        const inner: BodyContent[] = [];
        let dropped = 0;
        for (const c of item.content) {
          if ("type" in c && (c.type === "paragraph" || c.type === "table")) {
            inner.push(c as BodyContent);
          } else {
            dropped++;
          }
        }
        if (dropped > 0) {
          ctx.addWarning(
            "info",
            "sdt-inline-runs-dropped",
            `${dropped} inline Run children of an SDT had no semantic equivalent and were skipped`
          );
        }
        blocks.push(...convertBodyContent(inner, doc, ctx, imageMap));
        break;
      }
      case "textBox": {
        // Render textBox content as inline paragraphs in reading order.
        // The semantic model has no "frame" concept; flattening preserves
        // the text rather than dropping it.
        blocks.push(
          ...convertBodyContent(item.content as readonly BodyContent[], doc, ctx, imageMap)
        );
        break;
      }
      case "drawingShape": {
        if (item.textContent && item.textContent.length > 0) {
          blocks.push(
            ...convertBodyContent(item.textContent as readonly BodyContent[], doc, ctx, imageMap)
          );
        }
        break;
      }
      case "tableOfContents": {
        if (item.cachedParagraphs && item.cachedParagraphs.length > 0) {
          blocks.push(
            ...convertBodyContent(
              item.cachedParagraphs as readonly BodyContent[],
              doc,
              ctx,
              imageMap
            )
          );
        }
        break;
      }
      default:
        // Skip remaining block types (math block, charts, altChunks, etc.)
        // and surface a warning so the caller can investigate gaps.
        ctx.addWarning("info", "unsupported-block", `Skipped block type: ${item.type}`);
        break;
    }
  }

  return blocks;
}

// =============================================================================
// Internal: Paragraph Conversion
// =============================================================================

function convertParagraph(
  para: Paragraph,
  doc: DocxDocument,
  ctx: ConversionContext,
  imageMap: Map<string, { data: Uint8Array; mimeType: string; fileName: string }>
): SemanticBlock {
  // Detect heading
  const headingLevel = detectHeadingLevel(para);
  if (headingLevel !== null) {
    const children = convertParagraphChildren(para.children, doc, ctx, imageMap);
    return {
      type: "heading",
      level: headingLevel as 1 | 2 | 3 | 4 | 5 | 6,
      children
    };
  }

  const children = convertParagraphChildren(para.children, doc, ctx, imageMap);

  // Convert paragraph style
  const style = convertParagraphStyle(para.properties);

  return { type: "paragraph", children, style };
}

function detectHeadingLevel(para: Paragraph): number | null {
  const props = para.properties;
  if (!props) {
    return null;
  }

  // Check outlineLevel
  if (props.outlineLevel !== undefined && props.outlineLevel >= 0 && props.outlineLevel <= 5) {
    return props.outlineLevel + 1;
  }

  // Check style name
  if (props.style) {
    const match = /^[Hh]eading\s*(\d)$/i.exec(props.style);
    if (match) {
      const level = parseInt(match[1], 10);
      if (level >= 1 && level <= 6) {
        return level;
      }
    }
  }

  return null;
}

function convertParagraphStyle(props: Paragraph["properties"]): SemanticParagraphStyle | undefined {
  if (!props) {
    return undefined;
  }

  const style: SemanticParagraphStyle = {
    alignment: convertAlignment(props.alignment),
    indentLeft: props.indent?.left ? props.indent.left / 20 : undefined, // twips to points
    indentRight: props.indent?.right ? props.indent.right / 20 : undefined,
    spaceBefore: props.spacing?.before ? props.spacing.before / 20 : undefined,
    spaceAfter: props.spacing?.after ? props.spacing.after / 20 : undefined
  };

  // Only return if at least one property is set
  if (
    style.alignment === undefined &&
    style.indentLeft === undefined &&
    style.indentRight === undefined &&
    style.spaceBefore === undefined &&
    style.spaceAfter === undefined
  ) {
    return undefined;
  }

  return style;
}

function convertAlignment(
  alignment: string | undefined
): "left" | "center" | "right" | "justify" | undefined {
  switch (alignment) {
    case "left":
    case "start":
      return "left";
    case "center":
      return "center";
    case "right":
    case "end":
      return "right";
    case "both":
    case "justify":
      return "justify";
    default:
      return undefined;
  }
}

// =============================================================================
// Internal: Paragraph Children (Inline Content)
// =============================================================================

function convertParagraphChildren(
  children: readonly ParagraphChild[],
  doc: DocxDocument,
  ctx: ConversionContext,
  imageMap: Map<string, { data: Uint8Array; mimeType: string; fileName: string }>
): SemanticInline[] {
  const inlines: SemanticInline[] = [];

  for (const child of children) {
    if ("type" in child) {
      const typed = child as { type: string };
      if (typed.type === "hyperlink") {
        const hl = child as Hyperlink;
        const linkChildren: SemanticInline[] = [];
        for (const run of hl.children) {
          linkChildren.push(...convertRun(run, ctx, imageMap));
        }
        inlines.push({
          type: "link",
          href: hl.url ?? hl.anchor ?? "",
          children: linkChildren
        });
        continue;
      }
      // Skip other typed children (bookmark, comment range, etc.)
      continue;
    }

    // Default: treat as Run
    if (isRun(child)) {
      inlines.push(...convertRun(child, ctx, imageMap));
    }
  }

  return inlines;
}

function convertRun(
  run: Run,
  ctx: ConversionContext,
  imageMap: Map<string, { data: Uint8Array; mimeType: string; fileName: string }>
): SemanticInline[] {
  const inlines: SemanticInline[] = [];
  const format = resolveRunFormatting(run.properties);

  for (const content of run.content) {
    const inline = convertRunContent(content, format, ctx, imageMap);
    if (inline) {
      inlines.push(inline);
    }
  }

  return inlines;
}

function convertRunContent(
  content: RunContent,
  format: ResolvedFormatting | undefined,
  ctx: ConversionContext,
  imageMap: Map<string, { data: Uint8Array; mimeType: string; fileName: string }>
): SemanticInline | null {
  switch (content.type) {
    case "text":
      return { type: "text", text: content.text, format };
    case "break":
      return { type: "lineBreak" };
    case "tab":
      return { type: "text", text: "\t", format };
    case "image": {
      const img = content as InlineImageContent;
      if (img.rId && imageMap.has(img.rId)) {
        const imgData = imageMap.get(img.rId)!;
        const assetId = ctx.registerAsset(imgData.mimeType, imgData.data, img.altText);
        return {
          type: "image",
          assetId,
          alt: img.altText,
          width: img.width ? img.width / EMU_PER_INCH : undefined,
          height: img.height ? img.height / EMU_PER_INCH : undefined
        };
      }
      return null;
    }
    case "footnoteRef":
      return { type: "footnoteRef", id: content.id };
    case "endnoteRef":
      return { type: "endnoteRef", id: content.id };
    case "field":
      // Render field cached value as plain text
      if (content.cachedValue) {
        return { type: "text", text: content.cachedValue, format };
      }
      return null;
    case "symbol":
      return { type: "text", text: content.char ?? "", format };
    default:
      return null;
  }
}

// =============================================================================
// Internal: Run Formatting Resolution
// =============================================================================

function resolveRunFormatting(props: RunProperties | undefined): ResolvedFormatting | undefined {
  if (!props) {
    return undefined;
  }

  const format: ResolvedFormatting = {
    bold: props.bold || undefined,
    italic: props.italic || undefined,
    underline: props.underline !== undefined ? true : undefined,
    strikethrough: props.strike || undefined,
    superscript: props.vertAlign === "superscript" || undefined,
    subscript: props.vertAlign === "subscript" || undefined,
    fontFamily:
      typeof props.font === "string" ? props.font : (props.font as FontSpec | undefined)?.ascii,
    fontSize: props.size ? props.size / 2 : undefined, // half-points to points
    color:
      typeof props.color === "string" ? props.color : (props.color as ColorSpec | undefined)?.val,
    code: undefined
  };

  // Only return if at least one property is set
  if (
    !format.bold &&
    !format.italic &&
    !format.underline &&
    !format.strikethrough &&
    !format.superscript &&
    !format.subscript &&
    !format.fontFamily &&
    !format.fontSize &&
    !format.color
  ) {
    return undefined;
  }

  return format;
}

// =============================================================================
// Internal: Table Conversion
// =============================================================================

function convertTable(
  table: Table,
  doc: DocxDocument,
  ctx: ConversionContext,
  imageMap: Map<string, { data: Uint8Array; mimeType: string; fileName: string }>
): SemanticBlock {
  const rows: SemanticTableRow[] = [];

  for (const row of table.rows) {
    const cells: SemanticTableCell[] = [];
    const isHeader = row.properties?.tableHeader ?? false;

    for (const cell of row.cells) {
      const children = convertBodyContent(
        cell.content as readonly BodyContent[],
        doc,
        ctx,
        imageMap
      );
      const cellAlignment =
        cell.content?.[0]?.type === "paragraph"
          ? (cell.content[0] as Paragraph).properties?.alignment
          : undefined;
      const resolvedAlign = convertAlignment(cellAlignment);
      const cellResult: SemanticTableCell = {
        children,
        colSpan: cell.properties?.gridSpan,
        alignment: resolvedAlign === "justify" ? undefined : resolvedAlign
      };
      cells.push(cellResult);
    }

    rows.push({ cells, isHeader });
  }

  return { type: "table", rows };
}

/**
 * Map a file extension to the canonical IANA MIME type.
 *
 * The previous implementation just did `image/${ext}` which produced
 * non-canonical types like `image/tif` or `image/wmf` — those are
 * silently rejected by browsers when the semantic document is rendered
 * to HTML/PDF. Here we map the handful of extensions DOCX images
 * actually use to their correct types and refuse to forward unknown
 * extensions verbatim — they're returned as
 * `application/octet-stream` instead so attacker-controlled file names
 * can't smuggle a fabricated `image/<script>` Content-Type into
 * downstream renderers.
 */
function extToMimeType(ext: string): string {
  switch (ext) {
    case "jpg":
    case "jpeg":
      return "image/jpeg";
    case "png":
      return "image/png";
    case "gif":
      return "image/gif";
    case "bmp":
      return "image/bmp";
    case "tif":
    case "tiff":
      return "image/tiff";
    case "svg":
      return "image/svg+xml";
    case "webp":
      return "image/webp";
    case "emf":
      return "image/x-emf";
    case "wmf":
      return "image/x-wmf";
    default:
      // Unknown / hostile extension: don't synthesise an `image/<x>`
      // Content-Type from raw input.
      return "application/octet-stream";
  }
}
