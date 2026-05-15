/**
 * Conversion Intermediate Representation (IR)
 *
 * Format-agnostic semantic document model used by converters.
 * Sits between DocxDocument (OOXML-specific) and output formats (HTML, Markdown, etc.).
 *
 * Purpose:
 * - Normalize style resolution into direct formatting
 * - Resolve numbering into concrete list markers
 * - Flatten SDTs and tracked changes into final visible content
 * - Provide a common "loss report" mechanism
 * - Enable consistent behavior across all converters
 */

/** Severity of a conversion warning. */
export type ConversionSeverity = "info" | "warning" | "error";

/** A warning generated during conversion (feature loss, unsupported element, etc.). */
export interface ConversionWarning {
  readonly severity: ConversionSeverity;
  readonly code: string;
  readonly message: string;
  /** Path in the source document (e.g. "body[3].table.row[1].cell[0]"). */
  readonly path?: string;
}

/** A registered media asset for the conversion. */
export interface ConversionAsset {
  readonly id: string;
  readonly mimeType: string;
  readonly data: Uint8Array;
  readonly altText?: string;
  readonly width?: number; // in points
  readonly height?: number; // in points
}

/** Resolved inline formatting (no style inheritance — everything is explicit). */
export interface ResolvedFormatting {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly strikethrough?: boolean;
  readonly superscript?: boolean;
  readonly subscript?: boolean;
  readonly fontFamily?: string;
  readonly fontSize?: number; // in points
  readonly color?: string; // hex RGB
  readonly backgroundColor?: string;
  readonly code?: boolean; // monospace/code style
}

/** A semantic inline element. */
export type SemanticInline =
  | { readonly type: "text"; readonly text: string; readonly format?: ResolvedFormatting }
  | { readonly type: "lineBreak" }
  | {
      readonly type: "image";
      readonly assetId: string;
      readonly alt?: string;
      readonly width?: number;
      readonly height?: number;
    }
  | { readonly type: "link"; readonly href: string; readonly children: readonly SemanticInline[] }
  | { readonly type: "footnoteRef"; readonly id: number }
  | { readonly type: "endnoteRef"; readonly id: number }
  | { readonly type: "math"; readonly latex?: string; readonly text: string }
  | { readonly type: "code"; readonly text: string };

/** A semantic block element. */
export type SemanticBlock =
  | {
      readonly type: "paragraph";
      readonly children: readonly SemanticInline[];
      readonly style?: SemanticParagraphStyle;
    }
  | {
      readonly type: "heading";
      readonly level: 1 | 2 | 3 | 4 | 5 | 6;
      readonly children: readonly SemanticInline[];
    }
  | {
      readonly type: "list";
      readonly ordered: boolean;
      readonly items: readonly SemanticListItem[];
    }
  | {
      readonly type: "table";
      readonly rows: readonly SemanticTableRow[];
      readonly caption?: string;
    }
  | { readonly type: "codeBlock"; readonly language?: string; readonly text: string }
  | { readonly type: "blockquote"; readonly children: readonly SemanticBlock[] }
  | { readonly type: "horizontalRule" }
  | {
      readonly type: "image";
      readonly assetId: string;
      readonly alt?: string;
      readonly width?: number;
      readonly height?: number;
      readonly caption?: string;
    }
  | {
      /**
       * Block-level mathematical expression. `text` is a plain-text
       * fallback (always present); `mathML` is the full MathML
       * representation when the converter could derive one.
       */
      readonly type: "math";
      readonly text: string;
      readonly mathML?: string;
    }
  | {
      /**
       * A chart reference. The chart binary is preserved in `assets` only
       * when the source provided a rendered SVG; otherwise consumers
       * should fall back to the `title` / `altText` for a textual
       * placeholder.
       */
      readonly type: "chart";
      readonly chartId: string;
      readonly title?: string;
      readonly altText?: string;
      /** Asset id of an SVG rendering, when available. */
      readonly svgAssetId?: string;
    }
  | {
      /** Inline check-box state (carries `<w:checkBox>` semantics). */
      readonly type: "checkBox";
      readonly checked: boolean;
      readonly label?: string;
    }
  | {
      /**
       * Embedded foreign content (`<w:altChunk>`): HTML, RTF, plain text,
       * etc. The renderer can either inline the data directly (when it
       * knows how to handle `contentType`) or fall back to a placeholder.
       */
      readonly type: "embed";
      readonly contentType: string;
      readonly data?: Uint8Array;
      readonly fileName?: string;
    }
  | {
      /**
       * Opaque OOXML drawing markup that cannot be safely flattened into
       * the semantic model. Carried verbatim so consumers that want full
       * fidelity (e.g. an OOXML→OOXML pipeline) can preserve the source
       * markup; markdown / html renderers may emit a placeholder instead.
       */
      readonly type: "raw";
      readonly format: "ooxml-drawing";
      readonly xml: string;
    };

/** Paragraph style properties (resolved, not inherited). */
export interface SemanticParagraphStyle {
  readonly alignment?: "left" | "center" | "right" | "justify";
  readonly indentLeft?: number; // in points
  readonly indentRight?: number;
  readonly spaceBefore?: number;
  readonly spaceAfter?: number;
}

/** A list item (possibly nested). */
export interface SemanticListItem {
  readonly children: readonly SemanticInline[];
  readonly subList?: SemanticBlock; // nested list
}

/** A table row. */
export interface SemanticTableRow {
  readonly cells: readonly SemanticTableCell[];
  readonly isHeader?: boolean;
}

/** A table cell. */
export interface SemanticTableCell {
  readonly children: readonly SemanticBlock[];
  readonly colSpan?: number;
  readonly rowSpan?: number;
  readonly alignment?: "left" | "center" | "right";
}

/** A footnote/endnote definition. */
export interface SemanticNote {
  readonly id: number;
  readonly children: readonly SemanticBlock[];
}

/** The complete semantic document. */
export interface SemanticDocument {
  readonly blocks: readonly SemanticBlock[];
  readonly assets: readonly ConversionAsset[];
  readonly footnotes: readonly SemanticNote[];
  readonly endnotes: readonly SemanticNote[];
  readonly metadata?: {
    readonly title?: string;
    readonly author?: string;
    readonly subject?: string;
    readonly language?: string;
  };
}

/** Conversion context passed through the conversion pipeline. */
export interface ConversionContext {
  readonly warnings: ConversionWarning[];
  readonly assets: ConversionAsset[];
  addWarning(severity: ConversionSeverity, code: string, message: string, path?: string): void;
  registerAsset(mimeType: string, data: Uint8Array, altText?: string): string;
}

/** Create a new conversion context. */
export function createConversionContext(): ConversionContext {
  let nextAssetId = 1;
  const warnings: ConversionWarning[] = [];
  const assets: ConversionAsset[] = [];

  return {
    warnings,
    assets,
    addWarning(severity: ConversionSeverity, code: string, message: string, path?: string): void {
      warnings.push({ severity, code, message, path });
    },
    registerAsset(mimeType: string, data: Uint8Array, altText?: string): string {
      const id = `asset-${nextAssetId++}`;
      assets.push({ id, mimeType, data, altText });
      return id;
    }
  };
}
