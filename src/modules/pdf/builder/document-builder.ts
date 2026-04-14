/**
 * PDF document builder — high-level API for creating PDFs with free-form content.
 *
 * Unlike the table-oriented `pdf()` function, this builder gives direct control
 * over text positioning, vector drawing, images, and page management.
 *
 * @example Basic usage:
 * ```typescript
 * import { PdfDocumentBuilder } from "@cj-tech-master/excelts/pdf";
 *
 * const doc = new PdfDocumentBuilder();
 * const page = doc.addPage({ width: 595, height: 842 }); // A4
 *
 * page.drawText("Hello, World!", { x: 72, y: 750, fontSize: 24 });
 * page.drawRect({ x: 72, y: 700, width: 200, height: 30 });
 * page.drawCircle({ cx: 300, cy: 400, r: 50, fill: { r: 1, g: 0, b: 0 } });
 *
 * const bytes = await doc.build();
 * ```
 */

import { PdfContentStream } from "../core/pdf-stream";
import { PdfWriter } from "../core/pdf-writer";
import { PdfDict, pdfRef, pdfString, pdfNumber } from "../core/pdf-object";
import { writePdfAMetadata, writePdfAOutputIntent } from "../core/pdfa";
import { FontManager } from "../font/font-manager";
import { parseTtf } from "../font/ttf-parser";
import { wrapTextLines } from "../render/page-renderer";
import type { PdfColor, PdfExportOptions } from "../types";
import { initEncryption } from "../core/encryption";
import { writeImageXObject } from "./image-utils";

// =============================================================================
// Types
// =============================================================================

/** Page size configuration. */
export interface PageOptions {
  /** Page width in points (72pt = 1 inch). Default: 595.28 (A4). */
  width?: number;
  /** Page height in points (72pt = 1 inch). Default: 841.89 (A4). */
  height?: number;
}

/** Text drawing options. */
export interface DrawTextOptions {
  /** X position in points (from left edge). */
  x: number;
  /** Y position in points (from bottom edge — PDF coordinate system). */
  y: number;
  /** Font size in points. Default: 12. */
  fontSize?: number;
  /** Font family name. Default: "Helvetica". */
  fontFamily?: string;
  /** Bold. Default: false. */
  bold?: boolean;
  /** Italic. Default: false. */
  italic?: boolean;
  /** Text color. Default: black. */
  color?: PdfColor;
  /** Maximum width before word-wrap. Omit for no wrap. */
  maxWidth?: number;
  /** Line height multiplier. Default: 1.2. */
  lineHeight?: number;
}

/** Rectangle drawing options. */
export interface DrawRectOptions {
  x: number;
  y: number;
  width: number;
  height: number;
  /** Fill color. Omit for no fill. */
  fill?: PdfColor;
  /** Stroke color. Omit for no stroke. */
  stroke?: PdfColor;
  /** Line width for stroke. Default: 1. */
  lineWidth?: number;
  /** Corner radius for rounded rectangles. Default: 0. */
  borderRadius?: number;
}

/** Circle drawing options. */
export interface DrawCircleOptions {
  /** Center X. */
  cx: number;
  /** Center Y. */
  cy: number;
  /** Radius. */
  r: number;
  /** Fill color. Omit for no fill. */
  fill?: PdfColor;
  /** Stroke color. Omit for no stroke. */
  stroke?: PdfColor;
  /** Line width for stroke. Default: 1. */
  lineWidth?: number;
}

/** Ellipse drawing options. */
export interface DrawEllipseOptions {
  /** Center X. */
  cx: number;
  /** Center Y. */
  cy: number;
  /** Horizontal radius. */
  rx: number;
  /** Vertical radius. */
  ry: number;
  /** Fill color. Omit for no fill. */
  fill?: PdfColor;
  /** Stroke color. Omit for no stroke. */
  stroke?: PdfColor;
  /** Line width for stroke. Default: 1. */
  lineWidth?: number;
}

/** Line drawing options. */
export interface DrawLineOptions {
  /** Start X. */
  x1: number;
  /** Start Y. */
  y1: number;
  /** End X. */
  x2: number;
  /** End Y. */
  y2: number;
  /** Stroke color. Default: black. */
  color?: PdfColor;
  /** Line width. Default: 1. */
  lineWidth?: number;
  /** Dash pattern. Default: solid. */
  dashPattern?: number[];
}

/** Path drawing options. */
export interface DrawPathOptions {
  /** Fill color. Omit for no fill. */
  fill?: PdfColor;
  /** Stroke color. Omit for no stroke. */
  stroke?: PdfColor;
  /** Line width. Default: 1. */
  lineWidth?: number;
  /** Close the path before painting. Default: false. */
  closePath?: boolean;
}

/** A point in a path. */
export type PathOp =
  | { op: "move"; x: number; y: number }
  | { op: "line"; x: number; y: number }
  | { op: "curve"; x1: number; y1: number; x2: number; y2: number; x3: number; y3: number }
  | { op: "close" };

/** Image drawing options. */
export interface DrawImageOptions {
  /** Raw image bytes. */
  data: Uint8Array;
  /** Image format. */
  format: "jpeg" | "png";
  /** X position. */
  x: number;
  /** Y position (bottom edge of image in PDF coordinates). */
  y: number;
  /** Display width in points. */
  width: number;
  /** Display height in points. */
  height: number;
}

/** Document metadata. */
export interface DocumentMetadata {
  title?: string;
  author?: string;
  subject?: string;
  creator?: string;
}

/** Options for table of contents generation. */
export interface TocOptions {
  /** Title displayed at the top of the TOC page. Default: "Table of Contents". */
  title?: string;
  /** Font size for TOC entries in points. Default: 12. */
  fontSize?: number;
  /** Indentation in points per nesting level. Default: 20. */
  indent?: number;
}

/** @internal Bookmark node stored during document construction. */
interface BookmarkNode {
  /** Bookmark display title. */
  title: string;
  /** Zero-based page index this bookmark points to. */
  pageIndex: number;
  /** Child bookmarks. */
  children: BookmarkNode[];
}

/** @internal Link annotation recorded on a page for build-time serialization. */
interface PageAnnotation {
  /** Rectangle [x1, y1, x2, y2] in PDF coordinates. */
  rect: [number, number, number, number];
  /** Zero-based page index to link to. */
  destPageIndex: number;
}

// =============================================================================
// Annotation Types (for builder-created annotations)
// =============================================================================

/** Annotation types that can be created via the builder API. */
export type AnnotationType =
  | "Highlight"
  | "Underline"
  | "StrikeOut"
  | "Squiggly"
  | "Text"
  | "FreeText"
  | "Stamp";

/** Options for text markup annotations (Highlight, Underline, StrikeOut, Squiggly). */
export interface TextMarkupAnnotationOptions {
  /** Annotation subtype. */
  type: "Highlight" | "Underline" | "StrikeOut" | "Squiggly";
  /** Bounding rectangle [x1, y1, x2, y2]. */
  rect: [number, number, number, number];
  /**
   * QuadPoints — four pairs of (x,y) defining the marked text region.
   * Must be groups of 8 numbers (4 corners per quad). Order per PDF spec:
   * bottom-left, bottom-right, top-left, top-right (some viewers use
   * top-left, top-right, bottom-left, bottom-right — the spec is ambiguous).
   * If omitted, defaults to the corners of `rect`.
   */
  quadPoints?: number[];
  /** Annotation color (RGB, 0–1). Default: yellow for highlight, red for others. */
  color?: PdfColor;
  /** Text contents (e.g., comment text). */
  contents?: string;
  /** Author / title. */
  author?: string;
}

/** Options for a sticky note (Text) annotation. */
export interface TextAnnotationOptions {
  type: "Text";
  /** Position — the icon appears at this point. */
  rect: [number, number, number, number];
  /** Comment text. */
  contents?: string;
  /** Author. */
  author?: string;
  /** Icon name. Default: "Note". */
  iconName?: "Comment" | "Key" | "Note" | "Help" | "NewParagraph" | "Paragraph" | "Insert";
  /** Annotation color. Default: yellow. */
  color?: PdfColor;
  /** Whether the popup is initially open. Default: false. */
  open?: boolean;
}

/** Options for a free-text annotation (in-line text). */
export interface FreeTextAnnotationOptions {
  type: "FreeText";
  /** Bounding rectangle [x1, y1, x2, y2]. */
  rect: [number, number, number, number];
  /** The displayed text. */
  contents: string;
  /** Font size. Default: 12. */
  fontSize?: number;
  /** Text color. Default: black. */
  color?: PdfColor;
  /** Border color. Omit for no border. */
  borderColor?: PdfColor;
  /** Author. */
  author?: string;
}

/** Options for a rubber stamp annotation. */
export interface StampAnnotationOptions {
  type: "Stamp";
  /** Bounding rectangle [x1, y1, x2, y2]. */
  rect: [number, number, number, number];
  /** Standard stamp name. */
  stampName?:
    | "Approved"
    | "Experimental"
    | "NotApproved"
    | "AsIs"
    | "Expired"
    | "NotForPublicRelease"
    | "Confidential"
    | "Final"
    | "Sold"
    | "Departmental"
    | "ForComment"
    | "TopSecret"
    | "Draft"
    | "ForPublicRelease";
  /** Annotation color. */
  color?: PdfColor;
  /** Comment text. */
  contents?: string;
  /** Author. */
  author?: string;
}

/** Union of all annotation option types. */
export type AnnotationOptions =
  | TextMarkupAnnotationOptions
  | TextAnnotationOptions
  | FreeTextAnnotationOptions
  | StampAnnotationOptions;

/** @internal Stored annotation for build-time serialization. */
interface BuilderAnnotation {
  subtype: string;
  rect: [number, number, number, number];
  entries: Array<[string, string]>;
}

// =============================================================================
// Form Field Types (for builder-created forms)
// =============================================================================

/** Common options shared by all form field types. */
interface FormFieldBaseOptions {
  /** Fully qualified field name (e.g., "form.name"). */
  name: string;
  /** Bounding rectangle [x1, y1, x2, y2]. */
  rect: [number, number, number, number];
  /** Default value. */
  value?: string;
  /** Read-only. Default: false. */
  readOnly?: boolean;
  /** Required. Default: false. */
  required?: boolean;
}

/** Options for creating a text input field. */
export interface TextFieldOptions extends FormFieldBaseOptions {
  type: "text";
  /** Maximum character count. Omit for unlimited. */
  maxLength?: number;
  /** Multiline. Default: false. */
  multiline?: boolean;
  /** Password field (masked input). Default: false. */
  password?: boolean;
}

/** Options for creating a checkbox. */
export interface CheckboxOptions extends FormFieldBaseOptions {
  type: "checkbox";
  /** Whether initially checked. Default: false. */
  checked?: boolean;
}

/** Options for creating a dropdown (combo box). */
export interface DropdownOptions extends FormFieldBaseOptions {
  type: "dropdown";
  /** Available options. */
  options: string[];
  /** Allow typing a custom value. Default: false. */
  editable?: boolean;
}

/** Options for creating a radio button group. */
export interface RadioGroupOptions {
  type: "radio";
  /** Fully qualified field name for the group. */
  name: string;
  /** Individual radio buttons. */
  buttons: Array<{
    /** Bounding rectangle. */
    rect: [number, number, number, number];
    /** Export value for this button. */
    value: string;
  }>;
  /** Initially selected value. */
  selected?: string;
  /** Read-only. Default: false. */
  readOnly?: boolean;
  /** Required. Default: false. */
  required?: boolean;
}

/** Union of all form field creation options. */
export type FormFieldOptions =
  | TextFieldOptions
  | CheckboxOptions
  | DropdownOptions
  | RadioGroupOptions;

/** @internal Stored form field for build-time serialization. */
interface BuilderFormField {
  options: FormFieldOptions;
}

// =============================================================================
// Signature Options
// =============================================================================

/** Options for digitally signing a PDF. */
export interface PdfSignatureOptions {
  /** DER-encoded X.509 certificate. */
  certificate: Uint8Array;
  /** DER-encoded PKCS#8 private key. */
  privateKey: Uint8Array;
  /** Signer name (displayed in PDF viewers). */
  name?: string;
  /** Reason for signing. */
  reason?: string;
  /** Location of signing. */
  location?: string;
  /** Contact info. */
  contactInfo?: string;
}

// =============================================================================
// Constants
// =============================================================================

const DEFAULT_PAGE_WIDTH = 595.28; // A4
const DEFAULT_PAGE_HEIGHT = 841.89; // A4
const DEFAULT_FONT_SIZE = 12;
const DEFAULT_LINE_HEIGHT = 1.2;
const BLACK: PdfColor = { r: 0, g: 0, b: 0 };

// =============================================================================
// PdfPageBuilder
// =============================================================================

/**
 * Builder for a single PDF page.
 *
 * Provides methods for drawing text, shapes, and images at arbitrary positions.
 * All coordinates use PDF's coordinate system: origin at bottom-left, Y increases upward.
 */
export class PdfPageBuilder {
  /** @internal */
  readonly _stream = new PdfContentStream();
  /** @internal */
  readonly _width: number;
  /** @internal */
  readonly _height: number;
  /** @internal */
  readonly _images: DrawImageOptions[] = [];
  /** @internal */
  readonly _annotations: PageAnnotation[] = [];
  /** @internal */
  readonly _builderAnnotations: BuilderAnnotation[] = [];
  /** @internal */
  readonly _formFields: BuilderFormField[] = [];
  /** @internal */
  readonly _fontManager: FontManager;

  /** @internal */
  constructor(width: number, height: number, fontManager: FontManager) {
    this._width = width;
    this._height = height;
    this._fontManager = fontManager;
  }

  /** Page width in points. */
  get width(): number {
    return this._width;
  }

  /** Page height in points. */
  get height(): number {
    return this._height;
  }

  // ===========================================================================
  // Text
  // ===========================================================================

  /**
   * Draw text at a specific position.
   *
   * @param text - The text string to draw
   * @param options - Position, font, color, etc.
   */
  drawText(text: string, options: DrawTextOptions): this {
    const fontSize = options.fontSize ?? DEFAULT_FONT_SIZE;
    const color = options.color ?? BLACK;
    const lineHeightFactor = options.lineHeight ?? DEFAULT_LINE_HEIGHT;
    const bold = options.bold ?? false;
    const italic = options.italic ?? false;
    const fontFamily = options.fontFamily ?? "Helvetica";

    // Resolve font
    const resourceName = this._fontManager.resolveFont(fontFamily, bold, italic);
    const encodedText = this._fontManager.encodeText(text, resourceName);

    this._fontManager.trackText(text);

    if (options.maxWidth) {
      // Word-wrap (reuses the shared wrapTextLines from page-renderer)
      const measure = (s: string) => this._fontManager.measureText(s, resourceName, fontSize);
      const lines = wrapTextLines(text, measure, options.maxWidth);
      const leading = fontSize * lineHeightFactor;

      this._stream.save();
      this._stream.setFillColor(color);
      this._stream.beginText();
      this._stream.setFont(resourceName, fontSize);

      for (let i = 0; i < lines.length; i++) {
        const lineY = options.y - i * leading;
        this._stream.setTextMatrix(1, 0, 0, 1, options.x, lineY);

        const lineEncoded = this._fontManager.encodeText(lines[i], resourceName);
        if (lineEncoded) {
          this._stream.showTextHex(lineEncoded);
        } else {
          this._stream.showText(lines[i]);
        }
      }

      this._stream.endText();
      this._stream.restore();
    } else {
      // Single line
      this._stream.save();
      this._stream.setFillColor(color);
      this._stream.beginText();
      this._stream.setFont(resourceName, fontSize);
      this._stream.setTextMatrix(1, 0, 0, 1, options.x, options.y);

      if (encodedText) {
        this._stream.showTextHex(encodedText);
      } else {
        this._stream.showText(text);
      }

      this._stream.endText();
      this._stream.restore();
    }

    return this;
  }

  /**
   * Measure text width in points.
   */
  measureText(
    text: string,
    options?: { fontSize?: number; fontFamily?: string; bold?: boolean; italic?: boolean }
  ): number {
    const fontSize = options?.fontSize ?? DEFAULT_FONT_SIZE;
    const fontFamily = options?.fontFamily ?? "Helvetica";
    const bold = options?.bold ?? false;
    const italic = options?.italic ?? false;
    const resourceName = this._fontManager.resolveFont(fontFamily, bold, italic);
    return this._fontManager.measureText(text, resourceName, fontSize);
  }

  // ===========================================================================
  // Shapes
  // ===========================================================================

  /**
   * Draw a rectangle (filled and/or stroked).
   */
  drawRect(options: DrawRectOptions): this {
    this._stream.save();

    if (options.borderRadius && options.borderRadius > 0) {
      this._stream.roundedRect(
        options.x,
        options.y,
        options.width,
        options.height,
        options.borderRadius
      );
    } else {
      this._stream.rect(options.x, options.y, options.width, options.height);
    }

    this._paintPath(options.fill, options.stroke, options.lineWidth);
    this._stream.restore();
    return this;
  }

  /**
   * Draw a circle (filled and/or stroked).
   */
  drawCircle(options: DrawCircleOptions): this {
    this._stream.save();
    this._stream.circle(options.cx, options.cy, options.r);
    this._paintPath(options.fill, options.stroke, options.lineWidth);
    this._stream.restore();
    return this;
  }

  /**
   * Draw an ellipse (filled and/or stroked).
   */
  drawEllipse(options: DrawEllipseOptions): this {
    this._stream.save();
    this._stream.ellipse(options.cx, options.cy, options.rx, options.ry);
    this._paintPath(options.fill, options.stroke, options.lineWidth);
    this._stream.restore();
    return this;
  }

  /**
   * Draw a straight line.
   */
  drawLine(options: DrawLineOptions): this {
    const color = options.color ?? BLACK;
    const lineWidth = options.lineWidth ?? 1;

    this._stream.save();
    this._stream.setStrokeColor(color);
    this._stream.setLineWidth(lineWidth);
    if (options.dashPattern && options.dashPattern.length > 0) {
      this._stream.setDashPattern(options.dashPattern);
    }
    this._stream.moveTo(options.x1, options.y1);
    this._stream.lineTo(options.x2, options.y2);
    this._stream.stroke();
    this._stream.restore();
    return this;
  }

  /**
   * Draw a complex path from a list of path operations.
   */
  drawPath(ops: PathOp[], options?: DrawPathOptions): this {
    this._stream.save();

    for (const op of ops) {
      switch (op.op) {
        case "move":
          this._stream.moveTo(op.x, op.y);
          break;
        case "line":
          this._stream.lineTo(op.x, op.y);
          break;
        case "curve":
          this._stream.curveTo(op.x1, op.y1, op.x2, op.y2, op.x3, op.y3);
          break;
        case "close":
          this._stream.closePath();
          break;
      }
    }

    if (options?.closePath) {
      this._stream.closePath();
    }

    this._paintPath(options?.fill, options?.stroke, options?.lineWidth);
    this._stream.restore();
    return this;
  }

  // ===========================================================================
  // Images
  // ===========================================================================

  /**
   * Draw an image at a specific position.
   */
  drawImage(options: DrawImageOptions): this {
    this._images.push(options);
    // Image drawing is deferred to build time (needs object allocation)
    // We record a placeholder name based on index
    const imgName = `Im${this._images.length}`;
    this._stream.drawImage(imgName, options.x, options.y, options.width, options.height);
    return this;
  }

  // ===========================================================================
  // Annotations
  // ===========================================================================

  /**
   * Add an annotation to this page.
   *
   * Supports: Highlight, Underline, StrikeOut, Squiggly, Text (sticky note),
   * FreeText (inline text), and Stamp.
   */
  addAnnotation(options: AnnotationOptions): this {
    const entries: Array<[string, string]> = [];

    switch (options.type) {
      case "Highlight":
      case "Underline":
      case "StrikeOut":
      case "Squiggly": {
        const color =
          options.color ??
          (options.type === "Highlight" ? { r: 1, g: 1, b: 0 } : { r: 1, g: 0, b: 0 });
        entries.push(["C", `[${pdfNumber(color.r)} ${pdfNumber(color.g)} ${pdfNumber(color.b)}]`]);
        if (options.contents) {
          entries.push(["Contents", pdfString(options.contents)]);
        }
        if (options.author) {
          entries.push(["T", pdfString(options.author)]);
        }
        // QuadPoints
        const qp = options.quadPoints ?? [
          options.rect[0],
          options.rect[1],
          options.rect[2],
          options.rect[1],
          options.rect[0],
          options.rect[3],
          options.rect[2],
          options.rect[3]
        ];
        entries.push(["QuadPoints", `[${qp.map(v => pdfNumber(v)).join(" ")}]`]);
        break;
      }
      case "Text": {
        const color = options.color ?? { r: 1, g: 1, b: 0 };
        entries.push(["C", `[${pdfNumber(color.r)} ${pdfNumber(color.g)} ${pdfNumber(color.b)}]`]);
        if (options.contents) {
          entries.push(["Contents", pdfString(options.contents)]);
        }
        if (options.author) {
          entries.push(["T", pdfString(options.author)]);
        }
        entries.push(["Name", `/${options.iconName ?? "Note"}`]);
        if (options.open) {
          entries.push(["Open", "true"]);
        }
        break;
      }
      case "FreeText": {
        const fontSize = options.fontSize ?? 12;
        const color = options.color ?? BLACK;
        entries.push(["Contents", pdfString(options.contents)]);
        entries.push([
          "DA",
          pdfString(
            `/Helv ${pdfNumber(fontSize)} Tf ${pdfNumber(color.r)} ${pdfNumber(color.g)} ${pdfNumber(color.b)} rg`
          )
        ]);
        if (options.borderColor) {
          const bc = options.borderColor;
          entries.push(["C", `[${pdfNumber(bc.r)} ${pdfNumber(bc.g)} ${pdfNumber(bc.b)}]`]);
        }
        if (options.author) {
          entries.push(["T", pdfString(options.author)]);
        }
        break;
      }
      case "Stamp": {
        entries.push(["Name", `/${options.stampName ?? "Draft"}`]);
        if (options.color) {
          const c = options.color;
          entries.push(["C", `[${pdfNumber(c.r)} ${pdfNumber(c.g)} ${pdfNumber(c.b)}]`]);
        }
        if (options.contents) {
          entries.push(["Contents", pdfString(options.contents)]);
        }
        if (options.author) {
          entries.push(["T", pdfString(options.author)]);
        }
        break;
      }
    }

    this._builderAnnotations.push({
      subtype: options.type,
      rect: options.rect,
      entries
    });
    return this;
  }

  // ===========================================================================
  // Form Fields
  // ===========================================================================

  /**
   * Add a form field to this page.
   *
   * Supports: text input, checkbox, dropdown (combo box), and radio button groups.
   */
  addFormField(options: FormFieldOptions): this {
    this._formFields.push({ options });
    return this;
  }

  // ===========================================================================
  // SVG Path
  // ===========================================================================

  /**
   * Draw an SVG path from a `d` attribute string.
   *
   * Supports all SVG path commands: M, L, H, V, C, S, Q, T, A, Z
   * (both absolute and relative).
   *
   * @param d - The SVG path data string (e.g., "M10 10 L90 90 Z")
   * @param options - Fill/stroke options
   */
  drawSvgPath(d: string, options?: DrawPathOptions): this {
    const ops = parseSvgPath(d);
    return this.drawPath(ops, options);
  }

  // ===========================================================================
  // Raw content stream access
  // ===========================================================================

  /**
   * Get the raw content stream for advanced operations.
   * Use this when the high-level API doesn't cover your use case.
   */
  getContentStream(): PdfContentStream {
    return this._stream;
  }

  // ===========================================================================
  // Internal Helpers
  // ===========================================================================

  /** @internal */
  private _paintPath(
    fill: PdfColor | undefined,
    stroke: PdfColor | undefined,
    lineWidth: number | undefined
  ): void {
    const hasFill = fill !== undefined;
    const hasStroke = stroke !== undefined;

    if (hasFill) {
      this._stream.setFillColor(fill);
    }
    if (hasStroke) {
      this._stream.setStrokeColor(stroke);
      this._stream.setLineWidth(lineWidth ?? 1);
    }

    if (hasFill && hasStroke) {
      this._stream.fillAndStroke();
    } else if (hasFill) {
      this._stream.fill();
    } else if (hasStroke) {
      this._stream.stroke();
    } else {
      // Default: stroke with black, 1pt
      this._stream.setStrokeColor(BLACK);
      this._stream.setLineWidth(1);
      this._stream.stroke();
    }
  }
}

// =============================================================================
// PdfDocumentBuilder
// =============================================================================

/**
 * Builder for constructing multi-page PDF documents with free-form content.
 *
 * Provides fine-grained control over text positioning, vector graphics,
 * and page management — complementing the table-oriented `pdf()` function.
 */
export class PdfDocumentBuilder {
  private _pages: PdfPageBuilder[] = [];
  private _bookmarks: BookmarkNode[] = [];
  private _fontManager = new FontManager();
  private _metadata: DocumentMetadata = {};
  private _encryption: PdfExportOptions["encryption"];
  private _embeddedFont: Uint8Array | null = null;
  private _pdfA = false;
  private _signatureOptions: PdfSignatureOptions | null = null;

  /**
   * Add a new blank page to the document.
   *
   * @param options - Page dimensions. Default: A4 (595.28 x 841.89 points).
   * @returns A PdfPageBuilder for the new page.
   */
  addPage(options?: PageOptions): PdfPageBuilder {
    const width = options?.width ?? DEFAULT_PAGE_WIDTH;
    const height = options?.height ?? DEFAULT_PAGE_HEIGHT;
    const page = new PdfPageBuilder(width, height, this._fontManager);
    this._pages.push(page);
    return page;
  }

  /**
   * Set document metadata (title, author, etc.).
   */
  setMetadata(metadata: DocumentMetadata): this {
    this._metadata = metadata;
    return this;
  }

  /**
   * Set encryption options (AES-256).
   */
  setEncryption(encryption: PdfExportOptions["encryption"]): this {
    this._encryption = encryption;
    return this;
  }

  /**
   * Embed a TrueType font for Unicode/CJK support.
   *
   * @param fontBytes - Raw .ttf file bytes
   */
  embedFont(fontBytes: Uint8Array): this {
    this._embeddedFont = fontBytes;
    return this;
  }

  /**
   * Enable PDF/A compliance output.
   *
   * Currently supports PDF/A-1b (ISO 19005-1, Level B — visual appearance
   * preservation). When enabled, `build()` will:
   *
   * - Set PDF version to 1.4
   * - Write XMP metadata with `pdfaid:part=1` and `pdfaid:conformance=B`
   * - Write OutputIntents with an embedded sRGB ICC profile
   * - Add `/MarkInfo << /Marked true >>` to the catalog
   *
   * **Limitation:** Type1 base fonts (Helvetica, Times-Roman, Courier, etc.)
   * are not embedded. For strict PDF/A-1b font compliance, use `embedFont()`
   * to embed a TrueType font.
   *
   * @param _level - Conformance level. Currently only "1b" is supported.
   */
  setPdfACompliance(_level?: "1b"): this {
    this._pdfA = true;
    return this;
  }

  /**
   * Digitally sign the PDF during `build()`.
   *
   * When set, `build()` will:
   * 1. Embed a signature dictionary with placeholder in the PDF
   * 2. Compute the byte ranges and sign with RSA PKCS#1 v1.5 + SHA-256
   * 3. Return the fully signed PDF bytes
   *
   * @param options - Certificate, private key, and optional signer metadata
   *
   * @example
   * ```typescript
   * doc.sign({
   *   certificate: certDerBytes,
   *   privateKey: pkcs8DerBytes,
   *   name: "John Doe",
   *   reason: "Document approval"
   * });
   * const signedPdf = await doc.build();
   * ```
   */
  sign(options: PdfSignatureOptions): this {
    this._signatureOptions = options;
    return this;
  }

  // ===========================================================================
  // Bookmarks & Table of Contents
  // ===========================================================================

  /**
   * Add a bookmark (PDF outline entry) pointing to a specific page.
   *
   * @param title - Bookmark display title
   * @param pageIndex - Zero-based page index
   * @param parent - Index of a previously added top-level bookmark to nest under (zero-based in insertion order). Omit for top-level.
   * @returns this for chaining
   */
  addBookmark(title: string, pageIndex: number, parent?: number): this {
    const node: BookmarkNode = { title, pageIndex, children: [] };
    if (parent !== undefined) {
      if (parent < 0 || parent >= this._bookmarks.length) {
        throw new RangeError(
          `Bookmark parent index ${parent} is out of range (0..${this._bookmarks.length - 1})`
        );
      }
      this._bookmarks[parent].children.push(node);
    } else {
      this._bookmarks.push(node);
    }
    return this;
  }

  /**
   * Generate a table of contents page with clickable entries.
   *
   * Each entry displays the bookmark title and a right-aligned page number,
   * connected by a dotted leader. Entries link to their target pages.
   *
   * @param options - TOC formatting options
   * @returns The created PdfPageBuilder for further customization
   */
  generateTableOfContents(options?: TocOptions): PdfPageBuilder {
    const tocTitle = options?.title ?? "Table of Contents";
    const fontSize = options?.fontSize ?? DEFAULT_FONT_SIZE;
    const indent = options?.indent ?? 20;

    let page = this.addPage();
    const firstPage = page;

    const titleFontSize = fontSize + 6;
    const marginLeft = 72;
    const marginRight = 72;
    const marginBottom = 72;
    const usableWidth = page._width - marginLeft - marginRight;
    let cursorY = page._height - 72;

    // Draw TOC title
    page.drawText(tocTitle, {
      x: marginLeft,
      y: cursorY,
      fontSize: titleFontSize,
      bold: true
    });
    cursorY -= titleFontSize * 1.8;

    // Draw a separator line under the title
    page.drawLine({
      x1: marginLeft,
      y1: cursorY + fontSize * 0.4,
      x2: page._width - marginRight,
      y2: cursorY + fontSize * 0.4,
      color: { r: 0.6, g: 0.6, b: 0.6 },
      lineWidth: 0.5
    });
    cursorY -= fontSize * 0.6;

    // Flatten bookmarks with depth info for rendering
    const entries: Array<{ title: string; pageIndex: number; depth: number }> = [];
    const flattenBookmarks = (nodes: BookmarkNode[], depth: number): void => {
      for (const node of nodes) {
        entries.push({ title: node.title, pageIndex: node.pageIndex, depth });
        flattenBookmarks(node.children, depth + 1);
      }
    };
    flattenBookmarks(this._bookmarks, 0);

    const lineHeight = fontSize * 1.6;

    for (const entry of entries) {
      if (cursorY < marginBottom) {
        // Overflow — create a continuation page
        page = this.addPage();
        cursorY = page._height - 72;
      }

      const entryX = marginLeft + entry.depth * indent;
      // Measure title and page number
      const pageNumStr = String(entry.pageIndex + 1);
      const titleWidth = page.measureText(entry.title, { fontSize });
      const pageNumWidth = page.measureText(pageNumStr, { fontSize });
      const dotWidth = page.measureText(".", { fontSize });

      // Draw title text
      page.drawText(entry.title, {
        x: entryX,
        y: cursorY,
        fontSize
      });

      // Draw page number (right-aligned)
      const pageNumX = marginLeft + usableWidth - pageNumWidth;
      page.drawText(pageNumStr, {
        x: pageNumX,
        y: cursorY,
        fontSize
      });

      // Draw dot leaders between title and page number
      const dotsStartX = entryX + titleWidth + dotWidth;
      const dotsEndX = pageNumX - dotWidth;
      if (dotsEndX > dotsStartX && dotWidth > 0) {
        const dotSpacing = dotWidth * 2;
        let dotX = dotsStartX;
        const dots: string[] = [];
        while (dotX + dotWidth <= dotsEndX) {
          dots.push(".");
          dotX += dotSpacing;
        }
        if (dots.length > 0) {
          page.drawText(dots.join(" "), {
            x: dotsStartX,
            y: cursorY,
            fontSize,
            color: { r: 0.6, g: 0.6, b: 0.6 }
          });
        }
      }

      // Record a link annotation for this entry
      const annotY = cursorY - fontSize * 0.3;
      page._annotations.push({
        rect: [entryX, annotY, marginLeft + usableWidth, annotY + fontSize * 1.2],
        destPageIndex: entry.pageIndex
      });

      cursorY -= lineHeight;
    }

    return firstPage;
  }

  /** Get all pages. */
  get pages(): readonly PdfPageBuilder[] {
    return this._pages;
  }

  /**
   * Build the final PDF document.
   *
   * @returns The PDF file as Uint8Array.
   */
  async build(): Promise<Uint8Array> {
    const writer = new PdfWriter();

    // PDF/A-1b requires PDF 1.4
    if (this._pdfA) {
      writer.setVersion("1.4");
    }

    // Register embedded font if provided
    if (this._embeddedFont) {
      const ttfFont = parseTtf(this._embeddedFont);
      this._fontManager.registerEmbeddedFont(ttfFont);
    }

    // Write font resources
    const fontObjectMap = await this._fontManager.writeFontResources(writer);
    const fontDictStr = this._fontManager.buildFontDictString(fontObjectMap);

    // Build each page
    const pageObjNums: number[] = [];
    const pagesTreeObjNum = writer.allocObject();

    // Pre-allocate page object numbers so annotations can reference them
    for (let i = 0; i < this._pages.length; i++) {
      pageObjNums.push(writer.allocObject());
    }

    // Track content and resource refs per page for page dict construction
    const pageContentRefs: number[] = [];
    const pageResourceRefs: number[] = [];
    const allFormFieldRefs: number[] = [];

    // Write pages with their content, resources, and annotations
    for (let i = 0; i < this._pages.length; i++) {
      const page = this._pages[i];

      // Write image XObjects for this page
      const imageXObjectMap = new Map<string, number>();
      for (let j = 0; j < page._images.length; j++) {
        const img = page._images[j];
        const imgName = `Im${j + 1}`;
        const imgObjNum = this._writeImageXObject(writer, img);
        imageXObjectMap.set(imgName, imgObjNum);
      }

      // Build XObject dict string
      let xobjDictStr = "";
      if (imageXObjectMap.size > 0) {
        const entries = [...imageXObjectMap.entries()]
          .map(([name, objNum]) => `/${name} ${pdfRef(objNum)}`)
          .join(" ");
        xobjDictStr = `<< ${entries} >>`;
      }

      // Write content stream
      const contentObjNum = writer.allocObject();
      const contentDict = new PdfDict();
      writer.addStreamObject(contentObjNum, contentDict, page._stream);
      pageContentRefs.push(contentObjNum);

      // Write resources
      const resourcesObjNum = writer.allocObject();
      let resourcesStr = "<< ";
      if (fontDictStr) {
        resourcesStr += `/Font ${fontDictStr} `;
      }
      if (xobjDictStr) {
        resourcesStr += `/XObject ${xobjDictStr} `;
      }
      resourcesStr += ">>";
      writer.addObject(resourcesObjNum, resourcesStr);
      pageResourceRefs.push(resourcesObjNum);

      // Write link annotations
      const annotRefs: number[] = [];
      for (const annot of page._annotations) {
        const destPageObj = pageObjNums[annot.destPageIndex];
        if (destPageObj === undefined) {
          continue;
        }

        const annotObjNum = writer.allocObject();
        const rect = `[${pdfNumber(annot.rect[0])} ${pdfNumber(annot.rect[1])} ${pdfNumber(annot.rect[2])} ${pdfNumber(annot.rect[3])}]`;
        const annotDict = new PdfDict()
          .set("Type", "/Annot")
          .set("Subtype", "/Link")
          .set("Rect", rect)
          .set("Border", "[0 0 0]")
          .set("Dest", `[${pdfRef(destPageObj)} /Fit]`);
        writer.addObject(annotObjNum, annotDict);
        annotRefs.push(annotObjNum);
      }

      // Write builder-created annotations (Highlight, Text, FreeText, Stamp, etc.)
      for (const annot of page._builderAnnotations) {
        const annotObjNum = writer.allocObject();
        const rect = `[${pdfNumber(annot.rect[0])} ${pdfNumber(annot.rect[1])} ${pdfNumber(annot.rect[2])} ${pdfNumber(annot.rect[3])}]`;
        const annotDict = new PdfDict()
          .set("Type", "/Annot")
          .set("Subtype", `/${annot.subtype}`)
          .set("Rect", rect)
          .set("F", "4"); // Print flag — annotation is printable
        for (const [key, value] of annot.entries) {
          annotDict.set(key, value);
        }
        writer.addObject(annotObjNum, annotDict);
        annotRefs.push(annotObjNum);
      }

      // Write form field widget annotations
      for (const field of page._formFields) {
        const { fieldRefs, annotRefs: fieldAnnotRefs } = this._writeFormFieldAnnotation(
          writer,
          field.options,
          pageObjNums[i]
        );
        annotRefs.push(...fieldAnnotRefs);
        allFormFieldRefs.push(...fieldRefs);
      }

      // Write page object (using pre-allocated obj num)
      const pageObjNum = pageObjNums[i];
      const mediaBox = `[0 0 ${pdfNumber(page._width)} ${pdfNumber(page._height)}]`;
      const pageDict = new PdfDict()
        .set("Type", "/Page")
        .set("Parent", pdfRef(pagesTreeObjNum))
        .set("MediaBox", mediaBox)
        .set("Contents", pdfRef(contentObjNum))
        .set("Resources", pdfRef(resourcesObjNum));
      if (annotRefs.length > 0) {
        pageDict.set("Annots", "[" + annotRefs.map(r => pdfRef(r)).join(" ") + "]");
      }
      writer.addObject(pageObjNum, pageDict);
    }

    // Ensure at least one page
    if (pageObjNums.length === 0) {
      const emptyContentObjNum = writer.allocObject();
      writer.addStreamObject(emptyContentObjNum, new PdfDict(), new Uint8Array(0));
      const emptyResourcesObjNum = writer.allocObject();
      writer.addObject(emptyResourcesObjNum, "<< >>");
      const pageObjNum = writer.allocObject();
      const emptyPageDict = new PdfDict()
        .set("Type", "/Page")
        .set("Parent", pdfRef(pagesTreeObjNum))
        .set("MediaBox", `[0 0 ${pdfNumber(DEFAULT_PAGE_WIDTH)} ${pdfNumber(DEFAULT_PAGE_HEIGHT)}]`)
        .set("Contents", pdfRef(emptyContentObjNum))
        .set("Resources", pdfRef(emptyResourcesObjNum));
      writer.addObject(pageObjNum, emptyPageDict);
      pageObjNums.push(pageObjNum);
    }

    // Pages tree
    const kidsStr = pageObjNums.map(n => pdfRef(n)).join(" ");
    writer.addObject(
      pagesTreeObjNum,
      new PdfDict()
        .set("Type", "/Pages")
        .set("Kids", `[${kidsStr}]`)
        .set("Count", String(pageObjNums.length))
    );

    // Build outline tree from bookmarks
    let outlinesRef: number | undefined;
    if (this._bookmarks.length > 0) {
      outlinesRef = this._buildOutlines(writer, pageObjNums);
    }

    // Catalog — with optional PDF/A entries
    const catalogExtras: Array<[string, string]> = [];

    if (this._pdfA) {
      // Write XMP metadata stream
      const xmpObjNum = writePdfAMetadata(writer, this._metadata);
      catalogExtras.push(["Metadata", pdfRef(xmpObjNum)]);

      // Write OutputIntents with sRGB ICC profile
      const intentObjNum = writePdfAOutputIntent(writer);
      catalogExtras.push(["OutputIntents", `[${pdfRef(intentObjNum)}]`]);

      // Mark as tagged (minimal structural compliance)
      catalogExtras.push(["MarkInfo", "<< /Marked true >>"]);
    }

    // Build catalog — handle three cases:
    // 1. Simple: no form fields, no signing → addCatalog()
    // 2. Form fields only → rebuild catalog with AcroForm
    // 3. Signing (with or without form fields) → signing path builds the catalog
    const needsCustomCatalog = allFormFieldRefs.length > 0 || this._signatureOptions;

    if (!needsCustomCatalog) {
      writer.addCatalog(pagesTreeObjNum, {
        outlinesRef,
        extraEntries: catalogExtras.length > 0 ? catalogExtras : undefined
      });
    }

    // AcroForm — if any pages have form fields (and not signing — signing path builds its own catalog)
    if (allFormFieldRefs.length > 0 && !this._signatureOptions) {
      const catalogObjNum = writer.allocObject();
      const catalogDict = new PdfDict()
        .set("Type", "/Catalog")
        .set("Pages", pdfRef(pagesTreeObjNum));
      if (outlinesRef) {
        catalogDict.set("Outlines", pdfRef(outlinesRef));
        catalogDict.set("PageMode", "/UseOutlines");
      }
      for (const [key, value] of catalogExtras) {
        catalogDict.set(key, value);
      }
      const fieldsStr = allFormFieldRefs.map(r => pdfRef(r)).join(" ");
      const acroFormStr = `<< /Fields [${fieldsStr}] /NeedAppearances true /DR << /Font << /Helv << /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >> >> >> /DA (/Helv 0 Tf 0 g) >>`;
      catalogDict.set("AcroForm", acroFormStr);
      writer.addObject(catalogObjNum, catalogDict);
      writer.setCatalog(catalogObjNum);
    }

    // Info dict
    if (
      this._metadata.title ||
      this._metadata.author ||
      this._metadata.subject ||
      this._metadata.creator
    ) {
      writer.addInfoDict(this._metadata);
    }

    // Encryption
    if (this._encryption) {
      const encState = initEncryption(this._encryption);
      writer.setEncryption(encState);
    }

    // If signing is requested, we need to:
    // 1. Add the signature dict placeholder + widget to the PDF
    // 2. Build the PDF bytes
    // 3. Call signPdf() to fill in the real signature
    if (this._signatureOptions) {
      const { buildSignatureDictPlaceholder, signPdf } = await import("../core/digital-signature");

      const { dictString } = buildSignatureDictPlaceholder({
        name: this._signatureOptions.name,
        reason: this._signatureOptions.reason,
        location: this._signatureOptions.location,
        contactInfo: this._signatureOptions.contactInfo
      });

      // Write signature dict as indirect object
      const sigDictObjNum = writer.allocObject();
      writer.addObject(sigDictObjNum, dictString);

      // Write signature widget annotation
      const sigWidgetObjNum = writer.allocObject();
      const sigWidgetDict = new PdfDict()
        .set("Type", "/Annot")
        .set("Subtype", "/Widget")
        .set("FT", "/Sig")
        .set("Rect", "[0 0 0 0]")
        .set("T", pdfString("Signature1"))
        .set("V", pdfRef(sigDictObjNum))
        .set("F", "4");
      writer.addObject(sigWidgetObjNum, sigWidgetDict);

      // Patch catalog to include AcroForm with SigFlags
      // We need to rebuild the catalog with AcroForm
      const sigCatalogObjNum = writer.allocObject();
      const sigCatalogDict = new PdfDict()
        .set("Type", "/Catalog")
        .set("Pages", pdfRef(pagesTreeObjNum));
      if (outlinesRef) {
        sigCatalogDict.set("Outlines", pdfRef(outlinesRef));
        sigCatalogDict.set("PageMode", "/UseOutlines");
      }
      for (const [key, value] of catalogExtras) {
        sigCatalogDict.set(key, value);
      }

      // Merge existing form field refs with signature widget
      const allFields = [...allFormFieldRefs, sigWidgetObjNum];
      const fieldsStr = allFields.map(r => pdfRef(r)).join(" ");
      sigCatalogDict.set("AcroForm", `<< /Fields [${fieldsStr}] /SigFlags 3 >>`);

      // Add signature widget to first page's annotations
      // (We need to patch the first page dict to include the widget in /Annots)
      // For simplicity, add it as a document-level field (already in AcroForm /Fields)
      writer.addObject(sigCatalogObjNum, sigCatalogDict);
      writer.setCatalog(sigCatalogObjNum);

      const pdfWithPlaceholder = writer.build();

      // Sign the PDF
      return signPdf(
        pdfWithPlaceholder,
        this._signatureOptions.certificate,
        this._signatureOptions.privateKey
      );
    }

    return writer.build();
  }

  // ===========================================================================
  // Internal Helpers
  // ===========================================================================

  /** @internal */
  private _writeImageXObject(writer: PdfWriter, img: DrawImageOptions): number {
    return writeImageXObject(writer, img.data, img.format);
  }

  /**
   * Build a nested PDF outline (bookmark) tree.
   * @internal
   */
  private _buildOutlines(writer: PdfWriter, pageObjNums: number[]): number {
    const outlinesObjNum = writer.allocObject();

    // Allocate object numbers for all nodes (pre-order traversal)
    const allNodes: Array<{
      node: BookmarkNode;
      objNum: number;
      parentObjNum: number;
      depth: number;
    }> = [];

    const allocNodes = (nodes: BookmarkNode[], parentObjNum: number, depth: number): void => {
      for (const node of nodes) {
        const objNum = writer.allocObject();
        allNodes.push({ node, objNum, parentObjNum, depth });
        allocNodes(node.children, objNum, depth + 1);
      }
    };
    allocNodes(this._bookmarks, outlinesObjNum, 0);

    // Group children by parent for sibling linkage
    const childrenByParent = new Map<number, typeof allNodes>();
    for (const entry of allNodes) {
      const siblings = childrenByParent.get(entry.parentObjNum);
      if (siblings) {
        siblings.push(entry);
      } else {
        childrenByParent.set(entry.parentObjNum, [entry]);
      }
    }

    // Count all descendants (including self) for each node
    const countDescendants = (node: BookmarkNode): number => {
      let count = 1;
      for (const child of node.children) {
        count += countDescendants(child);
      }
      return count;
    };

    // Write each outline item
    for (const entry of allNodes) {
      const { node, objNum, parentObjNum } = entry;
      const pageObjNum = pageObjNums[node.pageIndex];
      if (pageObjNum === undefined) {
        continue;
      }

      const dict = new PdfDict()
        .set("Title", pdfString(node.title))
        .set("Parent", pdfRef(parentObjNum))
        .set("Dest", `[${pdfRef(pageObjNum)} /Fit]`);

      // Sibling linkage
      const siblings = childrenByParent.get(parentObjNum) ?? [];
      const idx = siblings.indexOf(entry);
      if (idx > 0) {
        dict.set("Prev", pdfRef(siblings[idx - 1].objNum));
      }
      if (idx < siblings.length - 1) {
        dict.set("Next", pdfRef(siblings[idx + 1].objNum));
      }

      // Children linkage
      const children = childrenByParent.get(objNum);
      if (children && children.length > 0) {
        dict.set("First", pdfRef(children[0].objNum));
        dict.set("Last", pdfRef(children[children.length - 1].objNum));
        // Negative count = initially closed, positive = initially open
        const totalChildren = node.children.reduce((sum, c) => sum + countDescendants(c), 0);
        dict.set("Count", String(-totalChildren));
      }

      writer.addObject(objNum, dict);
    }

    // Write outlines root
    const topLevel = childrenByParent.get(outlinesObjNum) ?? [];
    const totalCount = this._bookmarks.length;
    const outlinesDict = new PdfDict().set("Type", "/Outlines").set("Count", String(totalCount));
    if (topLevel.length > 0) {
      outlinesDict.set("First", pdfRef(topLevel[0].objNum));
      outlinesDict.set("Last", pdfRef(topLevel[topLevel.length - 1].objNum));
    }
    writer.addObject(outlinesObjNum, outlinesDict);

    return outlinesObjNum;
  }

  /**
   * Write form field annotation(s) as indirect objects.
   * @internal
   */
  private _writeFormFieldAnnotation(
    writer: PdfWriter,
    options: FormFieldOptions,
    pageObjNum: number
  ): { fieldRefs: number[]; annotRefs: number[] } {
    const fieldRefs: number[] = [];
    const annotRefs: number[] = [];

    if (options.type === "radio") {
      // Radio group: one parent field + one widget per button
      const parentObjNum = writer.allocObject();
      const childRefs: number[] = [];
      let ff = 1 << 15; // /Ff bit 16 = Radio
      ff |= 1 << 14; // /Ff bit 15 = NoToggleToOff
      if (options.readOnly) {
        ff |= 1;
      }
      if (options.required) {
        ff |= 1 << 1;
      }

      for (const btn of options.buttons) {
        const childObjNum = writer.allocObject();
        const rect = `[${btn.rect.map(v => pdfNumber(v)).join(" ")}]`;
        const isSelected = options.selected === btn.value;
        const apState = isSelected ? `/${btn.value}` : "/Off";

        const childDict = new PdfDict()
          .set("Type", "/Annot")
          .set("Subtype", "/Widget")
          .set("Rect", rect)
          .set("Parent", pdfRef(parentObjNum))
          .set("AS", apState)
          .set("AP", `<< /N << /${btn.value} null /Off null >> >>`);
        writer.addObject(childObjNum, childDict);
        childRefs.push(childObjNum);
      }

      const parentDict = new PdfDict()
        .set("FT", "/Btn")
        .set("T", pdfString(options.name))
        .set("Ff", String(ff))
        .set("Kids", `[${childRefs.map(r => pdfRef(r)).join(" ")}]`);
      if (options.selected) {
        parentDict.set("V", `/${options.selected}`);
      }
      writer.addObject(parentObjNum, parentDict);
      // Parent goes into AcroForm /Fields; children go into page /Annots
      fieldRefs.push(parentObjNum);
      annotRefs.push(...childRefs);
      return { fieldRefs, annotRefs };
    }

    // Single-widget fields: text, checkbox, dropdown
    const objNum = writer.allocObject();
    const r = options.rect;
    const rect = `[${pdfNumber(r[0])} ${pdfNumber(r[1])} ${pdfNumber(r[2])} ${pdfNumber(r[3])}]`;

    const dict = new PdfDict()
      .set("Type", "/Annot")
      .set("Subtype", "/Widget")
      .set("Rect", rect)
      .set("T", pdfString(options.name))
      .set("P", pdfRef(pageObjNum));

    let ff = 0;
    if (options.readOnly) {
      ff |= 1;
    }
    if (options.required) {
      ff |= 1 << 1;
    }

    switch (options.type) {
      case "text": {
        dict.set("FT", "/Tx");
        if (options.multiline) {
          ff |= 1 << 12;
        }
        if (options.password) {
          ff |= 1 << 13;
        }
        if (options.maxLength !== undefined) {
          dict.set("MaxLen", String(options.maxLength));
        }
        if (options.value) {
          dict.set("V", pdfString(options.value));
        }
        // Default appearance
        dict.set("DA", pdfString("/Helv 12 Tf 0 g"));
        break;
      }
      case "checkbox": {
        dict.set("FT", "/Btn");
        const checked = options.checked ?? false;
        dict.set("V", checked ? "/Yes" : "/Off");
        dict.set("AS", checked ? "/Yes" : "/Off");
        break;
      }
      case "dropdown": {
        dict.set("FT", "/Ch");
        ff |= 1 << 17; // Combo flag
        if (options.editable) {
          ff |= 1 << 18;
        }
        const optStr = options.options.map(o => pdfString(o)).join(" ");
        dict.set("Opt", `[${optStr}]`);
        if (options.value) {
          dict.set("V", pdfString(options.value));
        }
        dict.set("DA", pdfString("/Helv 12 Tf 0 g"));
        break;
      }
    }

    if (ff !== 0) {
      dict.set("Ff", String(ff));
    }

    writer.addObject(objNum, dict);
    // Single-widget fields go into both /Annots and /Fields
    fieldRefs.push(objNum);
    annotRefs.push(objNum);
    return { fieldRefs, annotRefs };
  }
}

// =============================================================================
// SVG Path Parser
// =============================================================================

/**
 * Parse an SVG path `d` attribute into PathOp array.
 *
 * Supports all SVG path commands:
 * - M/m (moveTo), L/l (lineTo), H/h (horizontal), V/v (vertical)
 * - C/c (cubic Bézier), S/s (smooth cubic)
 * - Q/q (quadratic Bézier), T/t (smooth quadratic)
 * - A/a (elliptical arc), Z/z (close)
 *
 * Arc commands are approximated with cubic Bézier curves.
 */
export function parseSvgPath(d: string): PathOp[] {
  const ops: PathOp[] = [];
  // Tokenize: split into commands + numbers
  const tokens = d.match(/[a-zA-Z]|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/g);
  if (!tokens) {
    return ops;
  }

  let i = 0;
  let cx = 0; // current x
  let cy = 0; // current y
  let sx = 0; // subpath start x
  let sy = 0; // subpath start y
  let lastCmd = "";
  // For smooth curves: last control point
  let lastCpX = 0;
  let lastCpY = 0;

  const num = (): number => {
    if (i >= tokens.length) {
      return 0;
    }
    return parseFloat(tokens[i++]);
  };

  const isNum = (): boolean => {
    if (i >= tokens.length) {
      return false;
    }
    const c = tokens[i].charCodeAt(0);
    return c === 0x2d || c === 0x2b || c === 0x2e || (c >= 0x30 && c <= 0x39);
  };

  while (i < tokens.length) {
    let cmd = tokens[i];
    if (/[a-zA-Z]/.test(cmd)) {
      i++;
    } else {
      // Implicit repeat of last command (except M becomes L, m becomes l)
      cmd = lastCmd === "M" ? "L" : lastCmd === "m" ? "l" : lastCmd;
    }

    switch (cmd) {
      case "M":
        cx = num();
        cy = num();
        ops.push({ op: "move", x: cx, y: cy });
        sx = cx;
        sy = cy;
        lastCmd = "M";
        while (isNum()) {
          cx = num();
          cy = num();
          ops.push({ op: "line", x: cx, y: cy });
        }
        break;
      case "m":
        cx += num();
        cy += num();
        ops.push({ op: "move", x: cx, y: cy });
        sx = cx;
        sy = cy;
        lastCmd = "m";
        while (isNum()) {
          cx += num();
          cy += num();
          ops.push({ op: "line", x: cx, y: cy });
        }
        break;
      case "L":
        do {
          cx = num();
          cy = num();
          ops.push({ op: "line", x: cx, y: cy });
        } while (isNum());
        lastCmd = "L";
        break;
      case "l":
        do {
          const dx = num();
          const dy = num();
          cx += dx;
          cy += dy;
          ops.push({ op: "line", x: cx, y: cy });
        } while (isNum());
        lastCmd = "l";
        break;
      case "H":
        do {
          cx = num();
          ops.push({ op: "line", x: cx, y: cy });
        } while (isNum());
        lastCmd = "H";
        break;
      case "h":
        do {
          cx += num();
          ops.push({ op: "line", x: cx, y: cy });
        } while (isNum());
        lastCmd = "h";
        break;
      case "V":
        do {
          cy = num();
          ops.push({ op: "line", x: cx, y: cy });
        } while (isNum());
        lastCmd = "V";
        break;
      case "v":
        do {
          cy += num();
          ops.push({ op: "line", x: cx, y: cy });
        } while (isNum());
        lastCmd = "v";
        break;
      case "C":
        do {
          const x1 = num(),
            y1 = num(),
            x2 = num(),
            y2 = num(),
            x = num(),
            y = num();
          ops.push({ op: "curve", x1, y1, x2, y2, x3: x, y3: y });
          lastCpX = x2;
          lastCpY = y2;
          cx = x;
          cy = y;
        } while (isNum());
        lastCmd = "C";
        break;
      case "c":
        do {
          const x1 = cx + num(),
            y1 = cy + num(),
            x2 = cx + num(),
            y2 = cy + num();
          const x = cx + num(),
            y = cy + num();
          ops.push({ op: "curve", x1, y1, x2, y2, x3: x, y3: y });
          lastCpX = x2;
          lastCpY = y2;
          cx = x;
          cy = y;
        } while (isNum());
        lastCmd = "c";
        break;
      case "S":
        do {
          const rx =
            lastCmd === "S" || lastCmd === "s" || lastCmd === "C" || lastCmd === "c"
              ? 2 * cx - lastCpX
              : cx;
          const ry =
            lastCmd === "S" || lastCmd === "s" || lastCmd === "C" || lastCmd === "c"
              ? 2 * cy - lastCpY
              : cy;
          const x2 = num(),
            y2 = num(),
            x = num(),
            y = num();
          ops.push({ op: "curve", x1: rx, y1: ry, x2, y2, x3: x, y3: y });
          lastCpX = x2;
          lastCpY = y2;
          cx = x;
          cy = y;
          lastCmd = "S";
        } while (isNum());
        break;
      case "s":
        do {
          const rx =
            lastCmd === "S" || lastCmd === "s" || lastCmd === "C" || lastCmd === "c"
              ? 2 * cx - lastCpX
              : cx;
          const ry =
            lastCmd === "S" || lastCmd === "s" || lastCmd === "C" || lastCmd === "c"
              ? 2 * cy - lastCpY
              : cy;
          const x2 = cx + num(),
            y2 = cy + num(),
            x = cx + num(),
            y = cy + num();
          ops.push({ op: "curve", x1: rx, y1: ry, x2, y2, x3: x, y3: y });
          lastCpX = x2;
          lastCpY = y2;
          cx = x;
          cy = y;
          lastCmd = "s";
        } while (isNum());
        break;
      case "Q":
        do {
          const qx = num(),
            qy = num(),
            x = num(),
            y = num();
          // Convert quadratic to cubic: CP1 = P0 + 2/3*(QP-P0), CP2 = P1 + 2/3*(QP-P1)
          const c1x = cx + (2 / 3) * (qx - cx),
            c1y = cy + (2 / 3) * (qy - cy);
          const c2x = x + (2 / 3) * (qx - x),
            c2y = y + (2 / 3) * (qy - y);
          ops.push({ op: "curve", x1: c1x, y1: c1y, x2: c2x, y2: c2y, x3: x, y3: y });
          lastCpX = qx;
          lastCpY = qy;
          cx = x;
          cy = y;
        } while (isNum());
        lastCmd = "Q";
        break;
      case "q":
        do {
          const qx = cx + num(),
            qy = cy + num(),
            x = cx + num(),
            y = cy + num();
          const c1x = cx + (2 / 3) * (qx - cx),
            c1y = cy + (2 / 3) * (qy - cy);
          const c2x = x + (2 / 3) * (qx - x),
            c2y = y + (2 / 3) * (qy - y);
          ops.push({ op: "curve", x1: c1x, y1: c1y, x2: c2x, y2: c2y, x3: x, y3: y });
          lastCpX = qx;
          lastCpY = qy;
          cx = x;
          cy = y;
        } while (isNum());
        lastCmd = "q";
        break;
      case "T":
        do {
          const qx =
            lastCmd === "Q" || lastCmd === "q" || lastCmd === "T" || lastCmd === "t"
              ? 2 * cx - lastCpX
              : cx;
          const qy =
            lastCmd === "Q" || lastCmd === "q" || lastCmd === "T" || lastCmd === "t"
              ? 2 * cy - lastCpY
              : cy;
          const x = num(),
            y = num();
          const c1x = cx + (2 / 3) * (qx - cx),
            c1y = cy + (2 / 3) * (qy - cy);
          const c2x = x + (2 / 3) * (qx - x),
            c2y = y + (2 / 3) * (qy - y);
          ops.push({ op: "curve", x1: c1x, y1: c1y, x2: c2x, y2: c2y, x3: x, y3: y });
          lastCpX = qx;
          lastCpY = qy;
          cx = x;
          cy = y;
          lastCmd = "T";
        } while (isNum());
        break;
      case "t":
        do {
          const qx =
            lastCmd === "Q" || lastCmd === "q" || lastCmd === "T" || lastCmd === "t"
              ? 2 * cx - lastCpX
              : cx;
          const qy =
            lastCmd === "Q" || lastCmd === "q" || lastCmd === "T" || lastCmd === "t"
              ? 2 * cy - lastCpY
              : cy;
          const x = cx + num(),
            y = cy + num();
          const c1x = cx + (2 / 3) * (qx - cx),
            c1y = cy + (2 / 3) * (qy - cy);
          const c2x = x + (2 / 3) * (qx - x),
            c2y = y + (2 / 3) * (qy - y);
          ops.push({ op: "curve", x1: c1x, y1: c1y, x2: c2x, y2: c2y, x3: x, y3: y });
          lastCpX = qx;
          lastCpY = qy;
          cx = x;
          cy = y;
          lastCmd = "t";
        } while (isNum());
        break;
      case "A":
      case "a": {
        const isRel = cmd === "a";
        do {
          const rx = Math.abs(num()),
            ry = Math.abs(num());
          const rotation = (num() * Math.PI) / 180;
          const largeArc = num() !== 0;
          const sweep = num() !== 0;
          const ex = isRel ? cx + num() : num();
          const ey = isRel ? cy + num() : num();
          arcToCurves(ops, cx, cy, rx, ry, rotation, largeArc, sweep, ex, ey);
          cx = ex;
          cy = ey;
        } while (isNum());
        lastCmd = cmd;
        break;
      }
      case "Z":
      case "z":
        ops.push({ op: "close" });
        cx = sx;
        cy = sy;
        lastCmd = cmd;
        break;
      default:
        // Unknown command — skip
        i++;
        break;
    }
  }

  return ops;
}

/**
 * Convert an SVG elliptical arc to cubic Bézier curves.
 * Follows the SVG spec's endpoint-to-center arc parameterization.
 * @internal
 */
function arcToCurves(
  ops: PathOp[],
  x1: number,
  y1: number,
  rx: number,
  ry: number,
  phi: number,
  largeArc: boolean,
  sweep: boolean,
  x2: number,
  y2: number
): void {
  if (rx === 0 || ry === 0) {
    ops.push({ op: "line", x: x2, y: y2 });
    return;
  }
  if (x1 === x2 && y1 === y2) {
    return;
  }

  const cosPhi = Math.cos(phi),
    sinPhi = Math.sin(phi);
  const dx = (x1 - x2) / 2,
    dy = (y1 - y2) / 2;
  const x1p = cosPhi * dx + sinPhi * dy;
  const y1p = -sinPhi * dx + cosPhi * dy;

  // Correct radii
  let rxSq = rx * rx,
    rySq = ry * ry;
  const x1pSq = x1p * x1p,
    y1pSq = y1p * y1p;
  const lambda = x1pSq / rxSq + y1pSq / rySq;
  if (lambda > 1) {
    const s = Math.sqrt(lambda);
    rx *= s;
    ry *= s;
    rxSq = rx * rx;
    rySq = ry * ry;
  }

  // Center parameterization
  let sq = (rxSq * rySq - rxSq * y1pSq - rySq * x1pSq) / (rxSq * y1pSq + rySq * x1pSq);
  if (sq < 0) {
    sq = 0;
  }
  let root = Math.sqrt(sq);
  if (largeArc === sweep) {
    root = -root;
  }
  const cxp = (root * rx * y1p) / ry;
  const cyp = (-root * ry * x1p) / rx;

  const cxr = cosPhi * cxp - sinPhi * cyp + (x1 + x2) / 2;
  const cyr = sinPhi * cxp + cosPhi * cyp + (y1 + y2) / 2;

  const angle = (ux: number, uy: number, vx: number, vy: number): number => {
    const dot = ux * vx + uy * vy;
    const len = Math.sqrt(ux * ux + uy * uy) * Math.sqrt(vx * vx + vy * vy);
    let a = Math.acos(Math.max(-1, Math.min(1, dot / len)));
    if (ux * vy - uy * vx < 0) {
      a = -a;
    }
    return a;
  };

  const theta1 = angle(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry);
  let dTheta = angle((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry);

  if (!sweep && dTheta > 0) {
    dTheta -= 2 * Math.PI;
  }
  if (sweep && dTheta < 0) {
    dTheta += 2 * Math.PI;
  }

  // Split into segments of at most π/2
  const segments = Math.ceil(Math.abs(dTheta) / (Math.PI / 2));
  const segAngle = dTheta / segments;

  for (let s = 0; s < segments; s++) {
    const t1 = theta1 + s * segAngle;
    const t2 = theta1 + (s + 1) * segAngle;
    const alpha = (4 * Math.tan((t2 - t1) / 4)) / 3;

    const cos1 = Math.cos(t1),
      sin1 = Math.sin(t1);
    const cos2 = Math.cos(t2),
      sin2 = Math.sin(t2);

    const ep1x = rx * cos1,
      ep1y = ry * sin1;
    const ep2x = rx * cos2,
      ep2y = ry * sin2;

    const cp1x = ep1x - alpha * rx * sin1;
    const cp1y = ep1y + alpha * ry * cos1;
    const cp2x = ep2x + alpha * rx * sin2;
    const cp2y = ep2y - alpha * ry * cos2;

    ops.push({
      op: "curve",
      x1: cosPhi * cp1x - sinPhi * cp1y + cxr,
      y1: sinPhi * cp1x + cosPhi * cp1y + cyr,
      x2: cosPhi * cp2x - sinPhi * cp2y + cxr,
      y2: sinPhi * cp2x + cosPhi * cp2y + cyr,
      x3: cosPhi * ep2x - sinPhi * ep2y + cxr,
      y3: sinPhi * ep2x + cosPhi * ep2y + cyr
    });
  }
}
