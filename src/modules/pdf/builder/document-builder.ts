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

    writer.addCatalog(pagesTreeObjNum, {
      outlinesRef,
      extraEntries: catalogExtras.length > 0 ? catalogExtras : undefined
    });

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
}
