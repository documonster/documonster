/**
 * Low-level PDF 1.4 document writer.
 *
 * Generates a valid PDF binary from structured objects.
 * Supports: text, graphics, images (JPEG/PNG), fonts (14 standard).
 *
 * Zero external dependencies – uses only TextEncoder and basic math.
 */

// =============================================================================
// Types
// =============================================================================

export interface PdfRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface PdfColor {
  r: number;
  g: number;
  b: number;
}

export interface PdfTextOp {
  kind: "text";
  x: number;
  y: number;
  text: string;
  fontRef: string;
  fontSize: number;
  color: PdfColor;
  underline?: boolean;
  strike?: boolean;
  rotation?: number;
}

export interface PdfRectOp {
  kind: "rect";
  x: number;
  y: number;
  width: number;
  height: number;
  fillColor?: PdfColor;
  strokeColor?: PdfColor;
  lineWidth?: number;
}

export interface PdfLineOp {
  kind: "line";
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  color: PdfColor;
  lineWidth: number;
  dash?: number[];
}

export interface PdfImageOp {
  kind: "image";
  imageKey: string;
  x: number;
  y: number;
  width: number;
  height: number;
}

export type PdfOp = PdfTextOp | PdfRectOp | PdfLineOp | PdfImageOp;

export interface PdfPageDef {
  width: number;
  height: number;
  ops: PdfOp[];
}

export interface PdfImageData {
  key: string;
  data: Uint8Array;
  width: number;
  height: number;
  format: "jpeg" | "png";
}

// =============================================================================
// Internal helpers
// =============================================================================

const encoder = new TextEncoder();

function encodeStr(s: string): Uint8Array {
  return encoder.encode(s);
}

/**
 * Escape a string for PDF text objects.
 * Replaces \, (, ) with backslash-escaped versions.
 */
function pdfEscapeString(s: string): string {
  return s
    .replaceAll("\\", String.raw`\\`)
    .replaceAll("(", String.raw`\(`)
    .replaceAll(")", String.raw`\)`);
}

/**
 * Format a number for PDF: max 4 decimal places, no trailing zeros.
 */
function n(v: number): string {
  return Number(v.toFixed(4)).toString();
}

// =============================================================================
// PNG Decoder (minimal – extracts raw pixel data for PDF)
// =============================================================================

function decodePng(data: Uint8Array): {
  width: number;
  height: number;
  colorType: number;
  bitDepth: number;
  rawPixels: Uint8Array;
  hasAlpha: boolean;
} {
  // Check PNG signature
  const sig = [137, 80, 78, 71, 13, 10, 26, 10];
  for (let i = 0; i < 8; i++) {
    if (data[i] !== sig[i]) {
      throw new Error("Invalid PNG signature");
    }
  }

  let pos = 8;
  let width = 0;
  let height = 0;
  let bitDepth = 0;
  let colorType = 0;
  const idatChunks: Uint8Array[] = [];

  while (pos < data.length) {
    const length = readUint32(data, pos);
    const type = String.fromCodePoint(data[pos + 4], data[pos + 5], data[pos + 6], data[pos + 7]);

    if (type === "IHDR") {
      width = readUint32(data, pos + 8);
      height = readUint32(data, pos + 12);
      bitDepth = data[pos + 16];
      colorType = data[pos + 17];
    } else if (type === "IDAT") {
      idatChunks.push(data.slice(pos + 8, pos + 8 + length));
    } else if (type === "IEND") {
      break;
    }

    pos += 12 + length;
  }

  // Concatenate IDAT chunks
  const totalLen = idatChunks.reduce((s, c) => s + c.length, 0);
  const compressed = new Uint8Array(totalLen);
  let offset = 0;
  for (const chunk of idatChunks) {
    compressed.set(chunk, offset);
    offset += chunk.length;
  }

  // Decompress using DecompressionStream (available in modern runtimes)
  // For PDF embedding, we store the raw compressed data and use FlateDecode
  // since PNG IDAT is zlib-compressed which is what PDF FlateDecode expects

  const hasAlpha = colorType === 4 || colorType === 6;

  return {
    width,
    height,
    colorType,
    bitDepth,
    rawPixels: compressed, // Keep compressed for FlateDecode
    hasAlpha
  };
}

function readUint32(data: Uint8Array, offset: number): number {
  return (
    ((data[offset] << 24) |
      (data[offset + 1] << 16) |
      (data[offset + 2] << 8) |
      data[offset + 3]) >>>
    0
  );
}

// =============================================================================
// PDF Writer
// =============================================================================

export class PdfWriter {
  private objects: Uint8Array[] = [];
  private nextObjId = 1;
  private fontObjects: Map<string, number> = new Map();
  private imageObjects: Map<string, { objId: number; width: number; height: number }> = new Map();

  // ==========================================================================
  // Object management
  // ==========================================================================

  private allocObj(): number {
    return this.nextObjId++;
  }

  private writeObj(objId: number, content: string): void {
    const data = encodeStr(`${objId} 0 obj\n${content}\nendobj\n`);
    this.objects[objId] = data;
  }

  private writeStreamObj(objId: number, dict: string, streamData: Uint8Array): void {
    const header = encodeStr(`${objId} 0 obj\n${dict}\nstream\n`);
    const footer = encodeStr("\nendstream\nendobj\n");
    const combined = new Uint8Array(header.length + streamData.length + footer.length);
    combined.set(header, 0);
    combined.set(streamData, header.length);
    combined.set(footer, header.length + streamData.length);
    this.objects[objId] = combined;
  }

  // ==========================================================================
  // Font registration
  // ==========================================================================

  registerFont(pdfFontName: string): string {
    const ref = "F" + this.fontObjects.size;
    if (this.fontObjects.has(pdfFontName)) {
      return this.getFontRef(pdfFontName);
    }
    const objId = this.allocObj();
    this.fontObjects.set(pdfFontName, objId);

    let encoding = "/Encoding /WinAnsiEncoding";
    if (pdfFontName === "Symbol" || pdfFontName === "ZapfDingbats") {
      encoding = "";
    }

    this.writeObj(objId, `<< /Type /Font /Subtype /Type1 /BaseFont /${pdfFontName} ${encoding} >>`);
    return ref;
  }

  private getFontRef(pdfFontName: string): string {
    const keys = Array.from(this.fontObjects.keys());
    const idx = keys.indexOf(pdfFontName);
    return "F" + idx;
  }

  getFontRefName(pdfFontName: string): string {
    if (!this.fontObjects.has(pdfFontName)) {
      this.registerFont(pdfFontName);
    }
    return this.getFontRef(pdfFontName);
  }

  // ==========================================================================
  // Image registration
  // ==========================================================================

  registerImage(img: PdfImageData): void {
    if (this.imageObjects.has(img.key)) {
      return;
    }

    const objId = this.allocObj();

    if (img.format === "jpeg") {
      this.writeStreamObj(
        objId,
        `<< /Type /XObject /Subtype /Image /Width ${img.width} /Height ${img.height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${img.data.length} >>`,
        img.data
      );
    } else {
      // PNG: Use the raw IDAT data with FlateDecode
      // We strip the PNG header and recombine the IDAT chunks
      const png = decodePng(img.data);

      const bpc = png.bitDepth;
      let colorSpace = "/DeviceRGB";
      let channels = 3;

      // Grayscale (type 0) and Gray+Alpha (type 4) → embed as gray
      if (png.colorType === 0 || png.colorType === 4) {
        colorSpace = "/DeviceGray";
        channels = 1;
      }
      // All other types (2=RGB, 3=indexed, 6=RGBA) → embed as RGB (default)

      // Use /DecodeParms with PNG predictor so PDF can decode the filtered rows
      const decodeParms = `/DecodeParms << /Predictor 15 /Colors ${channels} /BitsPerComponent ${bpc} /Columns ${png.width} >>`;

      this.writeStreamObj(
        objId,
        `<< /Type /XObject /Subtype /Image /Width ${png.width} /Height ${png.height} /ColorSpace ${colorSpace} /BitsPerComponent ${bpc} /Filter /FlateDecode ${decodeParms} /Length ${png.rawPixels.length} >>`,
        png.rawPixels
      );

      // Store actual PNG dimensions
      this.imageObjects.set(img.key, { objId, width: png.width, height: png.height });
      return;
    }

    this.imageObjects.set(img.key, { objId, width: img.width, height: img.height });
  }

  // ==========================================================================
  // Page content generation
  // ==========================================================================

  private generatePageContent(page: PdfPageDef): Uint8Array {
    const lines: string[] = [];

    for (const op of page.ops) {
      switch (op.kind) {
        case "rect": {
          if (op.fillColor) {
            lines.push(
              `${n(op.fillColor.r)} ${n(op.fillColor.g)} ${n(op.fillColor.b)} rg`,
              `${n(op.x)} ${n(op.y)} ${n(op.width)} ${n(op.height)} re f`
            );
          }
          if (op.strokeColor) {
            const lw = op.lineWidth ?? 0.5;
            lines.push(
              `${n(lw)} w`,
              `${n(op.strokeColor.r)} ${n(op.strokeColor.g)} ${n(op.strokeColor.b)} RG`,
              `${n(op.x)} ${n(op.y)} ${n(op.width)} ${n(op.height)} re S`
            );
          }
          break;
        }
        case "line": {
          const lw = op.lineWidth;
          lines.push(`${n(lw)} w`, `${n(op.color.r)} ${n(op.color.g)} ${n(op.color.b)} RG`);
          if (op.dash && op.dash.length > 0) {
            lines.push(`[${op.dash.map(n).join(" ")}] 0 d`);
          } else {
            lines.push("[] 0 d");
          }
          lines.push(`${n(op.x1)} ${n(op.y1)} m ${n(op.x2)} ${n(op.y2)} l S`);
          break;
        }
        case "text": {
          const fontRef = this.getFontRefName(op.fontRef);
          lines.push(
            "BT",
            `${n(op.color.r)} ${n(op.color.g)} ${n(op.color.b)} rg`,
            `/${fontRef} ${n(op.fontSize)} Tf`
          );
          if (op.rotation) {
            const rad = (op.rotation * Math.PI) / 180;
            const cos = Math.cos(rad);
            const sin = Math.sin(rad);
            lines.push(`${n(cos)} ${n(sin)} ${n(-sin)} ${n(cos)} ${n(op.x)} ${n(op.y)} Tm`);
          } else {
            lines.push(`${n(op.x)} ${n(op.y)} Td`);
          }
          lines.push(`(${pdfEscapeString(op.text)}) Tj`, "ET");

          // Underline: draw a line under the text
          if (op.underline) {
            const thickness = op.fontSize * 0.05;
            const yLine = op.y - op.fontSize * 0.15;
            lines.push(
              `${n(thickness)} w`,
              `${n(op.color.r)} ${n(op.color.g)} ${n(op.color.b)} RG`,
              `${n(op.x)} ${n(yLine)} m ${n(op.x + op.fontSize * op.text.length * 0.5)} ${n(yLine)} l S`
            );
          }

          // Strikethrough: draw a line through the middle of the text
          if (op.strike) {
            const thickness = op.fontSize * 0.05;
            const yLine = op.y + op.fontSize * 0.25;
            lines.push(
              `${n(thickness)} w`,
              `${n(op.color.r)} ${n(op.color.g)} ${n(op.color.b)} RG`,
              `${n(op.x)} ${n(yLine)} m ${n(op.x + op.fontSize * op.text.length * 0.5)} ${n(yLine)} l S`
            );
          }
          break;
        }
        case "image": {
          const imgInfo = this.imageObjects.get(op.imageKey);
          if (imgInfo) {
            lines.push(
              "q",
              `${n(op.width)} 0 0 ${n(op.height)} ${n(op.x)} ${n(op.y)} cm`,
              `/Im${op.imageKey} Do`,
              "Q"
            );
          }
          break;
        }
      }
    }

    return encodeStr(lines.join("\n"));
  }

  // ==========================================================================
  // Build the full PDF
  // ==========================================================================

  build(pages: PdfPageDef[], images: PdfImageData[]): Uint8Array {
    // Reset
    this.objects = [];
    this.nextObjId = 1;
    this.fontObjects = new Map();
    this.imageObjects = new Map();

    // Register images
    for (const img of images) {
      this.registerImage(img);
    }

    // Pre-register all standard PDF fonts. The layout engine references them
    // by base font name (e.g. "Helvetica"), so we register the full set upfront.
    const standardFonts = [
      "Helvetica",
      "Helvetica-Bold",
      "Helvetica-Oblique",
      "Helvetica-BoldOblique",
      "Times-Roman",
      "Times-Bold",
      "Times-Italic",
      "Times-BoldItalic",
      "Courier",
      "Courier-Bold",
      "Courier-Oblique",
      "Courier-BoldOblique"
    ];
    for (const f of standardFonts) {
      this.registerFont(f);
    }

    // Allocate catalog and pages objects
    const catalogId = this.allocObj();
    const pagesId = this.allocObj();

    // Build each page
    const pageObjIds: number[] = [];
    for (const pageDef of pages) {
      const contentData = this.generatePageContent(pageDef);
      const contentId = this.allocObj();
      this.writeStreamObj(contentId, `<< /Length ${contentData.length} >>`, contentData);

      // Build font dictionary for this page
      const fontEntries = Array.from(this.fontObjects.entries())
        .map(([name, objId], idx) => `/F${idx} ${objId} 0 R`)
        .join(" ");

      // Build image dictionary for this page
      const imageEntries = Array.from(this.imageObjects.entries())
        .map(([key, info]) => `/Im${key} ${info.objId} 0 R`)
        .join(" ");

      const xObjectDict = imageEntries ? `/XObject << ${imageEntries} >>` : "";
      const resourceDict = `<< /Font << ${fontEntries} >> ${xObjectDict} >>`;

      const pageId = this.allocObj();
      this.writeObj(
        pageId,
        `<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 ${n(pageDef.width)} ${n(pageDef.height)}] /Contents ${contentId} 0 R /Resources ${resourceDict} >>`
      );
      pageObjIds.push(pageId);
    }

    // Pages object
    const kidsStr = pageObjIds.map(id => `${id} 0 R`).join(" ");
    this.writeObj(pagesId, `<< /Type /Pages /Kids [${kidsStr}] /Count ${pageObjIds.length} >>`);

    // Catalog
    this.writeObj(catalogId, `<< /Type /Catalog /Pages ${pagesId} 0 R >>`);

    // Assemble the PDF
    return this.assemble(catalogId);
  }

  private assemble(catalogId: number): Uint8Array {
    const parts: Uint8Array[] = [];
    let currentOffset = 0;

    // Header
    const header = encodeStr("%PDF-1.4\n%\xC0\xC1\xC2\xC3\n");
    parts.push(header);
    currentOffset += header.length;

    // Object offsets
    const objOffsets: number[] = new Array(this.nextObjId).fill(0);

    // Write each object
    for (let i = 1; i < this.nextObjId; i++) {
      const data = this.objects[i];
      if (data) {
        objOffsets[i] = currentOffset;
        parts.push(data);
        currentOffset += data.length;
      }
    }

    // Cross-reference table
    const xrefOffset = currentOffset;
    const xrefLines: string[] = [];
    xrefLines.push("xref", `0 ${this.nextObjId}`, "0000000000 65535 f \n");
    for (let i = 1; i < this.nextObjId; i++) {
      const off = objOffsets[i].toString().padStart(10, "0");
      xrefLines.push(`${off} 00000 n \n`);
    }

    const xrefData = encodeStr(xrefLines.join(""));
    parts.push(xrefData);
    currentOffset += xrefData.length;

    // Trailer
    const trailer = encodeStr(
      `trailer\n<< /Size ${this.nextObjId} /Root ${catalogId} 0 R >>\nstartxref\n${xrefOffset}\n%%EOF\n`
    );
    parts.push(trailer);
    currentOffset += trailer.length;

    // Combine all parts
    const result = new Uint8Array(currentOffset);
    let offset = 0;
    for (const part of parts) {
      result.set(part, offset);
      offset += part.length;
    }

    return result;
  }
}
