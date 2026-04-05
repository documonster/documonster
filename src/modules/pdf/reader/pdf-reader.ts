/**
 * PDF reader — public API.
 *
 * Provides a high-level, zero-dependency interface for reading PDF files.
 * Supports:
 * - Text extraction with multilingual support (WinAnsi, MacRoman, CJK via
 *   ToUnicode CMap, Identity-H/V, Symbol, ZapfDingbats)
 * - Image extraction (JPEG, JPEG2000, raw/Flate, CCITT, JBIG2)
 * - Annotation extraction (links, comments, highlights, stamps, etc.)
 * - Form field extraction (AcroForm: text inputs, checkboxes, radio buttons, dropdowns)
 * - Metadata reading (Info dictionary + XMP)
 * - Encrypted PDFs:
 *   - RC4 (40-bit and 128-bit) — tested via roundtrip
 *   - AES-128 (V=4, R=4) — implemented, requires external test fixtures
 *   - AES-256 (V=5, R=5) — implemented, requires external test fixtures
 * - Cross-reference tables and streams (PDF 1.5+)
 * - Incremental updates and xref recovery
 *
 * @example Basic text extraction:
 * ```typescript
 * import { readPdf } from "excelts/pdf";
 *
 * const pdf = readPdf(pdfBytes);
 * console.log(pdf.text);           // All text from all pages
 * console.log(pdf.pages[0].text);  // Text from page 1
 * ```
 *
 * @example Image extraction:
 * ```typescript
 * const pdf = readPdf(pdfBytes);
 * for (const image of pdf.pages[0].images) {
 *   console.log(image.format, image.width, image.height);
 *   fs.writeFileSync(`image.${image.format}`, image.data);
 * }
 * ```
 *
 * @example Metadata:
 * ```typescript
 * const pdf = readPdf(pdfBytes);
 * console.log(pdf.metadata.title);
 * console.log(pdf.metadata.author);
 * console.log(pdf.metadata.pageCount);
 * ```
 *
 * @example Encrypted PDF:
 * ```typescript
 * const pdf = readPdf(pdfBytes, { password: "secret" });
 * ```
 */

import { PdfDocument } from "./pdf-document";
import type { PdfDictValue } from "./pdf-parser";
import { initDecryption, isEncrypted } from "./pdf-decrypt";
import { extractTextFromPage } from "./content-interpreter";
import { reconstructText, reconstructTextLines } from "./text-reconstruction";
import type { TextLine } from "./text-reconstruction";
import type { TextFragment } from "./content-interpreter";
import { extractImagesFromPage } from "./image-extractor";
import type { ExtractedImage } from "./image-extractor";
import { extractAnnotationsFromPage } from "./annotation-extractor";
import type { PdfAnnotation } from "./annotation-extractor";
import { extractFormFields } from "./form-extractor";
import type { PdfFormField } from "./form-extractor";
import { extractMetadata } from "./metadata-reader";
import type { PdfMetadata } from "./metadata-reader";
import { PdfStructureError } from "../errors";

// =============================================================================
// Types
// =============================================================================

/**
 * Options for reading a PDF.
 */
export interface ReadPdfOptions {
  /**
   * Password for encrypted PDFs.
   * Can be either the user password or owner password.
   * @default ""
   */
  password?: string;

  /**
   * Which pages to extract (1-based).
   * If omitted, all pages are extracted.
   * @example [1, 3, 5] — extract pages 1, 3, and 5
   */
  pages?: number[];

  /**
   * Whether to extract text.
   * @default true
   */
  extractText?: boolean;

  /**
   * Whether to extract images.
   * @default true
   */
  extractImages?: boolean;

  /**
   * Whether to extract metadata.
   * @default true
   */
  extractMetadata?: boolean;

  /**
   * Whether to extract annotations (links, comments, highlights, etc.).
   * @default true
   */
  extractAnnotations?: boolean;

  /**
   * Whether to extract form fields (AcroForm: text inputs, checkboxes, dropdowns, etc.).
   * @default true
   */
  extractFormFields?: boolean;
}

/**
 * A single page from a read PDF.
 */
export interface ReadPdfPage {
  /** 1-based page number */
  pageNumber: number;
  /** Extracted text content */
  text: string;
  /** Structured text lines with position information */
  textLines: TextLine[];
  /** Raw text fragments with exact positions */
  textFragments: TextFragment[];
  /** Extracted images */
  images: ExtractedImage[];
  /** Extracted annotations (links, comments, highlights, etc.) */
  annotations: PdfAnnotation[];
  /** Page width in points */
  width: number;
  /** Page height in points */
  height: number;
  /** Warnings encountered during extraction (non-fatal errors) */
  warnings: string[];
}

/**
 * Result of reading a PDF.
 */
export interface ReadPdfResult {
  /** All text from all pages concatenated */
  text: string;
  /** Per-page results */
  pages: ReadPdfPage[];
  /** Document metadata */
  metadata: PdfMetadata;
  /** Form fields extracted from AcroForm (document-level) */
  formFields: PdfFormField[];
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Read a PDF file and extract text, images, and metadata.
 *
 * @param data - Raw PDF file bytes
 * @param options - Extraction options
 * @returns Extracted content
 * @throws {PdfStructureError} If the PDF structure is invalid
 * @throws {PdfError} If decryption fails (wrong password)
 */
export function readPdf(data: Uint8Array, options?: ReadPdfOptions): ReadPdfResult {
  const opts = {
    password: options?.password ?? "",
    pages: options?.pages,
    extractText: options?.extractText ?? true,
    extractImages: options?.extractImages ?? true,
    extractMetadata: options?.extractMetadata ?? true,
    extractAnnotations: options?.extractAnnotations ?? true,
    extractFormFields: options?.extractFormFields ?? true
  };

  // Parse document structure
  const doc = new PdfDocument(data);

  // Handle encryption
  if (isEncrypted(doc)) {
    const success = initDecryption(doc, opts.password);
    if (!success) {
      throw new PdfStructureError("Failed to decrypt PDF: incorrect password");
    }
  }

  // Extract metadata
  const metadata = opts.extractMetadata ? extractMetadata(doc) : createEmptyMetadata();

  // Get pages (with object identity for correct decryption)
  const pagesInfo = doc.getPagesWithObjInfo();
  const pageIndicesToProcess = opts.pages
    ? opts.pages.map(p => p - 1).filter(p => p >= 0 && p < pagesInfo.length)
    : Array.from({ length: pagesInfo.length }, (_, i) => i);

  // Process each page
  const pages: ReadPdfPage[] = [];

  for (const pageIdx of pageIndicesToProcess) {
    const { dict: pageDict } = pagesInfo[pageIdx];
    const pageNumber = pageIdx + 1;
    const warnings: string[] = [];

    // Extract text
    let text = "";
    let textLines: TextLine[] = [];
    let textFragments: TextFragment[] = [];

    if (opts.extractText) {
      try {
        textFragments = extractTextFromPage(pageDict, doc);
        text = reconstructText(textFragments);
        textLines = reconstructTextLines(textFragments);
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        warnings.push(`Text extraction failed on page ${pageNumber}: ${msg}`);
      }
    }

    // Extract images
    let images: ExtractedImage[] = [];
    if (opts.extractImages) {
      try {
        images = extractImagesFromPage(pageDict, doc);
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        warnings.push(`Image extraction failed on page ${pageNumber}: ${msg}`);
      }
    }

    // Extract annotations
    let annotations: PdfAnnotation[] = [];
    if (opts.extractAnnotations) {
      try {
        annotations = extractAnnotationsFromPage(pageDict, doc);
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        warnings.push(`Annotation extraction failed on page ${pageNumber}: ${msg}`);
      }
    }

    // Get page dimensions
    const { width, height } = getPageDimensions(pageDict, doc);

    pages.push({
      pageNumber,
      text,
      textLines,
      textFragments,
      images,
      annotations,
      width,
      height,
      warnings
    });
  }

  // Concatenate all page text
  const allText = pages.map(p => p.text).join("\n\n");

  // Update page count in metadata
  if (opts.extractMetadata) {
    metadata.pageCount = pagesInfo.length;
  }

  // Extract form fields (document-level, not per-page)
  let formFields: PdfFormField[] = [];
  if (opts.extractFormFields) {
    try {
      formFields = extractFormFields(doc);
    } catch {
      // Non-fatal — just return empty
    }
  }

  return {
    text: allText,
    pages,
    metadata,
    formFields
  };
}

// =============================================================================
// Helpers
// =============================================================================

function getPageDimensions(
  pageDict: PdfDictValue,
  doc: PdfDocument
): { width: number; height: number } {
  return doc.resolvePageBox(pageDict) ?? { width: 612, height: 792 }; // Default: US Letter
}

function createEmptyMetadata(): PdfMetadata {
  return {
    title: "",
    author: "",
    subject: "",
    keywords: "",
    creator: "",
    producer: "",
    creationDate: null,
    modDate: null,
    pdfVersion: "",
    pageCount: 0,
    encrypted: false,
    pageSize: null,
    xmpXml: null,
    custom: {}
  };
}
