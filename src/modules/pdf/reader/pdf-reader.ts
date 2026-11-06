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
 * @example Text extraction:
 * ```typescript
 * import { Pdf } from "documonster/pdf";
 *
 * const doc = await Pdf.extract(pdfBytes);
 * console.log(doc.text);           // All text from all pages
 * console.log(doc.pages[0].text);  // Text from page 1
 * ```
 *
 * @example Image extraction:
 * ```typescript
 * const doc = await Pdf.extract(pdfBytes);
 * for (const image of doc.pages[0].images) {
 *   console.log(image.format, image.width, image.height);
 *   fs.writeFileSync(`image.${image.format}`, image.data);
 * }
 * ```
 *
 * @example Metadata:
 * ```typescript
 * const doc = await Pdf.extract(pdfBytes);
 * console.log(doc.metadata.title);
 * console.log(doc.metadata.author);
 * console.log(doc.metadata.pageCount);
 * ```
 *
 * @example Encrypted PDF:
 * ```typescript
 * const doc = await Pdf.extract(pdfBytes, { password: "secret" });
 * ```
 */

import { at, isBlob, pdfBytesOrBlob } from "@pdf/core/arguments";
import { PdfPasswordError } from "@pdf/errors";
import { extractAnnotationsFromPage } from "@pdf/reader/annotation-extractor";
import type { PdfAnnotation } from "@pdf/reader/annotation-extractor";
import { extractBookmarks } from "@pdf/reader/bookmark-extractor";
import type { PdfBookmark } from "@pdf/reader/bookmark-extractor";
import { extractTextFromPage } from "@pdf/reader/content-interpreter";
import type { TextFragment } from "@pdf/reader/content-interpreter";
import { extractFormFields } from "@pdf/reader/form-extractor";
import type { PdfFormField } from "@pdf/reader/form-extractor";
import { extractImagesFromPage } from "@pdf/reader/image-extractor";
import type { ExtractedImage } from "@pdf/reader/image-extractor";
import { extractMetadata } from "@pdf/reader/metadata-reader";
import type { PdfMetadata } from "@pdf/reader/metadata-reader";
import { initDecryption, isEncrypted } from "@pdf/reader/pdf-decrypt";
import { PdfDocument } from "@pdf/reader/pdf-document";
import type { PdfDictValue } from "@pdf/reader/pdf-parser";
import { extractTables } from "@pdf/reader/table-extractor";
import type { PdfTable } from "@pdf/reader/table-extractor";
import { reconstructText, reconstructTextLines } from "@pdf/reader/text-reconstruction";
import type { TextLine } from "@pdf/reader/text-reconstruction";
import { check, s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";
import { createYieldBudget } from "@utils/yield";

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
  pages?: readonly number[];

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
   * Whether to extract metadata. When `false`, `metadata` still reports `pageCount` and
   * `encrypted`, which cost nothing to know; every other field is empty (`""`, `null`, `{}`).
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

  /**
   * Whether to extract bookmarks (document outline / table of contents).
   * @default true
   */
  extractBookmarks?: boolean;

  /**
   * Whether to extract tables from pages using text positioning heuristics.
   * Opt-in since table detection is heavier than plain text extraction.
   * @default false
   */
  extractTables?: boolean;

  /**
   * Largest output, in bytes, a single stream filter may produce before the
   * read is aborted with a `PdfLimitExceededError` (`limit: "maxDecodedBytes"`). Guards
   * against decompression bombs. A number `>= 0`, or `Infinity` for no limit; anything else is
   * refused with `PdfArgumentError` (field `options.maxDecodedBytes`).
   * @default 268435456 (256 MiB)
   */
  maxDecodedBytes?: number;
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
  /** Tables detected from text fragment positioning (opt-in via extractTables) */
  tables: PdfTable[];
  /** Page width in points */
  width: number;
  /** Page height in points */
  height: number;
  /**
   * The page's `/Rotate` — how a viewer turns it for display, clockwise, in degrees — inherited
   * from the page tree as the format specifies. `width` and `height` are of the page as it is
   * stored, before this rotation. A value that is not a multiple of 90 reads as 0.
   */
  rotation: 0 | 90 | 180 | 270;
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
  /** Bookmarks (document outline) extracted from the outline tree */
  bookmarks: PdfBookmark[];
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Read a PDF file and extract text, images, and metadata.
 * Yields to the event loop between pages, whenever about 10 ms of work have passed since the
 * last yield, so a long document does not block it.
 *
 * Pages are named by 1-based number in `options.pages`; a number past the last page is
 * skipped (so the result may hold fewer pages than were asked for — each page carries its
 * `pageNumber`).
 *
 * @param data - Raw PDF file bytes — a `Uint8Array`, another `ArrayBufferView` (a `Buffer`, a
 *   `DataView`) or an `ArrayBuffer`, viewed rather than copied — or a `Blob` (a browser
 *   `File`), whose bytes are read first
 * @param options - Extraction options
 * @returns Promise of extracted content
 * @throws {PdfArgumentError} `data` is none of these, or an option is malformed
 * @throws {PdfStructureError} If the PDF structure is invalid
 * @throws {PdfError} If decryption fails (wrong password)
 */
export async function readPdf(
  data: Uint8Array | ArrayBuffer | ArrayBufferView | Blob,
  options?: ReadPdfOptions
): Promise<ReadPdfResult> {
  // Both checked before a Blob is read, so a bad argument costs no read.
  const source = pdfBytesOrBlob("Pdf.extract", "data", data);
  checkReadPdfArguments(options);
  const bytes = isBlob(source) ? new Uint8Array(await source.arrayBuffer()) : source;
  const { doc, opts, metadata, pagesInfo, pageIndicesToProcess } = prepareRead(bytes, options);

  const pages: ReadPdfPage[] = [];
  const yieldIfDue = createYieldBudget();
  for (let i = 0; i < pageIndicesToProcess.length; i++) {
    const pageIdx = pageIndicesToProcess[i];
    pages.push(processPage(pagesInfo[pageIdx].dict, pageIdx, doc, opts));
    // Page extraction tolerates malformed content, so a limit hit inside it is
    // recorded rather than thrown; stop before starting the next page.
    doc.throwIfLimitExceeded();
    if (i < pageIndicesToProcess.length - 1) {
      await yieldIfDue();
    }
  }

  return finalizeRead(pages, pagesInfo.length, metadata, opts, doc);
}

// =============================================================================
// Internal — Shared Pipeline
// =============================================================================

interface ResolvedReadOptions {
  password: string;
  pages: readonly number[] | undefined;
  extractText: boolean;
  extractImages: boolean;
  extractMetadata: boolean;
  extractAnnotations: boolean;
  extractFormFields: boolean;
  extractBookmarks: boolean;
  extractTables: boolean;
}

interface PreparedRead {
  doc: PdfDocument;
  opts: ResolvedReadOptions;
  metadata: PdfMetadata;
  pagesInfo: Array<{ dict: PdfDictValue; objNum: number; gen: number }>;
  pageIndicesToProcess: number[];
}

/**
 * Shared setup: parse document, handle encryption, extract metadata, resolve pages.
 */
function prepareRead(data: Uint8Array, options?: ReadPdfOptions): PreparedRead {
  const opts: ResolvedReadOptions = {
    password: options?.password ?? "",
    pages: options?.pages,
    extractText: options?.extractText ?? true,
    extractImages: options?.extractImages ?? true,
    extractMetadata: options?.extractMetadata ?? true,
    extractAnnotations: options?.extractAnnotations ?? true,
    extractFormFields: options?.extractFormFields ?? true,
    extractBookmarks: options?.extractBookmarks ?? true,
    extractTables: options?.extractTables ?? false
  };

  const doc = new PdfDocument(data, options?.maxDecodedBytes);

  if (isEncrypted(doc)) {
    const success = initDecryption(doc, opts.password);
    if (!success) {
      throw new PdfPasswordError("Failed to decrypt PDF: incorrect password");
    }
  }

  const metadata = opts.extractMetadata ? extractMetadata(doc) : createEmptyMetadata();
  const pagesInfo = doc.getPagesWithObjInfo();
  const pageIndicesToProcess = opts.pages
    ? opts.pages.map(p => p - 1).filter(p => p >= 0 && p < pagesInfo.length)
    : Array.from({ length: pagesInfo.length }, (_, i) => i);
  doc.throwIfLimitExceeded();

  return { doc, opts, metadata, pagesInfo, pageIndicesToProcess };
}

/**
 * Check the options `Pdf.extract` was given (its bytes are checked by `pdfBytesOrBlob`: a string
 * used to be reported as a corrupt file, "no objects found"). A `pages` entry
 * that is not a whole number crashed with a `TypeError`; and `pages: [0]`, the first page
 * counted from 0, quietly extracted nothing.
 * `maxDecodedBytes` threw a `RangeError` from `PdfDocument`, naming no field.
 */
const flag = () => s.boolean().optional();

const readPdfOptionsSchema = s.object(
  {
    password: s.string().optional(),
    pages: s
      .readonlyArray(s.number({ integer: true, min: 1 }).orUndefined(), {
        expected: "an array of 1-based page numbers"
      })
      .optional(),
    extractText: flag(),
    extractImages: flag(),
    extractMetadata: flag(),
    extractAnnotations: flag(),
    extractFormFields: flag(),
    extractBookmarks: flag(),
    extractTables: flag(),
    maxDecodedBytes: s.number({ min: 0, expected: "a number >= 0 or Infinity" }).optional()
  },
  { expected: "an options object" }
);
const _readPdfOptionsMatch: SchemaMatches<typeof readPdfOptionsSchema, ReadPdfOptions> = true;
void _readPdfOptionsMatch;

function checkReadPdfArguments(options: ReadPdfOptions | undefined): void {
  check(readPdfOptionsSchema.optional(), options, at("Pdf.extract", "options"));
}

/**
 * Process a single page: extract text, images, annotations, and dimensions.
 */
function processPage(
  pageDict: PdfDictValue,
  pageIdx: number,
  doc: PdfDocument,
  opts: ResolvedReadOptions
): ReadPdfPage {
  const pageNumber = pageIdx + 1;
  const warnings: string[] = [];

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

  let images: ExtractedImage[] = [];
  if (opts.extractImages) {
    try {
      images = extractImagesFromPage(pageDict, doc);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      warnings.push(`Image extraction failed on page ${pageNumber}: ${msg}`);
    }
  }

  let annotations: PdfAnnotation[] = [];
  if (opts.extractAnnotations) {
    try {
      annotations = extractAnnotationsFromPage(pageDict, doc);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      warnings.push(`Annotation extraction failed on page ${pageNumber}: ${msg}`);
    }
  }

  const { width, height } = getPageDimensions(pageDict, doc);

  let tables: PdfTable[] = [];
  if (opts.extractTables) {
    try {
      tables = extractTables(textFragments, width, height);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      warnings.push(`Table extraction failed on page ${pageNumber}: ${msg}`);
    }
  }

  return {
    pageNumber,
    text,
    textLines,
    textFragments,
    images,
    annotations,
    tables,
    width,
    height,
    rotation: getPageRotation(pageDict, doc),
    warnings
  };
}

/** `/Rotate`, inherited from the page tree (ISO 32000-2 §7.7.3.4), normalised to 0–270. */
function getPageRotation(pageDict: PdfDictValue, doc: PdfDocument): 0 | 90 | 180 | 270 {
  const seen = new Set<PdfDictValue>();
  let current: PdfDictValue | null = pageDict;
  while (current && !seen.has(current)) {
    seen.add(current);
    const raw = doc.deref(current.get("Rotate"));
    if (typeof raw === "number") {
      if (!Number.isInteger(raw) || raw % 90 !== 0) {
        return 0;
      }
      return (((raw % 360) + 360) % 360) as 0 | 90 | 180 | 270;
    }
    current = doc.derefDict(current.get("Parent"));
  }
  return 0;
}

/**
 * Finalize: concatenate text, update metadata page count, extract form fields.
 */
function finalizeRead(
  pages: ReadPdfPage[],
  totalPageCount: number,
  metadata: PdfMetadata,
  opts: ResolvedReadOptions,
  doc: PdfDocument
): ReadPdfResult {
  const allText = pages.map(p => p.text).join("\n\n");

  // The page count and the encryption are known whether or not the metadata was read: a
  // record reporting a two-page file as `pageCount: 0` was a wrong answer, not an absent one.
  metadata.pageCount = totalPageCount;
  if (!opts.extractMetadata) {
    metadata.encrypted = isEncrypted(doc);
  }

  let formFields: PdfFormField[] = [];
  if (opts.extractFormFields) {
    try {
      formFields = extractFormFields(doc);
    } catch {
      // Non-fatal — just return empty
    }
  }

  let bookmarks: PdfBookmark[] = [];
  if (opts.extractBookmarks) {
    try {
      bookmarks = extractBookmarks(doc);
    } catch {
      // Non-fatal — just return empty
    }
  }

  // Extraction tolerates malformed objects, so a resource limit hit anywhere
  // above was recorded rather than thrown; refuse the truncated result here.
  doc.throwIfLimitExceeded();

  return { text: allText, pages, metadata, formFields, bookmarks };
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
    pdfVersion: "",
    pageCount: 0,
    encrypted: false,
    custom: {}
  };
}
