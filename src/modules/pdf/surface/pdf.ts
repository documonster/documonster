/**
 * `Pdf` namespace surface — PDF writing, reading, building, editing,
 * conversion, and digital signatures.
 *
 * `import { Pdf } from "documonster/pdf"` →
 *   `Pdf.fromTable(rows)`, `Pdf.extract(bytes)`, `new Pdf.Builder()`,
 *   `Pdf.Editor.read(bytes)`, `await Pdf.verifySignatures(bytes)`, `Pdf.Signature.*`,
 *   `await Pdf.fromExcel(wb)`, `await Pdf.fromDocx(doc)`, `Pdf.createDrawSurface(page, rect)`.
 *
 * `Pdf.extract` pulls content out of a file (text, images, fields, tables) and returns
 * plain data; `Pdf.Editor.read` opens a file as an editable document. The two used to
 * share the verb `read`, which promised the same kind of result from both.
 *
 * The cross-module converters (`fromExcel`/`fromDocx`/`fromChart`) are **lazily loaded** via dynamic `import()`: they
 * statically pull the excel / word object models, so wiring them as plain
 * re-exports would force every `Pdf.fromTable` consumer to bundle excel+word
 * (~hundreds of KB). The dynamic boundary keeps the core PDF engine
 * tree-shakeable — a consumer who never calls a converter never bundles the
 * source module. (Verified by scripts/treeshake-verify.ts.)
 */
import type { DrawSurface } from "@draw/surface";
import {
  PdfDocumentBuilder,
  PdfPageBuilder,
  parseSvgPath as parseSvgPathImpl
} from "@pdf/builder/document-builder";
import type { PathOp } from "@pdf/builder/document-builder";
import {
  at,
  checkConverterOptions,
  checkDrawPage,
  checkMetadataText,
  checkPdfExportOptions,
  checkPdfInput,
  checkString,
  documentOptionsSchemaWith,
  documentOptionsSchemaWithReadonly
} from "@pdf/core/arguments";
import { signPdf, verifyPdfSignature } from "@pdf/core/digital-signature";
import type { SignPdfOptions } from "@pdf/core/digital-signature";
import { PdfArgumentError } from "@pdf/errors";
import type {
  ChartHandle,
  ChartToPdfOptions,
  ExcelToPdfOptions,
  Workbook
} from "@pdf/excel-bridge";
import { pdf } from "@pdf/pdf";
import type { PdfInput } from "@pdf/pdf";
import { readPdf } from "@pdf/reader/pdf-reader";
import { createPdfDrawSurface } from "@pdf/render/draw-surface";
import type { DrawSurfaceRect, PdfDrawPage, PdfDrawSurfaceOptions } from "@pdf/render/draw-surface";
import * as SignatureNs from "@pdf/surface/signature";
import type { PdfExportOptions } from "@pdf/types";
import type { DocxDocument, DocxToPdfOptions } from "@pdf/word-bridge";
import type { WordChart } from "@pdf/word-chart-bridge";
import { isRecordArgument } from "@utils/object";
import { check, fail, s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";

// --- Writing (core engine, statically linked) ---

/**
 * Lay out tabular data as a PDF table: a 2D array of rows, a sheet (rows with a name,
 * column widths and images) or a book of sheets, each its own section. The plain-data
 * counterpart of {@link fromExcel}.
 *
 * Yields to the event loop between output pages during layout and rendering.
 */
export async function fromTable(input: PdfInput, options?: PdfExportOptions): Promise<Uint8Array> {
  checkPdfInput("Pdf.fromTable", input);
  checkPdfExportOptions("Pdf.fromTable", options);
  return pdf(input, options);
}

/**
 * @deprecated Use {@link fromTable}. `create` is the name every other namespace uses for
 * a handle constructor, and this returns bytes. Removed in 1.0.
 */
export const create = fromTable;

// --- Extracting ---

/**
 * Extract the content of a PDF — text, pages, images, annotations, form fields,
 * bookmarks, tables and metadata — as plain data. To change a file, open it with
 * `Pdf.Editor.read` instead.
 */
export const extract = readPdf;

/**
 * @deprecated Use {@link extract}. `read` is the verb that opens a document as its
 * editable object (`Pdf.Editor.read`); this extracts content. Removed in 1.0.
 */
export const read = readPdf;

// --- Building (imported above; re-exported under namespace-friendly names) ---
export { PdfDocumentBuilder as Builder, PdfPageBuilder as PageBuilder };

/**
 * Parse an SVG path `d` attribute into path operations (M, L, H, V, C, S, Q, T, A, Z, absolute
 * and relative), for `drawPath`.
 *
 * @throws {PdfArgumentError} `d` is not a string.
 */
export function parseSvgPath(d: string): PathOp[] {
  checkString("Pdf.parseSvgPath", "d", d);
  return parseSvgPathImpl(d);
}

// --- Editing ---
export { PdfEditor as Editor, PdfEditorPage as EditorPage } from "@pdf/builder/pdf-editor";

// --- Digital signatures ---
export { verifyPdfSignatures as verifySignatures } from "@pdf/reader/signature-verifier";

/**
 * @deprecated Use `Pdf.Signature.verify`, or `Pdf.verifySignatures` to find and check
 * every signature in a file. Removed in 1.0.
 */
export const verifySignature = verifyPdfSignature;

/**
 * @deprecated Use `Pdf.Signature.fillPlaceholder(bytes, { certificate, privateKey })` —
 * this only completes a placeholder already in the file. Removed in 1.0.
 */
export function sign(pdfBytes: Uint8Array, options: SignPdfOptions): Promise<Uint8Array>;
/**
 * @deprecated Use `Pdf.Signature.fillPlaceholder(bytes, { certificate, privateKey })`.
 * Removed in 1.0.
 */
export function sign(
  pdfBytes: Uint8Array,
  certificate: Uint8Array,
  privateKey: Uint8Array
): Promise<Uint8Array>;
export function sign(
  pdfBytes: Uint8Array,
  optionsOrCertificate: SignPdfOptions | Uint8Array,
  privateKey?: Uint8Array
): Promise<Uint8Array> {
  // Anything but an options object is the positional certificate, handed on unchanged as the
  // positional form always did — a typed array from another realm fails `instanceof Uint8Array`.
  if (!isRecordArgument(optionsOrCertificate)) {
    return signPdf(pdfBytes, optionsOrCertificate, privateKey as Uint8Array);
  }
  return signPdf(pdfBytes, optionsOrCertificate.certificate, optionsOrCertificate.privateKey);
}

/** Low-level signing and verification primitives — see `Pdf.Signature`. */
export * as Signature from "@pdf/surface/signature";

/**
 * @deprecated Use `Pdf.Signature.buildDictPlaceholder`. Removed in 1.0.
 */
export const buildSignatureDictPlaceholder = SignatureNs.buildDictPlaceholder;
/** @deprecated Use `Pdf.Signature.parseAsn1`. Removed in 1.0. */
export const asn1Parse = SignatureNs.parseAsn1;

// --- Drawing (`documonster/draw` backend) ---

/**
 * The PDF backend for `documonster/draw`: a surface that draws a display list onto
 * `page`, its top-left corner at `rect`'s top-left (PDF points, `y` is the bottom edge) and
 * its coordinates multiplied by `options.scale`. The drawing is not clipped or fitted to
 * `rect`; size the rect as the list's size times the scale.
 *
 * A display list built anywhere — a chart, a diagram, a caller's own producer — draws
 * onto a page through this surface. SVG and pixels come from `documonster/draw`
 * itself; a PDF page needs a page builder, so its surface lives here.
 */
export function createDrawSurface(
  page: PdfDrawPage,
  rect: DrawSurfaceRect,
  options?: PdfDrawSurfaceOptions
): DrawSurface {
  checkDrawSurfaceArguments(page, rect, options);
  return createPdfDrawSurface(page, rect, options);
}

// --- Page-size presets ---
export { PageSizes } from "@pdf/types";

// --- Cross-module converters (lazy: pull excel / word only when called) ---

/**
 * Convert an Excel workbook to a PDF. Dynamically loads the excel bridge.
 *
 * @param workbook - The workbook to export.
 * @param options - Export options. Pass sheet names in `sheets`.
 */
export function fromExcel(
  workbook: Workbook,
  options?: Omit<ExcelToPdfOptions, "sheets"> & {
    /**
     * The sheets to include, by name, in this order. Default — and with an empty list — all
     * visible ones. Compared as Excel compares names (ignoring case); a selected sheet is
     * exported even when hidden; a name that matches no sheet is refused with
     * `PdfArgumentError`.
     */
    readonly sheets?: readonly string[];
  }
): Promise<Uint8Array>;
/**
 * @deprecated Pass sheet names. Numeric positions are deprecated (removed in 1.0): they
 * are 1-based here but 0-based in `WordExcel.toDocx` and sheet ids in
 * `Workbook.getWorksheet`.
 *
 * The form that takes numeric positions in `sheets`; it works exactly as
 * before until it is removed.
 */
export function fromExcel(
  workbook: Workbook,
  // `sheets` is required here, so only a value that certainly carries a list able to hold a
  // number — a literal with one, a variable inferred as `number[]` — resolves to this form.
  options: Omit<ExcelToPdfOptions, "sheets"> & { readonly sheets: readonly (string | number)[] }
): Promise<Uint8Array>;
/**
 * Convert an Excel workbook to a PDF, with options typed as the published
 * {@link ExcelToPdfOptions}. Not deprecated: such a value may well hold only names, and
 * striking it through would flag valid code. Prefer names in `sheets`.
 */
export function fromExcel(workbook: Workbook, options?: ExcelToPdfOptions): Promise<Uint8Array>;
export async function fromExcel(
  workbook: Workbook,
  input?: Omit<ExcelToPdfOptions, "sheets"> & { readonly sheets?: readonly (string | number)[] }
): Promise<Uint8Array> {
  // The exporter only reads `sheets`, so the read-only list of the names form is passed as is.
  const options = input as ExcelToPdfOptions | undefined;
  // Checked before the bridge runs `recalculate`, which edits the workbook.
  checkPdfExportOptions("Pdf.fromExcel", options);
  if (options?.recalculate !== undefined && typeof options.recalculate !== "function") {
    throw fail(
      at("Pdf.fromExcel", "options"),
      ".recalculate",
      "a function, such as ExcelFormula.calculate",
      options.recalculate
    );
  }
  checkWorkbookHandle("Pdf.fromExcel", workbook);
  const { excelToPdf } = await import("@pdf/excel-bridge");
  return excelToPdf(workbook, options);
}

/** Render a single chart handle to a PDF. Dynamically loads the excel bridge. */
export async function fromChart(
  chart: ChartHandle,
  options?: ChartToPdfOptions
): Promise<Uint8Array> {
  checkChartToPdfOptions(options);
  checkChartHandle("Pdf.fromChart", chart);
  const { chartToPdf } = await import("@pdf/excel-bridge");
  return chartToPdf(chart, options);
}

/** Convert a DOCX document to a PDF. Dynamically loads the word bridge. */
export async function fromDocx(doc: DocxDocument, options?: DocxToPdfOptions): Promise<Uint8Array> {
  checkDocxToPdfOptions(options);
  if (typeof doc !== "object" || doc === null || !Array.isArray((doc as { body?: unknown }).body)) {
    throw new PdfArgumentError(
      "doc",
      "Pdf.fromDocx: doc must be a DocxDocument (a document with a body)"
    );
  }
  const { docxToPdf } = await import("@pdf/word-bridge");
  return docxToPdf(doc, options);
}

/**
 * The Word-chart → PDF renderer callback that draws a classic chart with the Excel chart
 * engine, for `Pdf.fromDocx`'s `chartRenderer` option.
 *
 * @deprecated Omit `chartRenderer`: the default draws every chart — classic and ChartEx —
 * with the Excel engine, and this callback draws a classic chart exactly as the default
 * does. Removed in 1.0.
 */
export const wordChartRenderer = async (): Promise<
  (
    chart: WordChart,
    page: PdfPageBuilder,
    rect: { x: number; y: number; width: number; height: number }
  ) => void
> => {
  const { createWordChartPdfRenderer } = await import("@pdf/word-chart-bridge");
  return createWordChartPdfRenderer();
};

/** A handle argument: an object. What it is an object *of* is the bridge's to check. */
/**
 * A workbook handle, checked cheaply: an object with the worksheet list every workbook has.
 * Any object used to pass (`{}` then failed with "reading 'slice'"), and a non-object was
 * reported as "must be an options object".
 */
function checkWorkbookHandle(where: string, value: unknown): void {
  if (
    typeof value !== "object" ||
    value === null ||
    !Array.isArray((value as { _worksheets?: unknown })._worksheets)
  ) {
    throw new PdfArgumentError(
      "workbook",
      `${where}: workbook must be a workbook (from Workbook.create or Workbook.read), got ${describeHandle(value)}`
    );
  }
}

/** A chart handle, checked cheaply: its worksheet and its part number. */
function checkChartHandle(where: string, value: unknown): void {
  const chart = value as { worksheet?: unknown; chartNumber?: unknown } | null;
  if (
    typeof value !== "object" ||
    chart === null ||
    typeof chart.worksheet !== "object" ||
    chart.worksheet === null ||
    typeof chart.chartNumber !== "number"
  ) {
    throw new PdfArgumentError(
      "chart",
      `${where}: chart must be a chart handle (from Chart.place or Chart.list), got ${describeHandle(value)}`
    );
  }
}

function describeHandle(value: unknown): string {
  if (value === null || typeof value !== "object") {
    return String(value);
  }
  return Array.isArray(value) ? "an array" : "an object that is not one";
}

const finite = () => s.number({ finite: true });
const positive = () =>
  s
    .number({ finite: true, expected: "a finite number > 0" })
    .refine(v => v > 0, "a finite number > 0");
const nonNegative = () => s.number({ finite: true, min: 0 });
const optionsObject = { expected: "an options object" };

const chartToPdfOptionsSchema = documentOptionsSchemaWith({
  pageWidth: positive().optional(),
  pageHeight: positive().optional(),
  width: positive().optional(),
  height: positive().optional(),
  margin: nonNegative().optional(),
  forceRaster: s.boolean().optional(),
  rasterScale: positive().optional(),
  // The topology inside is the chart engine's to check (`checkChartRenderOptions`).
  regionMap: s
    .custom<NonNullable<ChartToPdfOptions["regionMap"]>>(
      v => typeof v === "object" && v !== null && !Array.isArray(v),
      "an options object"
    )
    .optional()
});
const _chartToPdfOptionsMatch: SchemaMatches<typeof chartToPdfOptionsSchema, ChartToPdfOptions> =
  true;
void _chartToPdfOptionsMatch;

function checkChartToPdfOptions(options: ChartToPdfOptions | undefined): void {
  const where = "Pdf.fromChart";
  checkConverterOptions(where, chartToPdfOptionsSchema, options);
  // The properties go through `Pdf.Builder#setMetadata`, which takes only text XML can carry.
  checkMetadataText(where, (options ?? {}) as Record<string, unknown>, "options.");
}

const docxToPdfOptionsSchema = documentOptionsSchemaWithReadonly({
  pageWidth: positive().optional(),
  pageHeight: positive().optional(),
  margins: s
    .object(
      {
        top: nonNegative().optional(),
        right: nonNegative().optional(),
        bottom: nonNegative().optional(),
        left: nonNegative().optional()
      },
      optionsObject
    )
    .optional(),
  marginTop: nonNegative().optional(),
  marginBottom: nonNegative().optional(),
  marginLeft: nonNegative().optional(),
  marginRight: nonNegative().optional(),
  defaultFontFamily: s.string().optional(),
  defaultFont: s.string().optional(),
  defaultFontSize: positive().optional(),
  headerMargin: nonNegative().optional(),
  footerMargin: nonNegative().optional(),
  chartRenderer: s.function<NonNullable<DocxToPdfOptions["chartRenderer"]>>().optional()
});
const _docxToPdfOptionsMatch: SchemaMatches<typeof docxToPdfOptionsSchema, DocxToPdfOptions> = true;
void _docxToPdfOptionsMatch;

function checkDocxToPdfOptions(options: DocxToPdfOptions | undefined): void {
  const where = "Pdf.fromDocx";
  checkConverterOptions(where, docxToPdfOptionsSchema, options);
  // The properties go through `Pdf.Builder#setMetadata`, which takes only text XML can carry.
  checkMetadataText(where, (options ?? {}) as Record<string, unknown>, "options.");
}

const drawSurfaceRectSchema = s.object(
  {
    x: finite(),
    y: finite(),
    width: s.number({ finite: true }).refine(v => v > 0, "a finite number > 0"),
    height: s.number({ finite: true }).refine(v => v > 0, "a finite number > 0")
  },
  { ...optionsObject, readonly: true }
);
const _drawSurfaceRectMatch: SchemaMatches<typeof drawSurfaceRectSchema, DrawSurfaceRect> = true;
void _drawSurfaceRectMatch;

const drawSurfaceOptionsSchema = s.object(
  { scale: positive().optional() },
  { ...optionsObject, readonly: true }
);
const _drawSurfaceOptionsMatch: SchemaMatches<
  typeof drawSurfaceOptionsSchema,
  PdfDrawSurfaceOptions
> = true;
void _drawSurfaceOptionsMatch;

function checkDrawSurfaceArguments(
  page: PdfDrawPage,
  rect: DrawSurfaceRect,
  options: PdfDrawSurfaceOptions | undefined
): void {
  const where = "Pdf.createDrawSurface";
  check(drawSurfaceRectSchema, rect, at(where, "rect"));
  checkDrawPage(where, page);
  check(drawSurfaceOptionsSchema.optional(), options, at(where, "options"));
}
