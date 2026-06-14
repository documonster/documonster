/**
 * workbook-core — the low-level Workbook container layer.
 *
 * Holds the plain-data `WorkbookData` record interface and the low-level
 * registry accessors (defined-names, chart/chartEx entries, images, sheet-name
 * validation, sheet removal) that the worksheet module reaches up into.
 *
 * Exists to break the worksheet ↔ workbook import cycle: worksheet needs these
 * workbook-level registries, and the heavy `workbook.browser` module needs to
 * create worksheets. Keeping the registries here — above worksheet, below the
 * heavy workbook module — yields a clean one-directional graph:
 *
 *     workbook.browser.ts  (addWorksheet / importSheet / chartsheet / model / IO)
 *        ↓
 *     workbook-core.ts  (WorkbookData + registry accessors)
 *        ↓
 *     worksheet.ts / worksheet-core.ts / cell / row / column
 *
 * No file below this one imports the heavy `workbook.browser`.
 */
import type { ChartEntry, ChartExEntry } from "@excel/chart/chart";
import type { DefinedNamesData } from "@excel/defined-names";
import { ImageError, WorksheetNameError } from "@excel/errors";
import type { PivotTable } from "@excel/pivot-table";
import type {
  Font,
  ImageData,
  WorkbookProperties,
  CalculationProperties,
  WorkbookView,
  ThreadedCommentPerson
} from "@excel/types";
import type {
  WorkbookMedia,
  WorkbookProtectionModel,
  ExternalLinkModel
} from "@excel/workbook.browser";
import type { Worksheet } from "@excel/worksheet";
import type { ChartsheetModel } from "@excel/xlsx/xform/sheet/chartsheet-xform";
import type { XLSX } from "@excel/xlsx/xlsx.browser";

export interface WorkbookData {
  category: string;
  company: string;
  created: Date;
  description: string;
  keywords: string;
  manager: string;
  modified: Date;
  subject: string;
  title: string;
  creator?: string;
  lastModifiedBy?: string;
  lastPrinted?: Date;
  language?: string;
  revision?: number;
  contentStatus?: string;
  properties: Partial<WorkbookProperties>;
  calcProperties: Partial<CalculationProperties>;
  views: WorkbookView[];
  media: WorkbookMedia[];
  pivotTables: PivotTable[];
  protection?: WorkbookProtectionModel;
  externalLinks: ExternalLinkModel[];
  _worksheets: Worksheet[];
  _definedNames: DefinedNamesData;
  _themes?: unknown;
  _defaultFont?: Partial<Font>;
  _writerExternalLinkCache: Map<string, ExternalLinkModel>;
  _tableNames: Set<string>;
  _chartEntries: Record<number, ChartEntry>;
  _chartRels: Record<number, any[]>;
  _chartStyles: Record<number, Uint8Array>;
  _chartColors: Record<number, Uint8Array>;
  _chartExStyles: Record<number, Uint8Array>;
  _chartExColors: Record<number, Uint8Array>;
  _chartExEntries: Record<number, Uint8Array>;
  _chartExRels: Record<number, any[]>;
  _chartExStructuredEntries: Record<number, ChartExEntry>;
  _chartsheets: ChartsheetModel[];
  _persons: ThreadedCommentPerson[];
  _slicerParts: Record<string, Uint8Array>;
  _slicerCacheParts: Record<string, Uint8Array>;
  _timelineParts: Record<string, Uint8Array>;
  _timelineCacheParts: Record<string, Uint8Array>;
  _xlsx?: XLSX;
  userFunctions?: Map<
    string,
    {
      minArity: number;
      maxArity: number;
      invoke: (args: unknown[]) => unknown;
      volatile?: boolean;
    }
  >;
}

export function removeWorksheetEx(wb: WorkbookData, worksheet: Worksheet): void {
  // Release any workbook-wide table names this sheet held so the names can
  // be reused by future tables on other sheets without spurious "name
  // already exists" errors.
  const tables = worksheet.tables;
  if (tables) {
    for (const tableName of Object.keys(tables)) {
      wb._tableNames.delete(tableName.toLowerCase());
    }
  }
  wb._worksheets[worksheet.id] = undefined!;
}

export function getImage(wb: WorkbookData, id: number | string): WorkbookMedia | undefined {
  return wb.media[Number(id)];
}

export function nextChartNumber(wb: WorkbookData): number {
  const existing = Object.keys(wb._chartEntries).map(Number);
  return existing.length > 0 ? Math.max(...existing) + 1 : 1;
}

export function addChartEntry(wb: WorkbookData, entry: ChartEntry): void {
  wb._chartEntries[entry.chartNumber] = entry;
}

export function setChartStyle(wb: WorkbookData, chartNumber: number, data: Uint8Array): void {
  wb._chartStyles[chartNumber] = data;
}

export function setChartColors(wb: WorkbookData, chartNumber: number, data: Uint8Array): void {
  wb._chartColors[chartNumber] = data;
}

export function getChartEntry(wb: WorkbookData, chartNumber: number): ChartEntry | undefined {
  return wb._chartEntries[chartNumber];
}

export function removeChartEntry(wb: WorkbookData, chartNumber: number): void {
  delete wb._chartEntries[chartNumber];
  delete wb._chartRels[chartNumber];
  delete wb._chartStyles[chartNumber];
  delete wb._chartColors[chartNumber];
}

export function nextChartExNumber(wb: WorkbookData): number {
  const rawKeys = Object.keys(wb._chartExEntries ?? {}).map(Number);
  const structKeys = Object.keys(wb._chartExStructuredEntries ?? {}).map(Number);
  const combined = [...rawKeys, ...structKeys];
  return combined.length > 0 ? Math.max(...combined) + 1 : 1;
}

export function addChartExStructuredEntry(wb: WorkbookData, entry: ChartExEntry): void {
  if (!wb._chartExStructuredEntries) {
    wb._chartExStructuredEntries = {};
  }
  wb._chartExStructuredEntries[entry.chartExNumber] = entry;
}

export function getChartExStructuredEntry(
  wb: WorkbookData,
  chartExNumber: number
): ChartExEntry | undefined {
  return wb._chartExStructuredEntries?.[chartExNumber];
}

export function validateSheetName(
  wb: WorkbookData,
  name: string,
  existing?: Worksheet | { name: string }
): string {
  if (typeof name !== "string") {
    throw new WorksheetNameError("The name has to be a string.");
  }
  if (name === "") {
    throw new WorksheetNameError("The name can't be empty.");
  }
  if (name === "History") {
    throw new WorksheetNameError('The name "History" is protected. Please use a different name.');
  }
  // Illegal characters per Excel's own naming rules: asterisk (*),
  // question mark (?), colon (:), forward slash (/), backslash (\),
  // left bracket ([), right bracket (]). The chartsheet regex used
  // to omit `\\`; unified here so both families enforce the same
  // char set.
  if (/[*?:/\\[\]]/.test(name)) {
    throw new WorksheetNameError(
      `Sheet name ${name} cannot include any of the following characters: * ? : \\ / [ ]`
    );
  }
  if (/(^')|('$)/.test(name)) {
    throw new WorksheetNameError(
      `The first or last character of sheet name cannot be a single quotation mark: ${name}`
    );
  }
  if (name.length > 31) {
    if (process.env.NODE_ENV !== "production") {
      console.warn(`Sheet name ${name} exceeds 31 chars. This will be truncated`);
    }
    name = name.substring(0, 31);
  }
  const nameLower = name.toLowerCase();
  const duplicateWorksheet = wb._worksheets.find(
    ws => ws && ws !== existing && ws._name.toLowerCase() === nameLower
  );
  const duplicateChartsheet = wb._chartsheets.find(
    cs => cs && cs !== existing && cs.name.toLowerCase() === nameLower
  );
  if (duplicateWorksheet || duplicateChartsheet) {
    throw new WorksheetNameError(`Sheet name already exists: ${name}`);
  }
  return name;
}

export function removeChartExStructuredEntry(wb: WorkbookData, chartExNumber: number): void {
  if (wb._chartExStructuredEntries) {
    delete wb._chartExStructuredEntries[chartExNumber];
  }
  delete wb._chartExEntries[chartExNumber];
  delete wb._chartExRels[chartExNumber];
  delete wb._chartExStyles[chartExNumber];
  delete wb._chartExColors[chartExNumber];
}

export function getDefinedNames(wb: WorkbookData): DefinedNamesData {
  return wb._definedNames;
}
export function addWorkbookImage(wb: WorkbookData, image: ImageData): number {
  const { svg, ...raster } = image;
  if (
    svg &&
    raster.link &&
    raster.buffer == null &&
    raster.base64 == null &&
    raster.filename == null
  ) {
    // An SVG companion needs an embedded raster fallback; a *linked* (external)
    // raster has no package part to attach the svgBlip extension to.
    throw new ImageError(
      "An SVG image requires an embedded raster fallback (buffer/base64/filename); it cannot be combined with an external link."
    );
  }
  const id = wb.media.length;
  const rasterMedia: WorkbookMedia = { ...raster, type: "image" };
  wb.media.push(rasterMedia);

  if (svg) {
    // Register the SVG as a second `type: "image"` media so it flows through
    // the existing media naming, content-types, and zip-writing paths. Link
    // it back to the raster blip so the drawing serializer can emit the
    // asvg:svgBlip extension.
    const svgId = wb.media.length;
    wb.media.push({ ...svg, type: "image", extension: "svg" });
    rasterMedia.svgMediaId = svgId;
  }

  return id;
}
