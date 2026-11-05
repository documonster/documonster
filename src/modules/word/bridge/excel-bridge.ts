/**
 * Excel-to-Word Bridge
 *
 * Converts an Excel Workbook into a DocxDocument model.
 * This is the ONLY file in the Word module that imports from @excel.
 *
 * @example
 * ```typescript
 * import { Workbook } from "documonster/excel";
 * import { WordExcel } from "documonster/word/excel";
 *
 * const wb = await Workbook.read(buffer);
 * const doc = WordExcel.toDocx(wb);
 * ```
 */

import { buildChartExModel } from "@excel/chart/build/chart-ex-builder";
import type {
  AddChartExOptions,
  ChartExModel,
  ChartExType
} from "@excel/chart/model/chart-ex-types";
import type {
  AxisDataSource,
  BarChartGroup,
  BarDirection,
  BarGrouping,
  BarSeries,
  ChartModel,
  ChartTypeGroup,
  LegendPosition,
  LineChartGroup,
  LineGrouping,
  LineSeries,
  NumberDataSource,
  PieChartGroup,
  PieSeries,
  AreaChartGroup,
  AreaSeries,
  ScatterChartGroup,
  ScatterSeries,
  RadarChartGroup,
  RadarSeries,
  DoughnutChartGroup,
  SurfaceChartGroup,
  SurfaceSeries,
  ChartAxis as ExcelChartAxis,
  DataPoint,
  ShapeProperties
} from "@excel/chart/model/types";
import { renderChartSvg } from "@excel/chart/render/chart-renderer";
import { checkChartRenderOptions } from "@excel/chart/render/render-options";
import { renderChartEx } from "@excel/chart/serialize/chart-ex-serialize";
import type { ChartThemeColors } from "@excel/chart/shared/chart-utils";
import type { CellData as ExcelCell } from "@excel/core/cell";
import {
  cellAlignment,
  cellBorder,
  cellFill,
  cellFont,
  cellGetValue,
  cellHyperlink,
  cellDisplayText,
  cellSetValue,
  cellType
} from "@excel/core/cell";
import { ValueType } from "@excel/core/enums";
import { rowCellCount } from "@excel/core/row";
import { addWorksheet, createWorkbook, getWorksheets } from "@excel/core/workbook";
// Use the browser base class so the public `excelToDocx(workbook)` signature
// is callable from both the Node entry (where `Workbook` is the Node subclass
// — trivially assignable to the base) and the browser entry (where `Workbook`
// is already the base). Importing the Node alias `@excel/workbook` would force
// browser consumers to satisfy `xlsx.readFile`/`writeFile`, which the browser
// XLSX surface intentionally omits.
import type { WorkbookData } from "@excel/core/workbook.browser";
import {
  getCell,
  getColumn,
  getRow,
  getRowCount,
  getSheetName,
  rowGetCell
} from "@excel/core/worksheet";
import type { Worksheet } from "@excel/core/worksheet";
import { ExcelChartOptionsError } from "@excel/errors";
import { Workbook } from "@excel/index";
import type {
  Font as ExcelFont,
  Fill as ExcelFill,
  Borders as ExcelBorders,
  Border as ExcelBorder,
  Alignment as ExcelAlignment,
  Color as ExcelColor
} from "@excel/types";
import { sheetNameKey } from "@utils/name-key";
import { CELL_THEME_PALETTE } from "@utils/theme-colors";
import {
  CHART_TYPES,
  checkArray,
  checkBoolean,
  checkChartSeries as checkWordChartSeries,
  checkFinite,
  checkHandle,
  checkHeadingLevel,
  checkInteger,
  checkOneOf,
  checkRecord,
  checkString,
  invalidArgument
} from "@word/core/argument-checks";
import {
  chartSpaceKeys,
  keptChild,
  keptPoints,
  plotAreaKeys,
  seriesKeys
} from "@word/core/chart-kept";
import type { Mutable } from "@word/core/internal-utils";
import { extractParagraphText } from "@word/core/text-utils";
import { DocxArgumentError } from "@word/errors";
import type {
  Alignment,
  Chart,
  ChartKeptNode,
  ChartSeries as WordChartSeries,
  ChartType as WordChartType,
  DocxDocument,
  BodyContent,
  Hyperlink,
  Paragraph,
  ParagraphChild,
  Run,
  RunProperties,
  StyleDef,
  Table,
  TableRow,
  TableCell,
  TableProperties,
  TableCellProperties,
  Border,
  Shading
} from "@word/types";
import { EMU_PER_INCH, charWidthToPixel, pixelToPoints, ptToTwips } from "@word/units";

// =============================================================================
// Public API
// =============================================================================

/** Options for Excel → DOCX conversion. */
export interface ExcelToDocxOptions {
  /**
   * Which worksheets to include, in this order. Default: all visible ones.
   *
   * Pass sheet names. Numeric positions are deprecated (removed in 1.0): they
   * are 0-based here but 1-based in `Pdf.fromExcel` and sheet ids in
   * `Workbook.getWorksheet`.
   *
   * A name is compared as Excel compares names (ignoring case). A number is a
   * 0-based position among the worksheets, not counting very hidden ones. A
   * name or position that matches no worksheet is refused with
   * `DocxArgumentError`.
   */
  readonly sheets?: readonly (string | number)[];
  /** Include sheet name as heading before each table. Default: true. */
  readonly includeSheetHeadings?: boolean;
  /** Heading level for sheet names (1-6). Default: 2. */
  readonly sheetHeadingLevel?: number;
  /** Include a title page with workbook metadata. Default: false. */
  readonly includeTitlePage?: boolean;
  /** Maximum columns to include per sheet (avoids excessively wide tables). Default: 50. */
  readonly maxColumns?: number;
  /** Maximum rows per sheet to include. Default: 10000. */
  readonly maxRows?: number;
  /** Preserve cell formatting (bold, italic, colors, etc.). Default: true. */
  readonly preserveFormatting?: boolean;
  /** Include cell borders in table. Default: true. */
  readonly includeBorders?: boolean;
}

function checkExcelToDocxOptions(options: ExcelToDocxOptions | undefined, where: string): void {
  checkRecord(options, "options", where, { optional: true });
  if (options?.sheets !== undefined) {
    checkArray(options.sheets, "options.sheets", where);
    options.sheets.forEach((sheet, i) => {
      if (typeof sheet !== "string") {
        checkInteger(sheet, `options.sheets[${i}]`, where, { min: 0 });
      }
    });
  }
  // Written as the `HeadingN` style of the sheet titles.
  if (options?.sheetHeadingLevel !== undefined) {
    checkHeadingLevel(options.sheetHeadingLevel, "options.sheetHeadingLevel", where);
  }
  checkInteger(options?.maxColumns, "options.maxColumns", where, { optional: true, min: 0 });
  checkInteger(options?.maxRows, "options.maxRows", where, { optional: true, min: 0 });
  for (const key of [
    "includeSheetHeadings",
    "includeTitlePage",
    "preserveFormatting",
    "includeBorders"
  ] as const) {
    checkBoolean(options?.[key], `options.${key}`, where, { optional: true });
  }
}

/**
 * Convert an Excel Workbook to a DocxDocument model.
 *
 * Creates a Word document with tables representing each worksheet's data.
 * Preserves cell formatting (fonts, colors, alignment) as run/paragraph properties.
 *
 * @param workbook - The Excel Workbook to convert.
 * @param options - Conversion options. Pass sheet names in `sheets`.
 * @returns A DocxDocument model ready for packaging, PDF conversion, or Markdown output.
 */
export function excelToDocx(
  workbook: WorkbookData,
  options?: Omit<ExcelToDocxOptions, "sheets"> & {
    /**
     * The worksheets to include, by name, in this order. Default: all
     * visible ones. Compared as Excel compares names (ignoring case); a name
     * that matches no worksheet is refused with `DocxArgumentError`.
     */
    readonly sheets?: readonly string[];
  }
): DocxDocument;
/**
 * @deprecated Pass sheet names. Numeric positions are deprecated (removed in 1.0): they
 * are 0-based here but 1-based in `Pdf.fromExcel` and sheet ids in
 * `Workbook.getWorksheet`.
 *
 * The form that takes numeric positions in `sheets`; it works exactly as
 * before until it is removed.
 */
export function excelToDocx(
  workbook: WorkbookData,
  // `sheets` is required here, so only a value that certainly carries a list able to hold a
  // number — a literal with one, a variable inferred as `number[]` — resolves to this form.
  options: Omit<ExcelToDocxOptions, "sheets"> & { readonly sheets: readonly (string | number)[] }
): DocxDocument;
/**
 * Convert an Excel Workbook to a DocxDocument model, with options typed as the published
 * {@link ExcelToDocxOptions}. Not deprecated: such a value may well hold only names, and
 * striking it through would flag valid code. Prefer names in `sheets`.
 */
export function excelToDocx(workbook: WorkbookData, options?: ExcelToDocxOptions): DocxDocument;
export function excelToDocx(workbook: WorkbookData, options?: ExcelToDocxOptions): DocxDocument {
  // Excel's own test for a workbook handle (`checkWorkbook`): it carries its sheet list.
  if ((workbook as { _worksheets?: unknown } | null | undefined)?._worksheets === undefined) {
    invalidArgument(
      "WordExcel.toDocx",
      "workbook",
      "a workbook (from Workbook.create or Workbook.read)",
      workbook
    );
  }
  checkExcelToDocxOptions(options, "WordExcel.toDocx");
  const opts: Required<ExcelToDocxOptions> = {
    sheets: options?.sheets ?? [],
    includeSheetHeadings: options?.includeSheetHeadings ?? true,
    sheetHeadingLevel: options?.sheetHeadingLevel ?? 2,
    includeTitlePage: options?.includeTitlePage ?? false,
    maxColumns: options?.maxColumns ?? 50,
    maxRows: options?.maxRows ?? 10000,
    preserveFormatting: options?.preserveFormatting ?? true,
    includeBorders: options?.includeBorders ?? true
  };

  const body: BodyContent[] = [];

  // Title page
  if (opts.includeTitlePage) {
    const title = workbook.creator || "Workbook";
    body.push(heading(title, 1));
    if (workbook.created) {
      body.push(textParagraph(`Created: ${fmtDate(workbook.created)}`));
    }
    if (workbook.lastModifiedBy) {
      body.push(textParagraph(`Author: ${workbook.lastModifiedBy}`));
    }
    body.push(textParagraph(""));
  }

  // Determine which sheets to process
  const worksheets = selectSheets(workbook, opts.sheets);

  for (const ws of worksheets) {
    if (opts.includeSheetHeadings) {
      body.push(heading(getSheetName(ws), opts.sheetHeadingLevel));
    }

    const table = sheetToTable(ws, opts);
    if (table) {
      body.push(table);
      body.push(textParagraph(""));
    }
  }

  if (body.length === 0) {
    body.push(textParagraph("(Empty workbook)"));
  }

  // Register the heading styles we actually used (plus Normal as the
  // basedOn root). Without these entries Word logs "missing referenced
  // style" warnings on every heading inserted by the bridge.
  const styles: StyleDef[] = [
    { type: "paragraph", styleId: "Normal", name: "Normal", isDefault: true, qFormat: true }
  ];
  const usedHeadingLevels = new Set<number>();
  if (opts.includeTitlePage) {
    usedHeadingLevels.add(1);
  }
  if (opts.includeSheetHeadings) {
    usedHeadingLevels.add(opts.sheetHeadingLevel);
  }
  for (const lvl of usedHeadingLevels) {
    styles.push({
      type: "paragraph",
      styleId: `Heading${lvl}`,
      name: `heading ${lvl}`,
      basedOn: "Normal",
      next: "Normal",
      qFormat: true,
      uiPriority: 9,
      runProperties: {
        font: "Calibri Light",
        color: lvl <= 2 ? "2F5496" : "1F3763",
        size: lvl === 1 ? 32 : lvl === 2 ? 26 : 24,
        bold: true
      }
    });
  }

  return {
    body,
    styles,
    coreProperties: {
      title: workbook.creator ? `${workbook.creator} - Export` : "Excel Export",
      creator: workbook.lastModifiedBy ?? workbook.creator,
      created: workbook.created instanceof Date ? workbook.created : undefined,
      modified: workbook.modified instanceof Date ? workbook.modified : undefined
    }
  };
}

// =============================================================================
// Sheet Selection
// =============================================================================

function selectSheets(workbook: WorkbookData, filter: readonly (string | number)[]): Worksheet[] {
  const worksheets = getWorksheets(workbook).filter(ws => ws);
  const all = worksheets.filter(ws => ws.state !== "veryHidden");

  if (filter.length === 0) {
    return all.filter(ws => ws.state !== "hidden");
  }

  // A selector that matches no sheet is refused, as `Pdf.fromExcel` refuses
  // it: skipping it dropped a requested sheet without a word.
  return filter.map((spec, i) => {
    const ws =
      typeof spec === "number"
        ? all[spec]
        : worksheets.find(s => sheetNameKey(getSheetName(s)) === sheetNameKey(spec));
    if (ws === undefined) {
      invalidArgument(
        "WordExcel.toDocx",
        `options.sheets[${i}]`,
        typeof spec === "number"
          ? `a 0-based worksheet position below ${all.length}`
          : `the name of one of the workbook's worksheets (${worksheets
              .map(s => JSON.stringify(getSheetName(s)))
              .join(", ")})`,
        spec
      );
    }
    return ws;
  });
}

// =============================================================================
// Worksheet → Table
// =============================================================================

function sheetToTable(ws: Worksheet, opts: Required<ExcelToDocxOptions>): Table | null {
  const rowCount = Math.min(getRowCount(ws), opts.maxRows);
  if (rowCount === 0) {
    return null;
  }

  // Find actual column range
  let maxCol = 0;
  for (let r = 1; r <= rowCount; r++) {
    const row = getRow(ws, r);
    if (rowCellCount(row) > maxCol) {
      maxCol = rowCellCount(row);
    }
  }
  maxCol = Math.min(maxCol, opts.maxColumns);
  if (maxCol === 0) {
    return null;
  }

  const rows: TableRow[] = [];

  for (let r = 1; r <= rowCount; r++) {
    const row = getRow(ws, r);

    // Skip completely empty rows
    let hasContent = false;
    for (let c = 1; c <= maxCol; c++) {
      const cell = rowGetCell(row, c);
      if (cellType(cell) !== ValueType.Null && cellType(cell) !== ValueType.Merge) {
        hasContent = true;
        break;
      }
    }
    if (!hasContent) {
      continue;
    }

    const cells: TableCell[] = [];
    for (let c = 1; c <= maxCol; c++) {
      const cell = rowGetCell(row, c);
      cells.push(convertCell(cell, opts));
    }
    rows.push({ cells });
  }

  if (rows.length === 0) {
    return null;
  }

  // Column widths from worksheet definitions (Excel char width → twips),
  // via the shared char-width→pixel→point chain so the table column widths
  // match the Excel / PDF column geometry. MDW = 7 (Calibri 11pt default).
  const columnWidths: number[] = [];
  for (let c = 1; c <= maxCol; c++) {
    const col = getColumn(ws, c);
    const charWidth = col.width ?? 10;
    const pt = pixelToPoints(charWidthToPixel(charWidth, 7));
    columnWidths.push(ptToTwips(pt));
  }

  const tableProps: TableProperties = {
    width: { type: "auto", value: 0 }
  };

  return {
    type: "table",
    properties: tableProps,
    rows,
    columnWidths
  };
}

// =============================================================================
// Cell → TableCell
// =============================================================================

function convertCell(cell: ExcelCell, opts: Required<ExcelToDocxOptions>): TableCell {
  const text = renderCellText(cell);
  const children: Run[] = [];

  if (text) {
    if (cellType(cell) === ValueType.RichText) {
      const value = cellGetValue(cell) as {
        richText?: ReadonlyArray<{ text?: string; font?: Partial<ExcelFont> }>;
      };
      if (value?.richText) {
        // Preserve rich text segments
        for (const segment of value.richText) {
          const runProps = opts.preserveFormatting ? fontToRun(segment.font) : undefined;
          children.push({
            properties: runProps,
            content: [{ type: "text", text: segment.text ?? "" }]
          });
        }
      }
    } else {
      const runProps = opts.preserveFormatting ? fontToRun(cellFont(cell)) : undefined;
      children.push({
        properties: runProps,
        content: [{ type: "text", text }]
      });
    }
  }

  const para: Paragraph = {
    type: "paragraph",
    properties: opts.preserveFormatting ? alignmentToParaProps(cellAlignment(cell)) : undefined,
    children: wrapHyperlink(cellHyperlink(cell), children)
  };

  const cellProps: Mutable<TableCellProperties> = {};

  if (opts.preserveFormatting && cellFill(cell)) {
    const shading = fillToShading(cellFill(cell));
    if (shading) {
      cellProps.shading = shading;
    }
  }

  if (opts.includeBorders && cellBorder(cell)) {
    const borders = bordersToCellBorders(cellBorder(cell));
    if (borders) {
      cellProps.borders = borders;
    }
  }

  return {
    content: [para],
    properties: Object.keys(cellProps).length > 0 ? cellProps : undefined
  };
}

/**
 * Build a paragraph's children for a converted cell, wrapping the runs
 * in a Word {@link Hyperlink} when the source Excel cell carries one.
 *
 * Excel cell hyperlinks are external URLs (or `#Sheet!A1` internal
 * references). We map `#…` targets to a Word anchor and everything else
 * to an external `url`; the packager assigns the relationship id on
 * write. An empty cell still produces a single empty run so the table
 * structure stays intact.
 */
function wrapHyperlink(
  hyperlink: string | undefined,
  runs: readonly Run[]
): readonly ParagraphChild[] {
  const children: Run[] = runs.length > 0 ? [...runs] : [{ content: [{ type: "text", text: "" }] }];
  if (!hyperlink) {
    return children;
  }
  const link: Hyperlink = hyperlink.startsWith("#")
    ? { type: "hyperlink", anchor: hyperlink.slice(1), children }
    : { type: "hyperlink", url: hyperlink, children };
  return [link];
}

function renderCellText(cell: ExcelCell): string {
  // As the cell shows in Excel, `Cell.getDisplayText` and `Pdf.fromExcel`: its
  // number format applied to the value (a date by its calendar fields, so the
  // text does not depend on the machine's timezone), the text of anything else.
  // A cell merged into another holds nothing of its own.
  return cellType(cell) === ValueType.Merge ? "" : cellDisplayText(cell);
}

// =============================================================================
// Style Conversion
// =============================================================================

function fontToRun(font: Partial<ExcelFont> | undefined): RunProperties | undefined {
  if (!font) {
    return undefined;
  }

  const props: Mutable<RunProperties> = {};
  let has = false;

  if (font.bold) {
    props.bold = true;
    has = true;
  }
  if (font.italic) {
    props.italic = true;
    has = true;
  }
  if (font.strike) {
    props.strike = true;
    has = true;
  }
  if (font.underline && font.underline !== "none") {
    props.underline =
      typeof font.underline === "string"
        ? (font.underline as RunProperties["underline"])
        : "single";
    has = true;
  }
  if (font.size) {
    props.size = Math.round(font.size * 2); // pt → half-points
    has = true;
  }
  if (font.name) {
    props.font = { ascii: font.name, hAnsi: font.name };
    has = true;
  }
  if (font.color) {
    const hex = excelColorToHex(font.color);
    if (hex) {
      props.color = hex;
      has = true;
    }
  }
  if (font.vertAlign === "superscript") {
    props.vertAlign = "superscript";
    has = true;
  } else if (font.vertAlign === "subscript") {
    props.vertAlign = "subscript";
    has = true;
  }

  return has ? props : undefined;
}

function alignmentToParaProps(
  alignment: Partial<ExcelAlignment> | undefined
): Paragraph["properties"] | undefined {
  if (!alignment) {
    return undefined;
  }

  const map: Record<string, string> = {
    left: "left",
    center: "center",
    right: "right",
    justify: "both",
    fill: "left",
    centerContinuous: "center",
    distributed: "distribute"
  };

  const mapped = alignment.horizontal ? map[alignment.horizontal] : undefined;
  if (mapped) {
    return { alignment: mapped as Alignment };
  }
  return undefined;
}

function fillToShading(fill: ExcelFill | undefined): Shading | undefined {
  if (!fill || fill.type !== "pattern" || !fill.fgColor) {
    return undefined;
  }
  const hex = excelColorToHex(fill.fgColor);
  if (hex) {
    return { fill: hex, pattern: "clear" };
  }
  return undefined;
}

function bordersToCellBorders(
  border: Partial<ExcelBorders> | undefined
): Record<string, Border> | undefined {
  if (!border) {
    return undefined;
  }

  const result: Record<string, Border> = {};
  let has = false;

  for (const side of ["top", "bottom", "left", "right"] as const) {
    if (border[side]) {
      const b = singleBorder(border[side]);
      if (b) {
        result[side] = b;
        has = true;
      }
    }
  }

  return has ? result : undefined;
}

function singleBorder(border: Partial<ExcelBorder> | undefined): Border | undefined {
  if (!border || !border.style) {
    return undefined;
  }

  const styleMap: Record<string, Border["style"]> = {
    thin: "single",
    medium: "single",
    thick: "thick",
    dotted: "dotted",
    dashed: "dashed",
    dashDot: "dashDotStroked",
    dashDotDot: "dashDotStroked",
    double: "double",
    hair: "single",
    mediumDashed: "dashed",
    mediumDashDot: "dashDotStroked",
    mediumDashDotDot: "dashDotStroked",
    slantDashDot: "dashDotStroked"
  };

  const docxStyle: Border["style"] = styleMap[border.style] ?? "single";
  const color = border.color ? (excelColorToHex(border.color) ?? "auto") : "auto";
  const size = border.style === "thick" || border.style === "medium" ? 12 : 4;

  return { style: docxStyle, size, space: 0, color };
}

// =============================================================================
// Utilities
// =============================================================================

function excelColorToHex(color: Partial<ExcelColor> | undefined): string | undefined {
  if (!color) {
    return undefined;
  }
  if (color.argb) {
    const argb = String(color.argb);
    return argb.length === 8 ? argb.slice(2) : argb;
  }
  if (color.theme !== undefined) {
    return CELL_THEME_PALETTE[color.theme] ?? "000000";
  }
  return undefined;
}

function heading(text: string, level: number): Paragraph {
  return {
    type: "paragraph",
    properties: { style: `Heading${Math.min(Math.max(level, 1), 6)}` },
    children: [{ properties: { bold: true }, content: [{ type: "text", text }] }]
  };
}

function textParagraph(text: string): Paragraph {
  return {
    type: "paragraph",
    children: [{ content: [{ type: "text", text }] }]
  };
}

function fmtDate(date: Date | string): string {
  if (typeof date === "string") {
    return date;
  }
  try {
    return date.toISOString().slice(0, 10);
  } catch {
    return "";
  }
}

// =============================================================================
// ChartEx Bridge — Expose ChartEx rendering to Word module
// =============================================================================

export type { AddChartExOptions, ChartExModel, ChartExType, ChartModel };

/** Options for creating a ChartEx in a Word document. */
export interface DocxChartExOptions {
  /** ChartEx type (sunburst, treemap, waterfall, funnel, histogram, pareto, boxWhisker, regionMap). */
  readonly type: ChartExType;
  /** Series data (with literal categories/values for Word context). */
  readonly series: readonly {
    readonly name: string;
    readonly categories?: readonly string[];
    readonly values: readonly number[];
  }[];
  /** Chart title. */
  readonly title?: string | null;
  /** Show legend. Default: true. */
  readonly showLegend?: boolean;
  /** Legend position. */
  readonly legendPosition?: "b" | "l" | "r" | "t" | "tr";
}

const CHART_EX_TYPES: ReadonlySet<string> = new Set<ChartExType>([
  "sunburst",
  "treemap",
  "waterfall",
  "funnel",
  "histogram",
  "pareto",
  "boxWhisker",
  "regionMap"
]);

const CHART_EX_LEGEND_POSITIONS: ReadonlySet<string> = new Set<
  NonNullable<DocxChartExOptions["legendPosition"]>
>(["b", "l", "r", "t", "tr"]);

/** Chart series as the Word chart helpers take them: a name, categories and finite values. */
function checkChartSeries(series: unknown, field: string, where: string): void {
  checkArray(series, field, where);
  (series as unknown[]).forEach((s, i) => {
    const at = `${field}[${i}]`;
    checkRecord(s, at, where);
    const v = s as { name?: unknown; categories?: unknown; values?: unknown };
    checkString(v.name, `${at}.name`, where);
    if (v.categories !== undefined) {
      checkArray(v.categories, `${at}.categories`, where);
      (v.categories as unknown[]).forEach((c, j) =>
        checkString(c, `${at}.categories[${j}]`, where)
      );
    }
    checkArray(v.values, `${at}.values`, where);
    (v.values as unknown[]).forEach((n, j) => checkFinite(n, `${at}.values[${j}]`, where));
  });
}

/**
 * Build a ChartEx XML part for embedding in a Word document.
 *
 * This bridges the Excel ChartEx renderer into the Word module,
 * generating a complete `cx:chartSpace` XML string from high-level options.
 *
 * @param options - Chart configuration.
 * @returns The complete ChartEx XML string for the chart part.
 */
export function buildWordChartExXml(options: DocxChartExOptions): string {
  const where = "WordExcel.toChartExXml";
  checkRecord(options, "options", where);
  checkOneOf(options.type, CHART_EX_TYPES, "options.type", where);
  checkChartSeries(options.series, "options.series", where);
  if (options.title !== undefined && options.title !== null) {
    checkString(options.title, "options.title", where);
  }
  checkBoolean(options.showLegend, "options.showLegend", where, { optional: true });
  checkOneOf(options.legendPosition, CHART_EX_LEGEND_POSITIONS, "options.legendPosition", where, {
    optional: true
  });
  // Convert Word-friendly options to Excel's AddChartExOptions format
  const excelOptions: AddChartExOptions = {
    type: options.type,
    series: options.series.map(s => ({
      name: s.name,
      values: "" as string, // headless chart — no worksheet formula, use literalValues instead
      literalCategories: s.categories ? [...s.categories] : undefined,
      literalValues: [...s.values]
    })) as AddChartExOptions["series"],
    title: options.title === null ? null : (options.title ?? undefined),
    showLegend: options.showLegend ?? true,
    legendPosition: options.legendPosition
  };

  // Build the model
  const model: ChartExModel = buildChartExModel(excelOptions);

  // Render to XML (force structural since it's a new chart)
  return renderChartEx(model, { forceStructural: true });
}

// =============================================================================
// Classic Chart Bridge — Render Word Chart definitions to SVG
// =============================================================================

// The chart engine's theme-colour type, named in `DocxChartSvgOptions`; re-exported here
// because only this bridge may reach into `@excel`.
export type { ChartThemeColors } from "@excel/chart/shared/chart-utils";

/** Options for {@link renderWordChartSvg} (`WordExcel.toChartSvg`). */
export interface DocxChartSvgOptions {
  /**
   * The colour scheme the chart's automatic series colours and scheme colours resolve against:
   * six hex digits per theme slot — the document's `doc.theme.colorScheme.colors` (which
   * `Layout.toSvg` hands its `renderChart` callback as `context.themeColors`). Default: the
   * default Office theme.
   */
  readonly themeColors?: ChartThemeColors;
}

/**
 * Render a Word Chart to SVG string for HTML embedding.
 *
 * Bridges Word's Chart definition to Excel's chart renderer by constructing
 * a ChartModel with literal (inline) data and delegating to renderChartSvg.
 *
 * @param chart - The Word Chart definition.
 * @param options - The document theme's colours (`themeColors`).
 * @returns SVG markup string.
 * @throws {DocxArgumentError} `chart` is not a Word chart, or `options` is malformed (field
 *   `options.themeColors.<slot>`, …).
 */
export function renderWordChartSvg(chart: Chart, options?: DocxChartSvgOptions): string {
  const where = "WordExcel.toChartSvg";
  checkWordChart(chart, where);
  checkRecord(options, "options", where, { optional: true });
  const themeColors = options?.themeColors;
  try {
    checkChartRenderOptions(where, { themeColors });
  } catch (error) {
    // The chart engine's check, in this module's class.
    if (error instanceof ExcelChartOptionsError) {
      throw new DocxArgumentError(error.field, error.message, { cause: error });
    }
    throw error;
  }
  const model = chartModelOf(chart);
  return renderChartSvg(model, {
    width: chart.width ? Math.round((chart.width / EMU_PER_INCH) * 96) : undefined,
    height: chart.height ? Math.round((chart.height / EMU_PER_INCH) * 96) : undefined,
    ...(themeColors === undefined ? {} : { themeColors })
  });
}

// =============================================================================
// Embedded XLSX Generation for Chart Data
// =============================================================================

/**
 * Generate an embedded xlsx for chart data.
 * Uses the full Excel xlsx writer for maximum compatibility.
 *
 * The resulting xlsx contains a single Sheet1 with:
 * - Row 1: header (empty cell A1, then series names in B1, C1, ...)
 * - Row 2+: category in column A, values in subsequent columns
 *
 * @param series - Chart data series
 * @returns xlsx file as Uint8Array
 */
export async function generateChartEmbeddedXlsx(
  series: readonly { name: string; categories: readonly string[]; values: readonly number[] }[]
): Promise<Uint8Array> {
  checkWordChartSeries(series, "series", "WordExcel.toChartXlsx");
  // `createWorkbook` comes from the static import at the top with `addWorksheet`, which the next
  // line calls. It was an `await import()` of the very same module, which bought nothing — the
  // module was already static in this file, so rolldown reported the pair as
  // `INEFFECTIVE_DYNAMIC_IMPORT` — and cost a round trip through the module registry per call.
  const wb = createWorkbook();
  const ws = addWorksheet(wb, "Sheet1");

  // Row 1: headers — A1 empty, B1..N1 = series names
  for (let c = 0; c < series.length; c++) {
    cellSetValue(getCell(ws, 1, c + 2), series[c].name);
  }

  // Rows 2+: categories and values
  const categories = series.length > 0 ? series[0].categories : [];
  for (let r = 0; r < categories.length; r++) {
    cellSetValue(getCell(ws, r + 2, 1), categories[r]);
    for (let c = 0; c < series.length; c++) {
      const val = r < series[c].values.length ? series[c].values[r] : 0;
      cellSetValue(getCell(ws, r + 2, c + 2), val);
    }
  }

  return new Uint8Array(await Workbook.toBuffer(wb));
}

// =============================================================================
// Word Chart → ChartModel mapping
// =============================================================================

/** Map Word's ChartType to Excel's chart type + modifiers. */
interface ExcelChartMapping {
  type:
    | "bar"
    | "bar3D"
    | "line"
    | "line3D"
    | "pie"
    | "pie3D"
    | "doughnut"
    | "area"
    | "area3D"
    | "scatter"
    | "radar"
    | "bubble"
    | "stock"
    | "surface"
    | "surface3D";
  barDir?: BarDirection;
  grouping?: BarGrouping | LineGrouping;
  scatterStyle?: "lineMarker" | "smooth";
  radarStyle?: "standard" | "filled";
  wireframe?: boolean;
}

function mapWordChartType(wordType: WordChartType): ExcelChartMapping {
  switch (wordType) {
    case "bar":
      return { type: "bar", barDir: "bar", grouping: "clustered" };
    case "barStacked":
      return { type: "bar", barDir: "bar", grouping: "stacked" };
    case "barPercentStacked":
      return { type: "bar", barDir: "bar", grouping: "percentStacked" };
    case "column":
      return { type: "bar", barDir: "col", grouping: "clustered" };
    case "columnStacked":
      return { type: "bar", barDir: "col", grouping: "stacked" };
    case "columnPercentStacked":
      return { type: "bar", barDir: "col", grouping: "percentStacked" };
    case "line":
      return { type: "line", grouping: "standard" };
    case "lineStacked":
      return { type: "line", grouping: "stacked" };
    case "lineMarked":
      return { type: "line", grouping: "standard" };
    case "pie":
      return { type: "pie" };
    case "pie3D":
      return { type: "pie3D" };
    case "doughnut":
      return { type: "doughnut" };
    case "area":
      return { type: "area", grouping: "standard" };
    case "areaStacked":
      return { type: "area", grouping: "stacked" };
    case "scatter":
      return { type: "scatter", scatterStyle: "lineMarker" };
    case "scatterSmooth":
      return { type: "scatter", scatterStyle: "smooth" };
    case "radar":
      return { type: "radar", radarStyle: "standard" };
    case "radarFilled":
      return { type: "radar", radarStyle: "filled" };
    case "bubble":
      return { type: "bubble" };
    case "stock":
      return { type: "stock" };
    case "surface":
      return { type: "surface" };
    case "surface3D":
      return { type: "surface3D" };
    case "surfaceWireframe":
      return { type: "surface", wireframe: true };
    case "surfaceWireframe3D":
      return { type: "surface3D", wireframe: true };
    default:
      return { type: "bar", barDir: "col", grouping: "clustered" };
  }
}

/** Build an AxisDataSource from literal string categories. */
function makeLiteralCat(categories: readonly string[]): AxisDataSource {
  return {
    strLit: {
      pointCount: categories.length,
      points: categories.map((value, index) => ({ index, value }))
    }
  };
}

/** Build a NumberDataSource from literal numeric values. */
function makeLiteralVal(values: readonly number[]): NumberDataSource {
  return {
    numLit: {
      pointCount: values.length,
      points: values.map((value, index) => ({ index, value }))
    }
  };
}

/** Map Word legend position to Excel LegendPosition. */
function mapLegendPosition(legend: string | undefined): LegendPosition | undefined {
  if (!legend || legend === "none") {
    return undefined;
  }
  const map: Record<string, LegendPosition> = {
    b: "b",
    l: "l",
    r: "r",
    t: "t",
    tr: "tr"
  };
  return map[legend] ?? "r";
}

/**
 * Convert a Word Chart definition to an Excel ChartModel.
 * This is the single source of truth for Word→Excel chart mapping.
 *
 * Used by:
 * - `renderWordChartSvg` (SVG rendering for HTML embedding)
 * - `createWordChartPdfRenderer` in `@pdf/excel-bridge` (PDF vector rendering)
 *
 * @param chart - The Word Chart definition.
 * @returns A fully constructed ChartModel suitable for Excel's chart renderer.
 */
export function wordChartToChartModel(chart: Chart): ChartModel {
  checkWordChart(chart, "WordExcel.toChartModel");
  return chartModelOf(chart);
}

/** A Word chart as `Build.chart` makes one: a known type and series of names, categories and finite values. */
function checkWordChart(chart: unknown, where: string): void {
  checkRecord(chart, "chart", where);
  const c = chart as { type?: unknown; series?: unknown };
  checkOneOf(c.type, CHART_TYPES, "chart.type", where);
  checkWordChartSeries(c.series, "chart.series", where);
}

function chartModelOf(chart: Chart): ChartModel {
  const mapping = mapWordChartType(chart.type);
  const kept = chart._kept;
  const chartTypeGroup = buildChartTypeGroupFromWord(chart, mapping);
  const axes = buildAxesForType(mapping, chart);
  const legend =
    chart.legend && chart.legend !== "none"
      ? {
          legendPos: mapLegendPosition(chart.legend),
          ...keptShape("spPr", keptChild(kept?.legend, "spPr", {})?.xml),
          ...keptShape("txPr", keptChild(kept?.legend, "txPr", {})?.xml)
        }
      : undefined;
  const plotAreaSpPr = areaShape(
    keptChild(kept?.plotArea, "spPr", plotAreaKeys(chart))?.xml,
    chart.plotAreaColor
  );
  const chartSpPr = areaShape(
    keptChild(kept?.space, "spPr", chartSpaceKeys(chart))?.xml,
    chart.chartAreaColor
  );
  const chartTxPr = keptChild(kept?.space, "txPr", {})?.xml;

  return {
    chart: {
      title: chart.title
        ? { text: { paragraphs: [{ runs: [{ text: chart.title }] }] } }
        : undefined,
      autoTitleDeleted: !chart.title,
      plotArea: {
        chartTypes: [chartTypeGroup],
        axes,
        ...(plotAreaSpPr !== undefined && { spPr: plotAreaSpPr })
      },
      legend,
      plotVisOnly: true
    },
    ...(chartSpPr !== undefined && { spPr: chartSpPr }),
    ...(chartTxPr !== undefined && { txPr: { _rawXml: chartTxPr } })
  };
}

/** A kept element as the Excel model holds one — its markup, which the renderer reads. */
function keptShape<K extends string>(
  key: K,
  markup: string | undefined
): { [P in K]?: { _rawXml: string } } {
  return (markup === undefined ? {} : { [key]: { _rawXml: markup } }) as {
    [P in K]?: { _rawXml: string };
  };
}

/** The plot or chart area's fill: the kept `c:spPr` while its colour is what was read, else the colour. */
function areaShape(
  kept: string | undefined,
  color: string | undefined
): ShapeProperties | undefined {
  if (kept !== undefined) {
    return { _rawXml: kept };
  }
  return color ? { fill: { solid: { srgb: color } } } : undefined;
}

/**
 * A Word series' formatting as the Excel model holds it, chosen as the Word
 * writer chooses it: the `c:spPr` and each `c:dPt` the reader kept while
 * `color` and the point's `pointColors` entry are what was read from them,
 * else those fields — a line series' colour strokes its line as well, as the
 * writer writes it.
 */
function seriesFormatting(
  chart: Chart,
  series: WordChartSeries,
  index: number,
  lineLike: boolean
): { spPr?: ShapeProperties; dataPoints?: DataPoint[] } {
  const keys = seriesKeys(series, index, series.chartType ?? chart.type);
  const keptSpPr = keptChild(series._kept?.node, "spPr", keys)?.xml;
  const color = series.color;
  const spPr: ShapeProperties | undefined =
    keptSpPr !== undefined
      ? { _rawXml: keptSpPr }
      : color
        ? {
            fill: { solid: { srgb: color } },
            ...(lineLike && { line: { width: 28575, color: { srgb: color } } })
          }
        : undefined;
  const dataPoints: DataPoint[] = [];
  const kept = new Set<number>();
  for (const point of keptPoints(series)) {
    kept.add(point.index!);
    if (point.spPr !== undefined) {
      dataPoints.push({ index: point.index!, spPr: { _rawXml: point.spPr } });
    }
  }
  series.pointColors?.forEach((pointColor, pointIndex) => {
    if (pointColor !== undefined && !kept.has(pointIndex)) {
      dataPoints.push({ index: pointIndex, spPr: { fill: { solid: { srgb: pointColor } } } });
    }
  });
  dataPoints.sort((a, b) => a.index - b.index);
  return {
    ...(spPr !== undefined && { spPr }),
    ...(dataPoints !== undefined && dataPoints.length > 0 && { dataPoints })
  };
}

/** An axis' kept gridlines, `c:spPr` and `c:txPr` as the Excel axis holds them. */
function axisFormatting(
  node: ChartKeptNode | undefined
): Pick<ExcelChartAxis, "majorGridlines" | "minorGridlines" | "spPr" | "txPr"> {
  const major = keptChild(node, "majorGridlines", {})?.xml;
  const minor = keptChild(node, "minorGridlines", {})?.xml;
  return {
    ...(major !== undefined && { majorGridlines: gridlineShape(major) }),
    ...(minor !== undefined && { minorGridlines: gridlineShape(minor) }),
    ...keptShape("spPr", keptChild(node, "spPr", {})?.xml),
    ...keptShape("txPr", keptChild(node, "txPr", {})?.xml)
  };
}

/** A gridlines element's `c:spPr` (its only child) — `{}`, Office's default line, when it has none. */
function gridlineShape(element: string): ShapeProperties {
  const inner = /^<[^>]*[^/]>([\s\S]*)<\/[^>]+>$/.exec(element)?.[1];
  return inner ? { _rawXml: inner } : {};
}

/** Build the chart type group from Word series data. */
function buildChartTypeGroupFromWord(chart: Chart, mapping: ExcelChartMapping): ChartTypeGroup {
  const { type } = mapping;
  const lineLike = type === "line" || type === "line3D" || type === "scatter";

  switch (type) {
    case "bar":
    case "bar3D": {
      const series: BarSeries[] = chart.series.map((s, idx) => ({
        index: idx,
        order: idx,
        tx: { value: s.name },
        ...seriesFormatting(chart, s, idx, lineLike),
        cat: makeLiteralCat(s.categories),
        val: makeLiteralVal(s.values)
      }));
      return {
        type,
        barDir: mapping.barDir ?? "col",
        grouping: (mapping.grouping as BarGrouping) ?? "clustered",
        varyColors: chart.series.length === 1,
        series,
        axisIds: [1, 2]
      } satisfies BarChartGroup;
    }

    case "line":
    case "line3D": {
      const series: LineSeries[] = chart.series.map((s, idx) => ({
        index: idx,
        order: idx,
        tx: { value: s.name },
        ...seriesFormatting(chart, s, idx, lineLike),
        cat: makeLiteralCat(s.categories),
        val: makeLiteralVal(s.values)
      }));
      return {
        type,
        grouping: (mapping.grouping as LineGrouping) ?? "standard",
        varyColors: false,
        series,
        axisIds: [1, 2]
      } satisfies LineChartGroup;
    }

    case "pie":
    case "pie3D": {
      const series: PieSeries[] = chart.series.map((s, idx) => ({
        index: idx,
        order: idx,
        tx: { value: s.name },
        ...seriesFormatting(chart, s, idx, lineLike),
        cat: makeLiteralCat(s.categories),
        val: makeLiteralVal(s.values)
      }));
      return {
        type,
        varyColors: true,
        series
      } satisfies PieChartGroup;
    }

    case "doughnut": {
      const series: PieSeries[] = chart.series.map((s, idx) => ({
        index: idx,
        order: idx,
        tx: { value: s.name },
        ...seriesFormatting(chart, s, idx, lineLike),
        cat: makeLiteralCat(s.categories),
        val: makeLiteralVal(s.values)
      }));
      return {
        type: "doughnut",
        varyColors: true,
        series,
        holeSize: 50
      } satisfies DoughnutChartGroup;
    }

    case "area":
    case "area3D": {
      const series: AreaSeries[] = chart.series.map((s, idx) => ({
        index: idx,
        order: idx,
        tx: { value: s.name },
        ...seriesFormatting(chart, s, idx, lineLike),
        cat: makeLiteralCat(s.categories),
        val: makeLiteralVal(s.values)
      }));
      return {
        type,
        grouping: (mapping.grouping as LineGrouping) ?? "standard",
        varyColors: false,
        series,
        axisIds: [1, 2]
      } satisfies AreaChartGroup;
    }

    case "scatter": {
      const series: ScatterSeries[] = chart.series.map((s, idx) => ({
        index: idx,
        order: idx,
        tx: { value: s.name },
        ...seriesFormatting(chart, s, idx, lineLike),
        xVal: makeLiteralCat(s.categories),
        yVal: makeLiteralVal(s.values)
      }));
      return {
        type: "scatter",
        scatterStyle: mapping.scatterStyle ?? "lineMarker",
        varyColors: false,
        series,
        axisIds: [1, 2]
      } satisfies ScatterChartGroup;
    }

    case "radar": {
      const series: RadarSeries[] = chart.series.map((s, idx) => ({
        index: idx,
        order: idx,
        tx: { value: s.name },
        ...seriesFormatting(chart, s, idx, lineLike),
        cat: makeLiteralCat(s.categories),
        val: makeLiteralVal(s.values)
      }));
      return {
        type: "radar",
        radarStyle: mapping.radarStyle ?? "standard",
        varyColors: false,
        series,
        axisIds: [1, 2]
      } satisfies RadarChartGroup;
    }

    case "surface":
    case "surface3D": {
      const series: SurfaceSeries[] = chart.series.map((s, idx) => ({
        index: idx,
        order: idx,
        tx: { value: s.name },
        ...seriesFormatting(chart, s, idx, lineLike),
        cat: makeLiteralCat(s.categories),
        val: makeLiteralVal(s.values)
      }));
      return {
        type,
        wireframe: mapping.wireframe,
        series,
        axisIds: [1, 2, 3]
      } satisfies SurfaceChartGroup;
    }

    case "bubble": {
      // Bubble charts need xVal, yVal, bubbleSize — use categories as x labels
      const series: ScatterSeries[] = chart.series.map((s, idx) => ({
        index: idx,
        order: idx,
        tx: { value: s.name },
        ...seriesFormatting(chart, s, idx, lineLike),
        xVal: makeLiteralCat(s.categories),
        yVal: makeLiteralVal(s.values)
      }));
      return {
        type: "scatter",
        scatterStyle: "lineMarker",
        varyColors: false,
        series,
        axisIds: [1, 2]
      } satisfies ScatterChartGroup;
    }

    case "stock": {
      const series: LineSeries[] = chart.series.map((s, idx) => ({
        index: idx,
        order: idx,
        tx: { value: s.name },
        ...seriesFormatting(chart, s, idx, lineLike),
        cat: makeLiteralCat(s.categories),
        val: makeLiteralVal(s.values)
      }));
      return {
        type: "line",
        grouping: "standard",
        varyColors: false,
        series,
        axisIds: [1, 2]
      } satisfies LineChartGroup;
    }

    default: {
      // Fallback to bar chart
      const series: BarSeries[] = chart.series.map((s, idx) => ({
        index: idx,
        order: idx,
        tx: { value: s.name },
        ...seriesFormatting(chart, s, idx, lineLike),
        cat: makeLiteralCat(s.categories),
        val: makeLiteralVal(s.values)
      }));
      return {
        type: "bar",
        barDir: "col",
        grouping: "clustered",
        varyColors: false,
        series,
        axisIds: [1, 2]
      } satisfies BarChartGroup;
    }
  }
}

/** Build axes appropriate for the chart type. */
function buildAxesForType(mapping: ExcelChartMapping, chart: Chart): ExcelChartAxis[] {
  const { type } = mapping;

  // Pie/doughnut charts have no axes
  if (type === "pie" || type === "pie3D" || type === "doughnut") {
    return [];
  }

  const catAxis: ExcelChartAxis = {
    axisType: "cat",
    axId: 1,
    axPos: "b",
    crossAx: 2,
    scaling: {},
    delete: chart.categoryAxis?.hidden,
    ...axisFormatting(chart._kept?.axes.find(a => a.role === "category")?.node)
  };

  const valAxis: ExcelChartAxis = {
    axisType: "val",
    axId: 2,
    axPos: "l",
    crossAx: 1,
    scaling: {
      min: chart.valueAxis?.min,
      max: chart.valueAxis?.max
    },
    majorUnit: chart.valueAxis?.majorUnit,
    delete: chart.valueAxis?.hidden,
    numFmt: chart.valueAxis?.numberFormat
      ? { formatCode: chart.valueAxis.numberFormat, sourceLinked: false }
      : undefined,
    ...axisFormatting(chart._kept?.axes.find(a => a.role === "value")?.node)
  };

  const axes: ExcelChartAxis[] = [catAxis, valAxis];

  // Surface charts need a series axis
  if (type === "surface" || type === "surface3D") {
    axes.push({
      axisType: "ser",
      axId: 3,
      axPos: "b",
      crossAx: 2,
      scaling: {}
    });
  }

  return axes;
}

/**
 * Extract tables from a Word document and convert them to Excel worksheet data.
 * Each table becomes a 2D array of cell values.
 *
 * @param doc - The DOCX document model.
 * @returns Array of table data objects, each with a name and 2D data array.
 */
export function extractTablesToExcel(
  doc: DocxDocument
): { name: string; data: (string | number | null)[][] }[] {
  checkHandle(doc, "document", "doc", "WordExcel.extractTables");
  const results: { name: string; data: (string | number | null)[][] }[] = [];
  let tableIndex = 0;

  for (const block of doc.body) {
    if ("type" in block && block.type === "table") {
      tableIndex++;
      const table = block as Table;
      const data = extractTableData(table);
      results.push({ name: `Table ${tableIndex}`, data });
    }
  }

  return results;
}

/**
 * Extract a single table into a 2D array, handling gridSpan.
 */
function extractTableData(table: Table): (string | number | null)[][] {
  const data: (string | number | null)[][] = [];

  for (const row of table.rows) {
    const rowData: (string | number | null)[] = [];

    for (const cell of row.cells) {
      const span = cell.properties?.gridSpan ?? 1;
      const text = extractCellPlainText(cell);
      const value = parseNumericValue(text);

      // First position gets the value
      rowData.push(value);

      // Remaining spanned cells get null
      for (let s = 1; s < span; s++) {
        rowData.push(null);
      }
    }

    data.push(rowData);
  }

  return data;
}

/**
 * Extract plain text from a table cell by traversing its content.
 */
function extractCellPlainText(cell: TableCell): string {
  const parts: string[] = [];

  for (const item of cell.content) {
    if ("type" in item && item.type === "paragraph") {
      const para = item as Paragraph;
      parts.push(extractParagraphText(para));
    } else if ("type" in item && item.type === "table") {
      // Nested table: extract text from each cell separated by tabs
      const nested = item as Table;
      for (const nRow of nested.rows) {
        const cellTexts: string[] = [];
        for (const nCell of nRow.cells) {
          cellTexts.push(extractCellPlainText(nCell));
        }
        parts.push(cellTexts.join("\t"));
      }
    }
  }

  return parts.join("\n").trim();
}

/**
 * Parse a string value as a number if it represents a numeric value.
 * Returns the number if parseable, otherwise the original string, or null if empty.
 */
function parseNumericValue(text: string): string | number | null {
  if (!text) {
    return null;
  }

  const trimmed = text.trim();
  if (!trimmed) {
    return null;
  }

  // Try to parse as number
  const num = Number(trimmed);
  if (!Number.isNaN(num) && trimmed !== "") {
    return num;
  }

  return text;
}
