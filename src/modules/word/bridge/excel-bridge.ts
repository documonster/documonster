/**
 * Excel-to-Word Bridge
 *
 * Converts an Excel Workbook into a DocxDocument model.
 * This is the ONLY file in the Word module that imports from @excel.
 *
 * @example
 * ```typescript
 * import { Workbook } from "excelts";
 * import { excelToDocx } from "excelts/word/excel";
 *
 * const wb = new Workbook();
 * await wb.xlsx.load(buffer);
 * const doc = excelToDocx(wb);
 * ```
 */

import type { Cell as ExcelCell } from "@excel/cell";
import { buildChartExModel } from "@excel/chart/chart-ex-builder";
import { renderChartEx } from "@excel/chart/chart-ex-renderer";
import type { AddChartExOptions, ChartExModel, ChartExType } from "@excel/chart/chart-ex-types";
import { renderChartSvg } from "@excel/chart/chart-renderer";
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
  ChartAxis as ExcelChartAxis
} from "@excel/chart/types";
import { ValueType } from "@excel/enums";
import type {
  Font as ExcelFont,
  Fill as ExcelFill,
  Borders as ExcelBorders,
  Border as ExcelBorder,
  Alignment as ExcelAlignment,
  Color as ExcelColor
} from "@excel/types";
// Use the browser base class so the public `excelToDocx(workbook)` signature
// is callable from both the Node entry (where `Workbook` is the Node subclass
// — trivially assignable to the base) and the browser entry (where `Workbook`
// is already the base). Importing the Node alias `@excel/workbook` would force
// browser consumers to satisfy `xlsx.readFile`/`writeFile`, which the browser
// XLSX surface intentionally omits — see issue #160.
import type { Workbook } from "@excel/workbook.browser";
import type { Worksheet } from "@excel/worksheet";

import { type Mutable } from "../core/internal-utils";
import { extractParagraphText } from "../core/text-utils";
import type {
  Alignment,
  Chart,
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
} from "../types";
import { EMU_PER_INCH } from "../units";

// =============================================================================
// Public API
// =============================================================================

/** Options for Excel → DOCX conversion. */
export interface ExcelToDocxOptions {
  /** Which sheets to include (by name or 0-based index). Default: all visible sheets. */
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

/**
 * Convert an Excel Workbook to a DocxDocument model.
 *
 * Creates a Word document with tables representing each worksheet's data.
 * Preserves cell formatting (fonts, colors, alignment) as run/paragraph properties.
 *
 * @param workbook - The Excel Workbook to convert.
 * @param options - Conversion options.
 * @returns A DocxDocument model ready for packaging, PDF conversion, or Markdown output.
 */
export function excelToDocx(workbook: Workbook, options?: ExcelToDocxOptions): DocxDocument {
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
      body.push(heading(ws.name, opts.sheetHeadingLevel));
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

function selectSheets(workbook: Workbook, filter: readonly (string | number)[]): Worksheet[] {
  const all = workbook.worksheets.filter(ws => ws && ws.state !== "veryHidden");

  if (filter.length === 0) {
    return all.filter(ws => ws.state !== "hidden");
  }

  const result: Worksheet[] = [];
  for (const spec of filter) {
    if (typeof spec === "number") {
      const ws = all[spec];
      if (ws) {
        result.push(ws);
      }
    } else {
      const ws = all.find(s => s.name === spec);
      if (ws) {
        result.push(ws);
      }
    }
  }
  return result;
}

// =============================================================================
// Worksheet → Table
// =============================================================================

function sheetToTable(ws: Worksheet, opts: Required<ExcelToDocxOptions>): Table | null {
  const rowCount = Math.min(ws.rowCount, opts.maxRows);
  if (rowCount === 0) {
    return null;
  }

  // Find actual column range
  let maxCol = 0;
  for (let r = 1; r <= rowCount; r++) {
    const row = ws.getRow(r);
    if (row.cellCount > maxCol) {
      maxCol = row.cellCount;
    }
  }
  maxCol = Math.min(maxCol, opts.maxColumns);
  if (maxCol === 0) {
    return null;
  }

  const rows: TableRow[] = [];

  for (let r = 1; r <= rowCount; r++) {
    const row = ws.getRow(r);

    // Skip completely empty rows
    let hasContent = false;
    for (let c = 1; c <= maxCol; c++) {
      const cell = row.getCell(c);
      if (cell.type !== ValueType.Null && cell.type !== ValueType.Merge) {
        hasContent = true;
        break;
      }
    }
    if (!hasContent) {
      continue;
    }

    const cells: TableCell[] = [];
    for (let c = 1; c <= maxCol; c++) {
      const cell = row.getCell(c);
      cells.push(convertCell(cell, opts));
    }
    rows.push({ cells });
  }

  if (rows.length === 0) {
    return null;
  }

  // Column widths from worksheet definitions (Excel char width → twips)
  const columnWidths: number[] = [];
  for (let c = 1; c <= maxCol; c++) {
    const col = ws.getColumn(c);
    const charWidth = col.width ?? 10;
    columnWidths.push(Math.round(charWidth * 140));
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
  const text = cellText(cell);
  const children: Run[] = [];

  if (text) {
    if (cell.type === ValueType.RichText) {
      const value = cell.value as {
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
      const runProps = opts.preserveFormatting ? fontToRun(cell.font) : undefined;
      children.push({
        properties: runProps,
        content: [{ type: "text", text }]
      });
    }
  }

  const para: Paragraph = {
    type: "paragraph",
    properties: opts.preserveFormatting ? alignmentToParaProps(cell.alignment) : undefined,
    children: wrapHyperlink(cell.hyperlink, children)
  };

  const cellProps: Mutable<TableCellProperties> = {};

  if (opts.preserveFormatting && cell.fill) {
    const shading = fillToShading(cell.fill);
    if (shading) {
      cellProps.shading = shading;
    }
  }

  if (opts.includeBorders && cell.border) {
    const borders = bordersToCellBorders(cell.border);
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

function cellText(cell: ExcelCell): string {
  if (cell.type === ValueType.Null || cell.type === ValueType.Merge) {
    return "";
  }

  if (cell.text !== undefined && cell.text !== null) {
    return String(cell.text);
  }

  const value = cell.value;
  if (value === null || value === undefined) {
    return "";
  }

  if (typeof value === "object") {
    if ("result" in value) {
      const result = value.result;
      if (result === null || result === undefined) {
        return "";
      }
      if (result instanceof Date) {
        return fmtDate(result);
      }
      return String(result);
    }
    if ("richText" in value) {
      return (value.richText as ReadonlyArray<{ text?: string }>).map(s => s.text ?? "").join("");
    }
    if ("text" in value) {
      return String(value.text);
    }
    if ("checkbox" in value) {
      return value.checkbox ? "\u2611" : "\u2610";
    }
    if ("error" in value) {
      return String(value.error);
    }
  }

  if (value instanceof Date) {
    return fmtDate(value);
  }

  return String(value);
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
    const defaults: Record<number, string> = {
      0: "FFFFFF",
      1: "000000",
      2: "44546A",
      3: "E7E6E6",
      4: "4472C4",
      5: "ED7D31",
      6: "A5A5A5",
      7: "FFC000",
      8: "5B9BD5",
      9: "70AD47"
    };
    return defaults[color.theme] ?? "000000";
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
export interface WordChartExOptions {
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

/**
 * Build a ChartEx XML part for embedding in a Word document.
 *
 * This bridges the Excel ChartEx renderer into the Word module,
 * generating a complete `cx:chartSpace` XML string from high-level options.
 *
 * @param options - Chart configuration.
 * @returns The complete ChartEx XML string for the chart part.
 */
export function buildWordChartExXml(options: WordChartExOptions): string {
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

/**
 * Render a Word Chart to SVG string for HTML embedding.
 *
 * Bridges Word's Chart definition to Excel's chart renderer by constructing
 * a ChartModel with literal (inline) data and delegating to renderChartSvg.
 *
 * @param chart - The Word Chart definition.
 * @returns SVG markup string.
 */
export function renderWordChartSvg(chart: Chart): string {
  const model = wordChartToChartModel(chart);
  return renderChartSvg(model, {
    width: chart.width ? Math.round((chart.width / EMU_PER_INCH) * 96) : undefined,
    height: chart.height ? Math.round((chart.height / EMU_PER_INCH) * 96) : undefined
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
  const { Workbook } = await import("@excel/workbook");
  const wb = new Workbook();
  const ws = wb.addWorksheet("Sheet1");

  // Row 1: headers — A1 empty, B1..N1 = series names
  for (let c = 0; c < series.length; c++) {
    ws.getCell(1, c + 2).value = series[c].name;
  }

  // Rows 2+: categories and values
  const categories = series.length > 0 ? series[0].categories : [];
  for (let r = 0; r < categories.length; r++) {
    ws.getCell(r + 2, 1).value = categories[r];
    for (let c = 0; c < series.length; c++) {
      const val = r < series[c].values.length ? series[c].values[r] : 0;
      ws.getCell(r + 2, c + 2).value = val;
    }
  }

  return new Uint8Array(await wb.xlsx.writeBuffer());
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
  const mapping = mapWordChartType(chart.type);
  const chartTypeGroup = buildChartTypeGroupFromWord(chart, mapping);
  const axes = buildAxesForType(mapping, chart);
  const legend =
    chart.legend && chart.legend !== "none"
      ? { legendPos: mapLegendPosition(chart.legend) }
      : undefined;

  return {
    chart: {
      title: chart.title
        ? { text: { paragraphs: [{ runs: [{ text: chart.title }] }] } }
        : undefined,
      autoTitleDeleted: !chart.title,
      plotArea: {
        chartTypes: [chartTypeGroup],
        axes
      },
      legend,
      plotVisOnly: true
    }
  };
}

/** Build the chart type group from Word series data. */
function buildChartTypeGroupFromWord(chart: Chart, mapping: ExcelChartMapping): ChartTypeGroup {
  const { type } = mapping;

  switch (type) {
    case "bar":
    case "bar3D": {
      const series: BarSeries[] = chart.series.map((s, idx) => ({
        index: idx,
        order: idx,
        tx: { value: s.name },
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
    delete: chart.categoryAxis?.hidden
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
      : undefined
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
