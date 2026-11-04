/**
 * DOCX Module - Template Chart Data Binding
 *
 * Extends the template engine with chart data binding support.
 * Allows replacing series data, categories, and titles in chart
 * definitions within a DocxDocument.
 *
 * @stability experimental
 */

import {
  checkArray,
  checkFiniteArray,
  checkHandle,
  checkRecord,
  checkRgbColor,
  checkString,
  checkStringArray,
  invalidArgument
} from "@word/core/argument-checks";
import type {
  DocxDocument,
  BodyContent,
  ChartContent,
  ChartSeries,
  Chart,
  ChartType,
  HexColor
} from "@word/types";

// =============================================================================
// Public Types
// =============================================================================

/** Data for a single chart series to bind. */
export interface ChartSeriesData {
  /** Series name (legend label). */
  readonly name: string;
  /** Numeric data values. */
  readonly values: readonly number[];
  /** Optional series color (6-digit hex RGB). */
  readonly color?: string;
}

/** Binding specification mapping a chart reference to new data. */
export interface ChartBinding {
  /** Chart index (0-based) or chart name. */
  readonly chartRef: number | string;
  /** New series data. */
  readonly series: readonly ChartSeriesData[];
  /** New category labels (X-axis). */
  readonly categories?: readonly string[];
  /** Chart title override. */
  readonly title?: string;
}

/** Aggregate template data for chart bindings. */
export interface ChartTemplateData {
  /** Array of chart bindings to apply. */
  readonly charts: readonly ChartBinding[];
}

// =============================================================================
// Supported chart type set (for validation)
// =============================================================================

const BASIC_CHART_TYPES: ReadonlySet<string> = new Set<ChartType>([
  "bar",
  "barStacked",
  "barPercentStacked",
  "column",
  "columnStacked",
  "columnPercentStacked",
  "line",
  "lineStacked",
  "lineMarked",
  "pie",
  "pie3D",
  "doughnut"
]);

// =============================================================================
// Core: bindChartData
// =============================================================================

/**
 * Replace chart data in a DocxDocument according to the provided bindings.
 *
 * Iterates through the document body, locates `ChartContent` blocks, and
 * replaces their series data, categories, and titles with the values
 * specified in each binding.
 *
 * Charts are matched by either:
 * - Numeric index (0-based position among all ChartContent blocks in body)
 * - String name (matched against `ChartContent.name`)
 *
 * Only basic chart types (bar, column, line, pie and their variants) are
 * supported. Bindings that reference unsupported chart types are skipped.
 *
 * A binding whose `chartRef` names no chart in the body — an index past the
 * last chart, or a name no chart has — is skipped, so one set of bindings can
 * be applied to documents that lack some of the charts. To tell which were
 * applied, compare the result's charts with the bindings.
 *
 * Every binding is checked before any is applied: `chartRef` must be an
 * integer ≥ 0 or a non-empty string, and values must be finite numbers, one
 * per category when the binding gives `categories`.
 *
 * Does not modify `doc`; the result may share unchanged nodes with it — see
 * `Document.clone`.
 *
 * @param doc - The DocxDocument to modify.
 * @param chartBindings - Array of chart bindings to apply.
 * @returns A new DocxDocument with chart data replaced.
 * @throws {DocxArgumentError} If a binding is malformed (including a
 * `chartRef` that is neither an integer ≥ 0 nor a non-empty string), a value
 * is not a finite number, or a series' values do not match the categories.
 *
 * @stability experimental
 */
export function bindChartData(
  doc: DocxDocument,
  chartBindings: readonly ChartBinding[]
): DocxDocument {
  checkHandle(doc, "document", "doc", "Template.bindChartData");
  const where = "Template.bindChartData";
  checkChartBindings(chartBindings, where);
  if (chartBindings.length === 0) {
    return doc;
  }

  // Build a map of chart index -> binding and name -> binding for fast lookup
  const indexBindings = new Map<number, ChartBinding>();
  const nameBindings = new Map<string, ChartBinding>();

  for (const binding of chartBindings) {
    if (typeof binding.chartRef === "number") {
      indexBindings.set(binding.chartRef, binding);
    } else {
      nameBindings.set(binding.chartRef, binding);
    }
  }

  // Walk through body content, track chart index, and apply bindings
  let chartIndex = 0;
  const newBody: BodyContent[] = [];

  for (const block of doc.body) {
    if (block.type === "chart") {
      const chartContent = block as ChartContent;
      const binding = resolveBinding(chartContent, chartIndex, indexBindings, nameBindings);

      if (binding && isSupportedChartType(chartContent.chart.type)) {
        newBody.push(applyBinding(chartContent, binding));
      } else {
        newBody.push(block);
      }
      chartIndex++;
    } else {
      newBody.push(block);
    }
  }

  return { ...doc, body: newBody };
}

/** The bindings' shape and values, before any is applied. */
function checkChartBindings(chartBindings: readonly ChartBinding[], where: string): void {
  checkArray(chartBindings, "chartBindings", where);
  chartBindings.forEach((binding, i) => {
    const at = `chartBindings[${i}]`;
    checkRecord(binding, at, where);
    const ref = binding.chartRef;
    if (
      !(typeof ref === "string" && ref !== "") &&
      !(Number.isInteger(ref) && (ref as number) >= 0)
    ) {
      invalidArgument(
        where,
        `${at}.chartRef`,
        "a chart index (an integer ≥ 0) or a chart name",
        ref
      );
    }
    checkStringArray(binding.categories, `${at}.categories`, where, { optional: true });
    checkString(binding.title, `${at}.title`, where, { optional: true });
    checkArray(binding.series, `${at}.series`, where);
    binding.series.forEach((series, j) => {
      const sAt = `${at}.series[${j}]`;
      checkRecord(series, sAt, where);
      checkString(series.name, `${sAt}.name`, where);
      checkFiniteArray(series.values, `${sAt}.values`, where);
      checkRgbColor(series.color, `${sAt}.color`, where, { optional: true });
      if (binding.categories && binding.categories.length > 0) {
        checkSeriesLength(series, binding.categories.length, sAt, where);
      }
    });
  });
}

function checkSeriesLength(
  series: ChartSeriesData,
  categories: number,
  field: string,
  where: string
): void {
  if (series.values.length !== categories) {
    invalidArgument(
      where,
      `${field}.values`,
      `${categories} value(s), one per category`,
      series.values
    );
  }
}

// =============================================================================
// Internal Helpers
// =============================================================================

/**
 * Resolve which binding (if any) applies to a given chart content block.
 */
function resolveBinding(
  chartContent: ChartContent,
  index: number,
  indexBindings: ReadonlyMap<number, ChartBinding>,
  nameBindings: ReadonlyMap<string, ChartBinding>
): ChartBinding | undefined {
  // Index-based binding takes priority
  const byIndex = indexBindings.get(index);
  if (byIndex) {
    return byIndex;
  }

  // Try name-based binding
  if (chartContent.name) {
    const byName = nameBindings.get(chartContent.name);
    if (byName) {
      return byName;
    }
  }

  return undefined;
}

/**
 * Check whether the chart type is in the supported set.
 */
function isSupportedChartType(chartType: ChartType): boolean {
  return BASIC_CHART_TYPES.has(chartType);
}

/**
 * Apply a binding to a ChartContent block, producing a new ChartContent
 * with updated series, categories, and title.
 */
function applyBinding(chartContent: ChartContent, binding: ChartBinding): ChartContent {
  const originalChart = chartContent.chart;
  const categories = binding.categories ?? extractOriginalCategories(originalChart);

  const newSeries: ChartSeries[] = [];
  for (const seriesData of binding.series) {
    newSeries.push(buildChartSeries(seriesData, categories));
  }

  const newChart: Chart = {
    ...originalChart,
    series: newSeries,
    title: binding.title !== undefined ? binding.title : originalChart.title
  };

  return {
    ...chartContent,
    chart: newChart
  };
}

/**
 * Extract categories from the original chart's first series (fallback).
 */
function extractOriginalCategories(chart: Chart): string[] {
  if (chart.series.length > 0) {
    return [...chart.series[0].categories];
  }
  return [];
}

/**
 * Build a ChartSeries from binding data and resolved categories.
 */
function buildChartSeries(data: ChartSeriesData, categories: readonly string[]): ChartSeries {
  const series: ChartSeries = {
    name: data.name,
    categories: categories,
    values: data.values
  };

  if (data.color) {
    return { ...series, color: data.color as HexColor };
  }

  return series;
}
