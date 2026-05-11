/**
 * DOCX Module - Template Chart Data Binding
 *
 * Extends the template engine with chart data binding support.
 * Allows replacing series data, categories, and titles in chart
 * definitions within a DocxDocument.
 *
 * @stability experimental
 */

import { DocxError } from "../errors";
import type {
  DocxDocument,
  BodyContent,
  ChartContent,
  ChartSeries,
  Chart,
  ChartType,
  HexColor
} from "../types";

// =============================================================================
// Public Types
// =============================================================================

/** Data for a single chart series to bind. */
export interface ChartSeriesData {
  /** Series name (legend label). */
  readonly name: string;
  /** Numeric data values. */
  readonly values: number[];
  /** Optional series color (6-digit hex RGB). */
  readonly color?: string;
}

/** Binding specification mapping a chart reference to new data. */
export interface ChartBinding {
  /** Chart index (0-based) or chart name. */
  readonly chartRef: number | string;
  /** New series data. */
  readonly series: ChartSeriesData[];
  /** New category labels (X-axis). */
  readonly categories?: string[];
  /** Chart title override. */
  readonly title?: string;
}

/** Aggregate template data for chart bindings. */
export interface ChartTemplateData {
  /** Array of chart bindings to apply. */
  readonly charts: ChartBinding[];
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
 * @param doc - The DocxDocument to modify.
 * @param chartBindings - Array of chart bindings to apply.
 * @returns A new DocxDocument with chart data replaced.
 *
 * @stability experimental
 */
export function bindChartData(doc: DocxDocument, chartBindings: ChartBinding[]): DocxDocument {
  if (chartBindings.length === 0) {
    return doc;
  }

  // Validate bindings: series values length must match categories length when both provided
  for (const binding of chartBindings) {
    if (binding.categories && binding.categories.length > 0) {
      for (const series of binding.series) {
        if (series.values.length !== binding.categories.length) {
          throw new DocxError(
            `Chart binding "${String(binding.chartRef)}": series "${series.name}" has ${series.values.length} values but ${binding.categories.length} categories were provided. Lengths must match.`
          );
        }
      }
    }
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
function buildChartSeries(data: ChartSeriesData, categories: string[]): ChartSeries {
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
