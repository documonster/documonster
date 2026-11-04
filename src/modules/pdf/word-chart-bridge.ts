/**
 * Word-chart → PDF rendering bridge.
 *
 * Word documents can embed two chart families:
 *
 *   - Classic `<c:chart>`   → Word `Chart` model
 *   - Modern `<cx:chartSpace>` ChartEx → carried as raw XML on the chart
 *
 * Both are rendered to PDF using **Excel's** vector chart engine
 * (`@excel/chart/render/*`), reusing the 8000+ lines of drawing logic
 * rather than duplicating it for Word. This is the single seam where a
 * Word chart definition meets the Excel rendering engine, so it is the
 * ONLY file in the PDF module that legitimately imports from BOTH
 * `@word` and `@excel`. (Registered as an explicit layering exception in
 * AGENTS.md, alongside `excel-bridge.ts` and `word-bridge.ts`.)
 *
 * The renderers are loaded by `word-bridge.ts` via dynamic `import()`,
 * so a consumer who never converts a DOCX containing charts never
 * bundles the Excel chart engine — keeping the core PDF path
 * tree-shakeable.
 */

import { canRenderChartExAsVectorPdf, drawChartExPdf } from "@excel/chart/render/chart-ex-renderer";
import { drawChartPdf } from "@excel/chart/render/chart-renderer";
import { parseChartEx } from "@excel/chart/serialize/chart-ex-parser";
import type { ChartThemeColors } from "@excel/chart/shared/chart-utils";
import type { PdfPageBuilder } from "@pdf/builder/document-builder";
import { wordChartToChartModel } from "@word/bridge/excel-bridge";
import type { LayoutChart } from "@word/layout/layout-model";
import type {
  Chart as WordChart,
  ChartContent as WordChartContent,
  ChartExContent as WordChartExContent
} from "@word/types";

// Re-export the Word chart type used in the renderer's public signature so
// the `Pdf` surface can type its lazy `wordChartRenderer` wrapper without
// importing from `@word` directly — only bridge files may cross into `@word`.
export type { Chart as WordChart } from "@word/types";

/**
 * The classic-chart renderer callback behind the deprecated `Pdf.wordChartRenderer()`.
 *
 * It draws a Word chart's classic model with the Excel chart engine — exactly what
 * `Pdf.fromDocx` does by default through {@link createWordLayoutChartPdfRenderer}, which
 * also covers ChartEx. Kept so the deprecated name returns the callback it always did.
 */
export function createWordChartPdfRenderer(): (
  chart: WordChart,
  page: PdfPageBuilder,
  rect: { x: number; y: number; width: number; height: number }
) => void {
  return (chart, page, rect) => {
    const model = wordChartToChartModel(chart);
    drawChartPdf(page, model, rect);
  };
}

/**
 * Create a layout-aware chart renderer for use as the internal
 * `RenderLayoutOptions.chartRenderer` of the Word→PDF bridge.
 *
 * Unlike {@link createWordChartPdfRenderer} (which only sees the inner
 * classic `Chart` model), this renderer receives the full
 * {@link LayoutChart} and therefore handles **both** chart families
 * with the full Excel rendering engine:
 *
 * - Classic `<c:chart>` (`chartKind === "chart"`) → `wordChartToChartModel`
 *   → `drawChartPdf` (vector).
 * - Modern `<cx:chartSpace>` ChartEx (`chartKind === "chartEx"`,
 *   e.g. sunburst / treemap / waterfall / funnel / boxWhisker /
 *   histogram / pareto / regionMap) → `parseChartEx` → `drawChartExPdf`
 *   (vector) when the layout is vector-capable, otherwise the
 *   pre-rendered SVG carried on the `LayoutChart` is left for the
 *   translator's fallback.
 *
 * Returns `false` to decline a chart so the translator's built-in
 * fallback (inline SVG, then a titled placeholder box) takes over. This
 * keeps "fail soft" behaviour: a chart the engine can't draw still
 * renders *something* rather than a blank slot.
 */
export function createWordLayoutChartPdfRenderer(
  options: { readonly themeColors?: ChartThemeColors } = {}
): (
  chart: LayoutChart,
  page: PdfPageBuilder,
  rect: { x: number; y: number; width: number; height: number }
) => boolean | void {
  // The document theme's colours: a Word chart's automatic and scheme colours are its theme's.
  const theme = options.themeColors === undefined ? {} : { themeColors: options.themeColors };
  return (layoutChart, page, rect) => {
    const source = layoutChart.source as WordChartContent | WordChartExContent | undefined;

    if (layoutChart.chartKind === "chart") {
      // Classic chart: prefer the structured source, fall back to nothing.
      if (source && source.type === "chart") {
        drawChartPdf(page, wordChartToChartModel(source.chart), { ...rect, ...theme });
        return;
      }
      return false;
    }

    // ChartEx. Parse the carried `cx:chartSpace` XML into a ChartExModel
    // and render it as vector PDF when the layout IDs are supported.
    if (source && source.type === "chartEx" && source.chartExXml) {
      let model;
      try {
        model = parseChartEx(source.chartExXml);
      } catch {
        return false; // Malformed XML — let the fallback path handle it.
      }
      if (model && canRenderChartExAsVectorPdf(model)) {
        drawChartExPdf(page, model, rect, {
          title: layoutChart.title,
          ...theme
        });
        return;
      }
    }
    // Not vector-capable (or no source): decline so the translator
    // falls back to the inline SVG / placeholder.
    return false;
  };
}
