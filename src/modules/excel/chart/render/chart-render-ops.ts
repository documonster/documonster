/**
 * Chart render operations — `chartToSVG` / `chartToPNG`.
 *
 * These are split out of `chart-handle.ts` on purpose: they are the **only**
 * chart-handle operations that pull the heavy SVG/PNG renderers
 * (`chart-renderer.ts` ≈ 312 KB, `chart-ex-renderer.ts` ≈ 238 KB). Keeping
 * them here means the chart *creation / mutation* surface (`registerChart`,
 * `chartMutate`, …) — which `worksheet.ts` and therefore `Workbook` statically
 * reach — never drags the renderers into a bundle. Only a consumer that
 * actually calls `Chart.toSvg` / `Chart.toPng` pays for them.
 *
 * This mirrors the module's "execute on call, never at load" contract:
 * renderers stay behind a leaf module that nothing on the create path pulls in.
 */

import {
  _chartRefreshChartExCaches,
  chartChartExModel,
  chartChartModel
} from "@excel/chart/chart-handle";
import { renderChartExPng, renderChartExSvg } from "@excel/chart/render/chart-ex-renderer";
import type { ChartRenderOptions } from "@excel/chart/render/chart-renderer";
import { renderChartPng, renderChartSvg } from "@excel/chart/render/chart-renderer";
import { workbookThemeColors } from "@excel/chart/render/workbook-theme";
import { checkChart } from "@excel/core/feature-handles";
import type { ChartHandle } from "@excel/core/worksheet-core";
import { ExcelChartOptionsError } from "@excel/errors";
import { asPlainBytes } from "@utils/binary";

/**
 * Render this chart as a **zero-dependency deterministic preview** SVG.
 *
 * The output is suitable for thumbnails, email attachments, server-side report
 * generation, CI smoke tests, and README images. It is **not** an
 * Excel-pixel-perfect compositor — text layout, font metrics, and 3D
 * projection are approximated for a stable preview rather than reproduced from
 * Excel's internal renderer.
 *
 * For production-grade rendering (Excel-identical layout, real 3D for non-bar
 * types, exact font hinting), round-trip the `.xlsx` through headless
 * LibreOffice (`soffice --convert-to pdf`).
 *
 * See `src/modules/excel/README.md` → "Rendering scope" for the complete
 * boundary list.
 */
export function chartToSVG(c: ChartHandle, options: ChartRenderOptions = {}): string {
  checkChart(c);
  options = withWorkbookTheme(c, options);
  const model = chartChartModel(c);
  if (model) {
    return renderChartSvg(model, options, "Chart.toSvg");
  }
  const chartEx = chartChartExModel(c);
  if (chartEx) {
    _chartRefreshChartExCaches(c);
    return renderChartExSvg(chartEx, options, "Chart.toSvg");
  }
  throw new ExcelChartOptionsError("c", "Cannot render chart because no chart model is available.");
}

/**
 * Render this chart as a **zero-dependency deterministic preview** PNG.
 *
 * Browsers use a `<canvas>` pipeline; Node.js uses the built-in
 * `BasicRasterCanvas` rasteriser (a pure-JS SVG-subset rasteriser — no native
 * canvas dependency). DrawingML effect filters
 * (shadow/glow/soft-edge/blur/reflection) round-trip through XML and emit as
 * SVG `<filter>`, but the Node PNG rasteriser silently drops them; the browser
 * canvas path renders them natively.
 *
 * See {@link chartToSVG} for the full scope-boundary note. For pixel-perfect
 * output, convert through LibreOffice.
 */
export async function chartToPNG(
  c: ChartHandle,
  options: ChartRenderOptions = {}
): Promise<Uint8Array<ArrayBuffer>> {
  checkChart(c);
  options = withWorkbookTheme(c, options);
  // `async` makes the "no model" branch reject the returned promise instead of
  // throwing synchronously, honouring the `Promise<Uint8Array>` contract.
  const model = chartChartModel(c);
  if (model) {
    return asPlainBytes(await renderChartPng(model, options, "Chart.toPng"));
  }
  const chartEx = chartChartExModel(c);
  if (chartEx) {
    _chartRefreshChartExCaches(c);
    return asPlainBytes(await renderChartExPng(chartEx, options, "Chart.toPng"));
  }
  throw new ExcelChartOptionsError("c", "Cannot render chart because no chart model is available.");
}

/**
 * `options` with the chart's workbook theme as `themeColors`, so the chart's scheme colours
 * resolve as Excel resolves them; a `themeColors` the caller gives wins. Options that are not
 * an object are left for the renderer to refuse.
 */
function withWorkbookTheme(c: ChartHandle, options: ChartRenderOptions): ChartRenderOptions {
  if (options === null || typeof options !== "object" || options.themeColors !== undefined) {
    return options;
  }
  const workbook = (c.worksheet as { _workbook?: unknown } | undefined)?._workbook;
  const themeColors = workbookThemeColors(workbook as never);
  return themeColors === undefined ? options : { ...options, themeColors };
}
