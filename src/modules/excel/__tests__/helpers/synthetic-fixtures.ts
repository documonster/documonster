/**
 * Synthetic chart fixture builders.
 *
 * The chart test corpus historically contained a single real-Excel
 * workbook (`__tests__/data/chart-pivot-sample.xlsx`). The synthetic
 * fixtures in this file fill the rest of the matrix programmatically so
 * `chart-synthetic-corpus.integration.test.ts`,
 * `chart-mutation-roundtrip.integration.test.ts`, and
 * `chartsheet-roundtrip.integration.test.ts` can run on every CI without
 * checking proprietary `.xlsx` blobs into the repo.
 *
 * IMPORTANT: these are SYNTHETIC fixtures — the workbook bytes are
 * produced by excelts itself. They are NOT a substitute for true Office
 * round-trip oracle coverage. The fixtures are deliberately marked with
 * a non-conformant XML comment via the `syntheticMarker` helper so the
 * mutation tests can prove that loaded raw XML (including unknown
 * extensions and unsupported child elements) survives a high-level
 * mutation followed by `writeBuffer`.
 *
 * Each builder returns an `Uint8Array` of xlsx bytes plus a reusable
 * description for the `expect(...)` failure messages and audit traces.
 *
 * The builders intentionally exercise every classic chart type, every
 * ChartEx layoutId, and the cross-cutting features the round-trip
 * pipeline regresses on most often (combo + secondary axis, date axis,
 * log axis, pivot chart with field buttons, chartsheet hosting, raw
 * passthrough with extension XML).
 */

import { extractAll, type ExtractedFile } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import {
  applyChartPreset,
  EXCEL_CHART_PRESETS,
  type ChartExType,
  type ExcelChartPreset
} from "@excel/chart";
import { Workbook } from "@excel/workbook";

const textDecoder = new TextDecoder();
const textEncoder = new TextEncoder();

export interface SyntheticFixture {
  /** Stable identifier used in test failure messages. */
  id: string;
  /** Human-readable description (e.g. "classic bar with dataLabels + trendline"). */
  description: string;
  /** xlsx bytes ready to pass to `Workbook.xlsx.load()`. */
  bytes: Uint8Array;
}

/** Marker every synthetic fixture is decorated with. */
export const SYNTHETIC_MARKER =
  "<!-- SYNTHETIC-FIXTURE: produced by excelts (not by Excel/LibreOffice/WPS) -->";

/** Pre-compiled `<c:chart>` decoration regex used by `injectSyntheticMarker`. */
const C_CHART_OPEN = /<c:chart(\s|>)/;
const CX_CHART_OPEN = /<cx:chart(\s|>)/;

/**
 * Decorate the chart XML inside an xlsx buffer with the synthetic marker
 * AND a vendor-extension `extLst` so mutation tests can prove that
 * unknown extension XML survives high-level mutations + writeBuffer.
 *
 * The decoration is intentionally idempotent: applying it twice is a
 * no-op (`includes(SYNTHETIC_MARKER)`). The function does not validate
 * that the buffer contains a chart — callers that need a chart fixture
 * should use the dedicated builders below.
 */
export async function decorateWithSyntheticMarkers(
  bytes: Uint8Array,
  extLstFragment: string = vendorExtLstFragment("{synthetic-roundtrip-fixture}")
): Promise<Uint8Array> {
  const entries = await extractAll(bytes);
  const updated = new Map<string, ExtractedFile>(entries);
  for (const [name, file] of entries) {
    if (!/^xl\/charts\/chart(?:Ex)?\d+[.]xml$/.test(name)) {
      continue;
    }
    const xml = textDecoder.decode(file.data);
    if (xml.includes(SYNTHETIC_MARKER)) {
      continue;
    }
    let next = xml;
    const isChartEx = name.includes("chartEx");
    const opener = isChartEx ? CX_CHART_OPEN : C_CHART_OPEN;
    const match = opener.exec(next);
    if (match) {
      next = `${next.slice(0, match.index)}${SYNTHETIC_MARKER}${next.slice(match.index)}`;
    }
    if (!isChartEx) {
      // Append a vendor `extLst` extension just inside `</c:chart>` so
      // round-trip code paths must preserve it across passthrough,
      // structured re-render, and raw-patch mutations.
      next = next.replace("</c:chart>", `${extLstFragment}</c:chart>`);
    } else {
      // ChartEx variant: append a `cx:extLst` extension at the
      // `cx:chartSpace` level so the structured renderer (which
      // preserves `chartSpace.extLst` via the raw-element passthrough)
      // must round-trip it. The renderer drops unknown XML inside
      // `cx:chart`, so we anchor to the chartSpace closer.
      const cxFragment = cxVendorExtLstFragment("{synthetic-cx-roundtrip-fixture}");
      next = next.replace("</cx:chartSpace>", `${cxFragment}</cx:chartSpace>`);
    }
    updated.set(name, { ...file, data: textEncoder.encode(next) });

    // Inject style/colors sidecars + chart rels for classic charts so
    // mutation tests can prove the writer preserves them byte-for-byte
    // across the round-trip.
    if (!isChartEx) {
      const m = /^xl\/charts\/chart(\d+)\.xml$/.exec(name);
      if (m) {
        const n = m[1];
        injectChartSidecars(updated, Number(n));
      }
    }
  }

  // Patch [Content_Types].xml in one pass to add Override entries for
  // every newly-injected style{N}.xml / colors{N}.xml part so the
  // package audit accepts the synthetic fixture as well-formed.
  patchContentTypesForChartSidecars(updated);

  return createZip([...updated.entries()].map(([name, file]) => ({ name, data: file.data })));
}

function patchContentTypesForChartSidecars(updated: Map<string, ExtractedFile>): void {
  const ctEntry = updated.get("[Content_Types].xml");
  if (!ctEntry) {
    return;
  }
  let ct = textDecoder.decode(ctEntry.data);
  const stylePaths = [...updated.keys()].filter(k => /^xl\/charts\/style\d+\.xml$/.test(k));
  const colorsPaths = [...updated.keys()].filter(k => /^xl\/charts\/colors\d+\.xml$/.test(k));
  const additions: string[] = [];
  for (const p of stylePaths) {
    if (!ct.includes(`PartName="/${p}"`)) {
      additions.push(
        `<Override PartName="/${p}" ContentType="application/vnd.ms-office.chartstyle+xml"/>`
      );
    }
  }
  for (const p of colorsPaths) {
    if (!ct.includes(`PartName="/${p}"`)) {
      additions.push(
        `<Override PartName="/${p}" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>`
      );
    }
  }
  if (additions.length === 0) {
    return;
  }
  ct = ct.replace("</Types>", `${additions.join("")}</Types>`);
  updated.set("[Content_Types].xml", { ...ctEntry, data: textEncoder.encode(ct) });
}

/** Injects synthetic style{n}.xml + colors{n}.xml sidecars + chart{n}.xml.rels
 *  if they are not already present. The XML bodies use the minimal
 *  schemas required for `validateXlsxBuffer` and Excel's own loader to
 *  accept them. */
function injectChartSidecars(updated: Map<string, ExtractedFile>, n: number): void {
  const stylePath = `xl/charts/style${n}.xml`;
  const colorsPath = `xl/charts/colors${n}.xml`;
  const relsPath = `xl/charts/_rels/chart${n}.xml.rels`;
  if (!updated.has(stylePath)) {
    const styleXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"` +
      ` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="201"/>`;
    const data = textEncoder.encode(styleXml);
    updated.set(stylePath, { path: stylePath, data, type: "file", size: data.length, mode: 0 });
  }
  if (!updated.has(colorsPath)) {
    const colorsXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"` +
      ` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10"/>`;
    const data = textEncoder.encode(colorsXml);
    updated.set(colorsPath, { path: colorsPath, data, type: "file", size: data.length, mode: 0 });
  }
  if (!updated.has(relsPath)) {
    const relsXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1"` +
      ` Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle"` +
      ` Target="style${n}.xml"/>` +
      `<Relationship Id="rId2"` +
      ` Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle"` +
      ` Target="colors${n}.xml"/>` +
      `</Relationships>`;
    const data = textEncoder.encode(relsXml);
    updated.set(relsPath, { path: relsPath, data, type: "file", size: data.length, mode: 0 });
  }
}

/**
 * Standard vendor `extLst` fragment used by raw-passthrough mutation tests.
 * The URI is namespaced with a `{…}` UUID-shaped placeholder so it cannot
 * collide with a real Office extension.
 */
export function vendorExtLstFragment(uri: string): string {
  return (
    `<c:extLst><c:ext uri="${uri}">` +
    `<c15:foo xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart" val="42"/>` +
    `</c:ext></c:extLst>`
  );
}

/**
 * ChartEx (Office 2016+) variant: `<cx:extLst>` lives in the
 * `cx` namespace, so a separate fragment is required. The renderer
 * preserves these via `chartSpace.extLst` (raw bytes) — see
 * `chart-ex-parser.ts` and `chart-ex-renderer.ts`.
 */
export function cxVendorExtLstFragment(uri: string): string {
  return (
    `<cx:extLst><cx:ext uri="${uri}">` +
    `<cx15:foo xmlns:cx15="http://schemas.microsoft.com/office/drawing/2014/chartex" val="42"/>` +
    `</cx:ext></cx:extLst>`
  );
}

// ---------------------------------------------------------------------------
// Classic per-type fixture
// ---------------------------------------------------------------------------

const SAMPLE_ROWS: Array<Array<string | number>> = [
  ["Cat", "Region", "Open", "High", "Low", "Close", "Value", "Value2", "Size"],
  ["A", "North", 10, 15, 8, 12, 30, 3, 5],
  ["B", "North", 20, 25, 18, 21, 20, 2, 7],
  ["C", "South", 12, 18, 9, 16, 10, 1, 9]
];

function seedDataSheet(wb: Workbook): void {
  const sheet = wb.addWorksheet("Data");
  sheet.addRows(SAMPLE_ROWS);
}

/**
 * Build one xlsx buffer per classic chart preset (16 types). Each chart
 * carries the cross-cutting features that historically broke during
 * round-trip: title, legend, axis options, dataLabels, dataPoints,
 * trendline, errorBars, fill/line colours, vendor extLst.
 */
export async function buildClassicPresetFixtures(): Promise<SyntheticFixture[]> {
  const fixtures: SyntheticFixture[] = [];
  for (const preset of EXCEL_CHART_PRESETS as ExcelChartPreset[]) {
    const presetType = applyChartPreset(preset, { series: [] }).type;
    const noAxis =
      presetType === "ofPie" ||
      presetType === "pie" ||
      presetType === "pie3D" ||
      presetType === "doughnut";
    const noTrendline = noAxis || presetType === "surface" || presetType === "surface3D";
    const isScatter = presetType === "scatter";
    const isBubble = presetType === "bubble";
    const isStock = presetType === "stock";

    // Pick a `dLblPos` that Excel accepts for this preset's chart
    // type. See `VALID_DLBL_POSITIONS_BY_TYPE` in `chart-builder.ts`
    // for the per-type allow-lists. `doughnut` rejects every
    // position value, and `surface` / `surface3D` do not support
    // series-level data labels at all, so skip the field entirely
    // for those presets.
    const skipDataLabels = presetType === "surface" || presetType === "surface3D";
    const dataLabelPosition: "outEnd" | "t" | "ctr" | undefined =
      presetType === "doughnut"
        ? undefined
        : presetType === "bar" ||
            presetType === "bar3D" ||
            presetType === "pie" ||
            presetType === "pie3D" ||
            presetType === "ofPie"
          ? "outEnd"
          : presetType === "area" || presetType === "area3D"
            ? "ctr"
            : "t";

    const wb = new Workbook();
    seedDataSheet(wb);
    const sheet = wb.getWorksheet("Data")!;

    const baseSeries = {
      name: String(preset),
      categories: "Data!$A$2:$A$4",
      values: "Data!$G$2:$G$4",
      fill: "4472C4",
      line: "ED7D31",
      dataLabels: skipDataLabels
        ? undefined
        : dataLabelPosition !== undefined
          ? { showVal: true, position: dataLabelPosition }
          : { showVal: true },
      trendline: noTrendline ? undefined : { type: "linear" as const, displayEq: true },
      errorBars: noAxis ? undefined : { type: "fixedVal" as const, value: 1 },
      dataPoints: [{ index: 0, fill: "FFC000" }]
    };

    const series = isScatter
      ? [
          {
            name: String(preset),
            xValues: "Data!$H$2:$H$4",
            values: "Data!$G$2:$G$4",
            marker: { symbol: "circle" as const, size: 6 }
          }
        ]
      : isBubble
        ? [{ ...baseSeries, xValues: "Data!$H$2:$H$4", bubbleSize: "Data!$I$2:$I$4" }]
        : isStock
          ? [
              { name: "Open", categories: "Data!$A$2:$A$4", values: "Data!$C$2:$C$4" },
              { name: "High", categories: "Data!$A$2:$A$4", values: "Data!$D$2:$D$4" },
              { name: "Low", categories: "Data!$A$2:$A$4", values: "Data!$E$2:$E$4" },
              { name: "Close", categories: "Data!$A$2:$A$4", values: "Data!$F$2:$F$4" }
            ]
          : [baseSeries];

    const options = applyChartPreset(preset, {
      series,
      title: `Synthetic ${preset}`,
      valueAxis: noAxis
        ? undefined
        : { numFmt: "#,##0", majorGridlines: true, lineColor: "70AD47" },
      categoryAxis: noAxis ? undefined : { textRotation: -30 },
      plotAreaOptions: { layout: { manualLayout: { x: 0.1, y: 0.1, w: 0.8, h: 0.8 } } }
    });
    sheet.addChart(options, "K1:Q12");

    const raw = new Uint8Array(await wb.xlsx.writeBuffer());
    const decorated = await decorateWithSyntheticMarkers(raw);
    fixtures.push({
      id: `classic-${preset}`,
      description: `synthetic classic preset "${preset}" with title/legend/axis/dataLabels/trendline/errorBars/extLst`,
      bytes: decorated
    });
  }
  return fixtures;
}

// ---------------------------------------------------------------------------
// ChartEx per-layoutId fixture
// ---------------------------------------------------------------------------

export const CHART_EX_TYPES: ChartExType[] = [
  "sunburst",
  "treemap",
  "waterfall",
  "funnel",
  "histogram",
  "pareto",
  "boxWhisker",
  "regionMap"
];

export async function buildChartExFixtures(): Promise<SyntheticFixture[]> {
  const fixtures: SyntheticFixture[] = [];
  for (const type of CHART_EX_TYPES) {
    const wb = new Workbook();
    seedDataSheet(wb);
    const sheet = wb.getWorksheet("Data")!;

    sheet.addChartEx(
      {
        type,
        categories: "Data!$A$2:$A$4",
        series: [
          {
            name: type,
            values: "Data!$G$2:$G$4",
            hierarchy: type === "sunburst" || type === "treemap" ? ["Data!$B$2:$B$4"] : undefined,
            subtotals: type === "waterfall" ? [2] : undefined
          }
        ],
        title: `Synthetic ${type}`,
        ...(type === "histogram" || type === "pareto" ? { binning: { binCount: 3 } } : {})
      },
      "K1:Q12"
    );

    const raw = new Uint8Array(await wb.xlsx.writeBuffer());
    // ChartEx fixtures get both the SYNTHETIC marker comment and a
    // vendor `cx:extLst` on `cx:chartSpace`. The structured renderer
    // preserves the latter via `chartSpace.extLst` raw passthrough so
    // mutation tests can prove unknown ChartEx extension XML survives
    // structured re-render.
    const decorated = await decorateWithSyntheticMarkers(raw);
    fixtures.push({
      id: `chartEx-${type}`,
      description: `synthetic ChartEx layoutId "${type}" with title/categories/series`,
      bytes: decorated
    });
  }
  return fixtures;
}

// ---------------------------------------------------------------------------
// Combo / axis matrix fixture
// ---------------------------------------------------------------------------

export async function buildComboAxisFixtures(): Promise<SyntheticFixture[]> {
  const fixtures: SyntheticFixture[] = [];

  // 1. Stacked bar + line on a secondary axis.
  {
    const wb = new Workbook();
    seedDataSheet(wb);
    wb.getWorksheet("Data")!.addComboChart(
      {
        groups: [
          {
            type: "bar",
            barDir: "col",
            grouping: "stacked",
            series: [
              { name: "Rev1", categories: "Data!$A$2:$A$4", values: "Data!$C$2:$C$4" },
              { name: "Rev2", categories: "Data!$A$2:$A$4", values: "Data!$D$2:$D$4" }
            ]
          },
          {
            type: "line",
            useSecondaryAxis: true,
            series: [{ name: "Growth", categories: "Data!$A$2:$A$4", values: "Data!$E$2:$E$4" }]
          }
        ],
        title: "Stacked bar + secondary line"
      },
      "K1:Q14"
    );
    const raw = new Uint8Array(await wb.xlsx.writeBuffer());
    fixtures.push({
      id: "combo-stacked-secondary",
      description: "stacked bar + line with secondary value axis",
      bytes: await decorateWithSyntheticMarkers(raw)
    });
  }

  // 2. Scatter + line combo.
  {
    const wb = new Workbook();
    seedDataSheet(wb);
    wb.getWorksheet("Data")!.addComboChart(
      {
        groups: [
          {
            type: "scatter",
            scatterStyle: "lineMarker",
            series: [
              {
                name: "Cloud",
                xValues: "Data!$C$2:$C$4",
                values: "Data!$D$2:$D$4",
                marker: { symbol: "diamond", size: 8 }
              }
            ]
          },
          {
            type: "line",
            useSecondaryAxis: true,
            series: [{ name: "Trend", categories: "Data!$A$2:$A$4", values: "Data!$E$2:$E$4" }]
          }
        ],
        title: "Scatter + secondary line"
      },
      "K1:Q14"
    );
    const raw = new Uint8Array(await wb.xlsx.writeBuffer());
    fixtures.push({
      id: "combo-scatter-line",
      description: "scatter + line combo with secondary axis",
      bytes: await decorateWithSyntheticMarkers(raw)
    });
  }

  // 3. Log-scale value axis.
  {
    const wb = new Workbook();
    seedDataSheet(wb);
    wb.getWorksheet("Data")!.addChart(
      {
        type: "line",
        series: [{ categories: "Data!$A$2:$A$4", values: "Data!$G$2:$G$4" }],
        title: "Log axis line",
        valueAxis: { logBase: 10, min: 1, max: 100, majorGridlines: true }
      },
      "K1:Q14"
    );
    const raw = new Uint8Array(await wb.xlsx.writeBuffer());
    fixtures.push({
      id: "axis-log",
      description: "line chart with log10 value axis",
      bytes: await decorateWithSyntheticMarkers(raw)
    });
  }

  // 4. Multiple secondary axes (bar primary, line secondary, area secondary).
  {
    const wb = new Workbook();
    seedDataSheet(wb);
    wb.getWorksheet("Data")!.addComboChart(
      {
        groups: [
          {
            type: "bar",
            barDir: "col",
            series: [{ categories: "Data!$A$2:$A$4", values: "Data!$C$2:$C$4" }]
          },
          {
            type: "line",
            useSecondaryAxis: true,
            series: [{ categories: "Data!$A$2:$A$4", values: "Data!$D$2:$D$4" }]
          },
          {
            type: "area",
            useSecondaryAxis: true,
            series: [{ categories: "Data!$A$2:$A$4", values: "Data!$E$2:$E$4" }]
          }
        ],
        title: "Three groups, two share secondary axis"
      },
      "K1:Q14"
    );
    const raw = new Uint8Array(await wb.xlsx.writeBuffer());
    fixtures.push({
      id: "combo-three-groups-shared-secondary",
      description: "three combo groups, two sharing the same secondary axis",
      bytes: await decorateWithSyntheticMarkers(raw)
    });
  }

  // 5. 3D bar combo + line.
  {
    const wb = new Workbook();
    seedDataSheet(wb);
    wb.getWorksheet("Data")!.addComboChart(
      {
        groups: [
          {
            type: "bar3D",
            series: [{ categories: "Data!$A$2:$A$4", values: "Data!$C$2:$C$4" }]
          },
          {
            type: "line",
            series: [{ categories: "Data!$A$2:$A$4", values: "Data!$D$2:$D$4" }]
          }
        ],
        title: "3D bar + line"
      },
      "K1:Q14"
    );
    const raw = new Uint8Array(await wb.xlsx.writeBuffer());
    fixtures.push({
      id: "combo-3d-bar-line",
      description: "3D bar group with serAx + 2D line group",
      bytes: await decorateWithSyntheticMarkers(raw)
    });
  }

  return fixtures;
}

// ---------------------------------------------------------------------------
// Chartsheet fixture
// ---------------------------------------------------------------------------

export async function buildChartsheetFixtures(): Promise<SyntheticFixture[]> {
  const fixtures: SyntheticFixture[] = [];
  const wb = new Workbook();
  seedDataSheet(wb);

  wb.addChartsheet("Pie Chart Sheet", {
    chart: {
      type: "pie",
      series: [{ categories: "Data!$A$2:$A$4", values: "Data!$G$2:$G$4" }],
      title: "Synthetic Pie Sheet"
    },
    tabSelected: false,
    zoomScale: 80,
    zoomToFit: true,
    pageMargins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
    pageSetup: { orientation: "landscape", paperSize: 9 }
  });

  wb.addChartsheet("Funnel Chart Sheet", {
    chart: {
      type: "funnel",
      categories: "Data!$A$2:$A$4",
      series: [{ values: "Data!$G$2:$G$4" }],
      title: "Synthetic Funnel Sheet"
    },
    state: "visible"
  });

  wb.addChartsheet("Combo Chart Sheet", {
    chart: {
      groups: [
        {
          type: "bar",
          barDir: "col",
          series: [{ categories: "Data!$A$2:$A$4", values: "Data!$G$2:$G$4" }]
        },
        {
          type: "line",
          useSecondaryAxis: true,
          series: [{ categories: "Data!$A$2:$A$4", values: "Data!$E$2:$E$4" }]
        }
      ],
      title: "Synthetic Combo Sheet"
    }
  });

  // Hidden chartsheet — exercises the `state="hidden"` write path.
  wb.addChartsheet("Hidden Sheet", {
    state: "hidden",
    chart: {
      type: "bar",
      series: [{ categories: "Data!$A$2:$A$4", values: "Data!$G$2:$G$4" }]
    }
  });

  const raw = new Uint8Array(await wb.xlsx.writeBuffer());
  fixtures.push({
    id: "chartsheet-mixed",
    description: "four chartsheets (pie, funnel, combo, hidden) plus a Data worksheet",
    bytes: await decorateWithSyntheticMarkers(raw)
  });
  return fixtures;
}

// ---------------------------------------------------------------------------
// Pivot chart fixture
// ---------------------------------------------------------------------------

export async function buildPivotChartFixtures(): Promise<SyntheticFixture[]> {
  const fixtures: SyntheticFixture[] = [];

  // Single pivot table + chart with field buttons + refreshOnOpen.
  {
    const wb = new Workbook();
    const data = wb.addWorksheet("Data");
    data.addRows([
      ["Region", "Product", "Revenue", "Units"],
      ["West", "A", 10, 1],
      ["West", "B", 20, 2],
      ["East", "A", 30, 3],
      ["East", "B", 40, 4]
    ]);
    const pivotSheet = wb.addWorksheet("Pivot");
    const pivot = pivotSheet.addPivotTable({
      sourceSheet: data,
      rows: ["Region"],
      columns: ["Product"],
      values: ["Revenue", "Units"],
      metric: "sum"
    });
    pivotSheet.addPivotChart(
      pivot,
      {
        type: "bar",
        barDir: "col",
        series: [
          {
            name: "Pivot Revenue",
            categories: "'Pivot'!$A$4:$A$5",
            values: "'Pivot'!$B$4:$B$5"
          }
        ],
        title: "Synthetic pivot chart",
        pivotChartOptions: {
          refreshOnOpen: true,
          dropZonesVisible: true,
          dropZoneCategories: true,
          dropZoneData: true,
          dropZoneSeries: true
        }
      },
      "D1:J12"
    );
    const raw = new Uint8Array(await wb.xlsx.writeBuffer());
    fixtures.push({
      id: "pivot-chart-multi-value",
      description:
        "pivot table with row+column+two value fields plus pivot chart with all drop zone buttons",
      bytes: await decorateWithSyntheticMarkers(raw)
    });
  }

  // Multiple pivot caches + multiple pivot charts.
  {
    const wb = new Workbook();
    const sales = wb.addWorksheet("Sales");
    sales.addRows([
      ["Region", "Revenue"],
      ["West", 10],
      ["East", 20]
    ]);
    const ops = wb.addWorksheet("Ops");
    ops.addRows([
      ["Team", "Tickets"],
      ["Alpha", 100],
      ["Beta", 200]
    ]);
    const salesPivot = wb.addWorksheet("SalesPivot");
    const salesP = salesPivot.addPivotTable({
      sourceSheet: sales,
      rows: ["Region"],
      values: ["Revenue"],
      metric: "sum"
    });
    salesPivot.addPivotChart(
      salesP,
      {
        type: "pie",
        series: [
          {
            categories: "'SalesPivot'!$A$4:$A$5",
            values: "'SalesPivot'!$B$4:$B$5"
          }
        ],
        pivotChartOptions: { refreshOnOpen: true, dropZonesVisible: false }
      },
      "D1:J12"
    );
    const opsPivot = wb.addWorksheet("OpsPivot");
    const opsP = opsPivot.addPivotTable({
      sourceSheet: ops,
      rows: ["Team"],
      values: ["Tickets"],
      metric: "sum"
    });
    opsPivot.addPivotChart(
      opsP,
      {
        type: "bar",
        barDir: "bar",
        series: [
          {
            categories: "'OpsPivot'!$A$4:$A$5",
            values: "'OpsPivot'!$B$4:$B$5"
          }
        ]
      },
      "D1:J12"
    );
    const raw = new Uint8Array(await wb.xlsx.writeBuffer());
    fixtures.push({
      id: "pivot-chart-multi-cache",
      description: "two independent pivot caches each driving a pivot chart",
      bytes: await decorateWithSyntheticMarkers(raw)
    });
  }

  return fixtures;
}

// ---------------------------------------------------------------------------
// Common helpers exported for tests
// ---------------------------------------------------------------------------

export async function readChartXmlEntries(
  bytes: Uint8Array
): Promise<{ classic: Map<string, string>; chartEx: Map<string, string> }> {
  const entries = await extractAll(bytes);
  const classic = new Map<string, string>();
  const chartEx = new Map<string, string>();
  for (const [name, file] of entries) {
    if (/^xl\/charts\/chart\d+[.]xml$/.test(name)) {
      classic.set(name, textDecoder.decode(file.data));
    } else if (/^xl\/charts\/chartEx\d+[.]xml$/.test(name)) {
      chartEx.set(name, textDecoder.decode(file.data));
    }
  }
  return { classic, chartEx };
}

export async function getEntryText(bytes: Uint8Array, path: string): Promise<string | undefined> {
  const entries = await extractAll(bytes);
  const entry = entries.get(path);
  return entry ? textDecoder.decode(entry.data) : undefined;
}
