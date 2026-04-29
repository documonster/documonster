import type { AddChartExOptions, ChartExType } from "./chart-ex-types";
import type { AddChartOptions, AddChartSeriesOptions } from "./types";

type PresetSeriesDefaults = Partial<Pick<AddChartSeriesOptions, "bubble3D" | "explosion">>;

interface PresetConfig {
  options: Partial<AddChartOptions>;
  seriesDefaults?: PresetSeriesDefaults;
}

const barPreset = (
  barDir: NonNullable<AddChartOptions["barDir"]>,
  grouping: Extract<
    NonNullable<AddChartOptions["grouping"]>,
    "clustered" | "stacked" | "percentStacked"
  >,
  shape?: NonNullable<AddChartOptions["shape"]>
): PresetConfig => ({
  options: { type: shape ? "bar3D" : "bar", barDir, grouping, shape }
});

const bar3DPreset = (
  barDir: NonNullable<AddChartOptions["barDir"]>,
  grouping: Extract<
    NonNullable<AddChartOptions["grouping"]>,
    "clustered" | "stacked" | "percentStacked"
  >,
  shape?: NonNullable<AddChartOptions["shape"]>
): PresetConfig => ({
  options: { type: "bar3D", barDir, grouping, shape }
});

export const CHART_PRESETS = {
  columnClustered: barPreset("col", "clustered"),
  columnStacked: barPreset("col", "stacked"),
  columnStacked100: barPreset("col", "percentStacked"),
  colClustered: barPreset("col", "clustered"),
  colStacked: barPreset("col", "stacked"),
  colPercentStacked: barPreset("col", "percentStacked"),
  colStacked100: barPreset("col", "percentStacked"),
  barClustered: barPreset("bar", "clustered"),
  barStacked: barPreset("bar", "stacked"),
  barPercentStacked: barPreset("bar", "percentStacked"),
  barStacked100: barPreset("bar", "percentStacked"),
  pieOfPie: { options: { type: "ofPie", ofPieType: "pie" } },
  barOfPie: { options: { type: "ofPie", ofPieType: "bar" } },

  column3DClustered: bar3DPreset("col", "clustered"),
  column3DStacked: bar3DPreset("col", "stacked"),
  column3DStacked100: bar3DPreset("col", "percentStacked"),
  col3DClustered: bar3DPreset("col", "clustered"),
  col3DStacked: bar3DPreset("col", "stacked"),
  col3DStacked100: bar3DPreset("col", "percentStacked"),
  bar3DClustered: bar3DPreset("bar", "clustered"),
  bar3DStacked: bar3DPreset("bar", "stacked"),
  bar3DStacked100: bar3DPreset("bar", "percentStacked"),

  col3DConeClustered: bar3DPreset("col", "clustered", "cone"),
  col3DConeStacked: bar3DPreset("col", "stacked", "cone"),
  col3DConeStacked100: bar3DPreset("col", "percentStacked", "cone"),
  coneColClustered: bar3DPreset("col", "clustered", "cone"),
  coneColStacked: bar3DPreset("col", "stacked", "cone"),
  coneColStacked100: bar3DPreset("col", "percentStacked", "cone"),
  bar3DConeClustered: bar3DPreset("bar", "clustered", "cone"),
  bar3DConeStacked: bar3DPreset("bar", "stacked", "cone"),
  bar3DConeStacked100: bar3DPreset("bar", "percentStacked", "cone"),
  coneBarClustered: bar3DPreset("bar", "clustered", "cone"),
  coneBarStacked: bar3DPreset("bar", "stacked", "cone"),
  coneBarStacked100: bar3DPreset("bar", "percentStacked", "cone"),

  col3DCylinderClustered: bar3DPreset("col", "clustered", "cylinder"),
  col3DCylinderStacked: bar3DPreset("col", "stacked", "cylinder"),
  col3DCylinderStacked100: bar3DPreset("col", "percentStacked", "cylinder"),
  cylinderColClustered: bar3DPreset("col", "clustered", "cylinder"),
  cylinderColStacked: bar3DPreset("col", "stacked", "cylinder"),
  cylinderColStacked100: bar3DPreset("col", "percentStacked", "cylinder"),
  bar3DCylinderClustered: bar3DPreset("bar", "clustered", "cylinder"),
  bar3DCylinderStacked: bar3DPreset("bar", "stacked", "cylinder"),
  bar3DCylinderStacked100: bar3DPreset("bar", "percentStacked", "cylinder"),
  cylinderBarClustered: bar3DPreset("bar", "clustered", "cylinder"),
  cylinderBarStacked: bar3DPreset("bar", "stacked", "cylinder"),
  cylinderBarStacked100: bar3DPreset("bar", "percentStacked", "cylinder"),

  col3DPyramidClustered: bar3DPreset("col", "clustered", "pyramid"),
  col3DPyramidStacked: bar3DPreset("col", "stacked", "pyramid"),
  col3DPyramidStacked100: bar3DPreset("col", "percentStacked", "pyramid"),
  pyramidColClustered: bar3DPreset("col", "clustered", "pyramid"),
  pyramidColStacked: bar3DPreset("col", "stacked", "pyramid"),
  pyramidColStacked100: bar3DPreset("col", "percentStacked", "pyramid"),
  bar3DPyramidClustered: bar3DPreset("bar", "clustered", "pyramid"),
  bar3DPyramidStacked: bar3DPreset("bar", "stacked", "pyramid"),
  bar3DPyramidStacked100: bar3DPreset("bar", "percentStacked", "pyramid"),
  pyramidBarClustered: bar3DPreset("bar", "clustered", "pyramid"),
  pyramidBarStacked: bar3DPreset("bar", "stacked", "pyramid"),
  pyramidBarStacked100: bar3DPreset("bar", "percentStacked", "pyramid"),

  line: { options: { type: "line" } },
  lineMarkers: { options: { type: "line", showMarker: true } },
  lineStacked: { options: { type: "line", grouping: "stacked" } },
  lineStacked100: { options: { type: "line", grouping: "percentStacked" } },
  pie: { options: { type: "pie" } },
  pieExploded: { options: { type: "pie" }, seriesDefaults: { explosion: 25 } },
  pie3D: { options: { type: "pie3D" } },
  pieExploded3D: { options: { type: "pie3D" }, seriesDefaults: { explosion: 25 } },
  doughnut: { options: { type: "doughnut" } },
  doughnutExploded: {
    options: { type: "doughnut", holeSize: 50 },
    seriesDefaults: { explosion: 25 }
  },
  area: { options: { type: "area" } },
  areaStacked: { options: { type: "area", grouping: "stacked" } },
  areaStacked100: { options: { type: "area", grouping: "percentStacked" } },
  area3D: { options: { type: "area3D" } },
  area3DStacked: { options: { type: "area3D", grouping: "stacked" } },
  area3DStacked100: { options: { type: "area3D", grouping: "percentStacked" } },
  scatter: { options: { type: "scatter", scatterStyle: "marker" } },
  scatterMarker: { options: { type: "scatter", scatterStyle: "marker" } },
  scatterStraight: { options: { type: "scatter", scatterStyle: "lineMarker" } },
  scatterLines: { options: { type: "scatter", scatterStyle: "lineMarker" } },
  scatterLinesNoMarkers: { options: { type: "scatter", scatterStyle: "line" } },
  scatterSmooth: { options: { type: "scatter", scatterStyle: "smoothMarker" } },
  scatterSmoothMarker: { options: { type: "scatter", scatterStyle: "smoothMarker" } },
  scatterSmoothNoMarkers: { options: { type: "scatter", scatterStyle: "smooth" } },
  bubble: { options: { type: "bubble" } },
  bubble3D: { options: { type: "bubble" }, seriesDefaults: { bubble3D: true } },
  radar: { options: { type: "radar", radarStyle: "standard" } },
  radarMarkers: { options: { type: "radar", radarStyle: "marker" } },
  radarFilled: { options: { type: "radar", radarStyle: "filled" } },
  stockHLC: { options: { type: "stock", hiLowLines: true } },
  stockOHLC: { options: { type: "stock", upDownBars: true, hiLowLines: true } },
  stockVHLC: { options: { type: "stock", hiLowLines: true } },
  stockVOHLC: { options: { type: "stock", hiLowLines: true, upDownBars: true } },
  surface: { options: { type: "surface" } },
  surface3D: { options: { type: "surface3D" } },
  surfaceTopView: { options: { type: "surface" } },
  contour: { options: { type: "surface" } },
  wireframeSurface: { options: { type: "surface3D", wireframe: true } },
  surface3DWireframe: { options: { type: "surface3D", wireframe: true } },
  surfaceWireframe: { options: { type: "surface3D", wireframe: true } },
  wireframeContour: { options: { type: "surface", wireframe: true } },
  topViewWireframe: { options: { type: "surface", wireframe: true } }
} satisfies Readonly<Record<string, PresetConfig>>;

export type ExcelChartPreset = keyof typeof CHART_PRESETS;

export const EXCEL_CHART_PRESETS = Object.keys(CHART_PRESETS) as ExcelChartPreset[];

export const CHART_EX_PRESETS = {
  histogram: { type: "histogram" },
  pareto: { type: "pareto" },
  waterfall: { type: "waterfall" },
  funnel: { type: "funnel" },
  treemap: { type: "treemap" },
  sunburst: { type: "sunburst" },
  boxWhisker: { type: "boxWhisker" },
  boxAndWhisker: { type: "boxWhisker" },
  regionMap: { type: "regionMap" },
  map: { type: "regionMap" }
} satisfies Readonly<Record<string, { type: ChartExType }>>;

export type ExcelChartExPreset = keyof typeof CHART_EX_PRESETS;

export const EXCEL_CHART_EX_PRESETS = Object.keys(CHART_EX_PRESETS) as ExcelChartExPreset[];

export function applyChartPreset(
  preset: ExcelChartPreset,
  options: Omit<AddChartOptions, "type"> & Partial<Pick<AddChartOptions, "type">>
): AddChartOptions {
  const config: PresetConfig | undefined = CHART_PRESETS[preset];
  if (!config) {
    throw new Error(`Unknown chart preset: ${preset}`);
  }
  const merged = {
    ...config.options,
    ...options,
    type: options.type ?? config.options.type!
  } as AddChartOptions;
  if (config.seriesDefaults && merged.series) {
    merged.series = merged.series.map(series => ({ ...config.seriesDefaults, ...series }));
  }
  return merged;
}

export function applyChartExPreset(
  preset: ExcelChartExPreset,
  options: Omit<AddChartExOptions, "type"> & Partial<Pick<AddChartExOptions, "type">>
): AddChartExOptions {
  const config = CHART_EX_PRESETS[preset];
  if (!config) {
    throw new Error(`Unknown chartEx preset: ${preset}`);
  }
  return { ...options, type: options.type ?? config.type } as AddChartExOptions;
}
