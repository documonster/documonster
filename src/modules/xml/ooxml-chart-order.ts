/**
 * Where each content model of a DrawingML chart applies — shared by the workbook validator (`check-chart.ts`)
 * and the document order check (`writer/order-check.ts`), because both modules write charts and a rule written
 * twice is a rule that drifts.
 *
 * The order of each model is not here; it is the generated table each module carries
 * (`ooxml-sequences.ts`, from `scripts/extract-ooxml-sequences.ts`). A rule names its type by key, and each
 * module's tests hold every key to its own table.
 *
 * Rules match by local name anywhere in a chart part: every name below is unambiguous there, including the
 * DrawingML ones (`spPr`, `p`, `rPr`), which a chart part has no other use for. A series is the exception —
 * `c:ser` has a different model under each chart type — so its rules also name the parent.
 */
/**
 * A chart element and its content model. ECMA-376 §21.2 declares each chart element’s children as an
 * `xsd:sequence`, and Excel refuses — or silently repairs — a chart whose children are out of order. Members of an
 * `xsd:choice` share one position in the generated tables, which is what lets `c:plotArea` hold its chart types and
 * its axes in any order among themselves.
 */
export interface ChartOrderRule {
  /** Parent element local name (namespace prefix stripped). */
  parent: string;
  /** The parent's content model. */
  type: ChartSequenceType;
  /**
   * Whether the content model is the *complete* set of permitted children, so a child outside it is an error
   * rather than something to skip.
   *
   * Off by default, and deliberately opt-in: the order check must stay tolerant of extension elements and of
   * children this table has not enumerated, or it would report a healthy file. But for a type whose content
   * model is short and closed, silence is the wrong answer — a `<c:overlay>` inside `<c:dispUnitsLbl>` ranked
   * -1, was skipped as "not in the table", and shipped in a file Excel repaired while this validator reported
   * zero problems.
   */
  closed?: boolean;
}

/** A content model's key in a module's generated table: `c:CT_…` for the chart's own, `a:CT_…` for DrawingML. */
export type ChartSequenceType = `c:CT_${string}` | `a:CT_${string}`;

export const CHART_ORDER_RULES: readonly ChartOrderRule[] = [
  { parent: "chartSpace", type: "c:CT_ChartSpace" },
  { parent: "chart", type: "c:CT_Chart" },
  { parent: "plotArea", type: "c:CT_PlotArea" },
  // CT_DLbls — the biggest historical offender: `separator` must come before `showLeaderLines`.
  { parent: "dLbls", type: "c:CT_DLbls" },
  // CT_DLbl. Its `delete` branch excludes the display flags; that choice is `checkClassicDLblChoice`'s.
  { parent: "dLbl", type: "c:CT_DLbl" },
  { parent: "scaling", type: "c:CT_Scaling" },
  { parent: "catAx", type: "c:CT_CatAx" },
  { parent: "valAx", type: "c:CT_ValAx" },
  { parent: "dateAx", type: "c:CT_DateAx" },
  { parent: "serAx", type: "c:CT_SerAx" },
  { parent: "dispUnits", type: "c:CT_DispUnits" },
  // CT_DispUnitsLbl. **Not CT_Title**, though it is easy to render as one: the sequence is `layout` before
  // `tx` — the reverse of a title's — and there is no `overlay` and no `extLst`. Closed, because those two
  // absences are the whole point of listing it.
  { parent: "dispUnitsLbl", type: "c:CT_DispUnitsLbl", closed: true },
  { parent: "trendline", type: "c:CT_Trendline" },
  { parent: "trendlineLbl", type: "c:CT_TrendlineLbl" },
  { parent: "errBars", type: "c:CT_ErrBars" },
  { parent: "view3D", type: "c:CT_View3D" },
  { parent: "legend", type: "c:CT_Legend" },
  { parent: "legendEntry", type: "c:CT_LegendEntry" },
  { parent: "title", type: "c:CT_Title" },
  // Chart-type elements. Excel tolerates many mis-orderings here, but LibreOffice strict mode rejects them, and
  // so do third-party OOXML validators (Microsoft's SDK, Calligra, xmllint + XSD).
  { parent: "areaChart", type: "c:CT_AreaChart" },
  { parent: "area3DChart", type: "c:CT_Area3DChart" },
  { parent: "barChart", type: "c:CT_BarChart" },
  { parent: "bar3DChart", type: "c:CT_Bar3DChart" },
  { parent: "lineChart", type: "c:CT_LineChart" },
  { parent: "line3DChart", type: "c:CT_Line3DChart" },
  { parent: "pieChart", type: "c:CT_PieChart" },
  { parent: "pie3DChart", type: "c:CT_Pie3DChart" },
  { parent: "doughnutChart", type: "c:CT_DoughnutChart" },
  { parent: "ofPieChart", type: "c:CT_OfPieChart" },
  { parent: "scatterChart", type: "c:CT_ScatterChart" },
  { parent: "bubbleChart", type: "c:CT_BubbleChart" },
  { parent: "radarChart", type: "c:CT_RadarChart" },
  { parent: "stockChart", type: "c:CT_StockChart" },
  { parent: "surfaceChart", type: "c:CT_SurfaceChart" },
  { parent: "surface3DChart", type: "c:CT_Surface3DChart" },
  // Auxiliary types that appear throughout the chart graph.
  { parent: "dPt", type: "c:CT_DPt" },
  // A series' `c:marker` is CT_Marker; a line chart's `c:marker` is a CT_Boolean with no children to rank.
  { parent: "marker", type: "c:CT_Marker" },
  { parent: "upDownBars", type: "c:CT_UpDownBars" },
  { parent: "numRef", type: "c:CT_NumRef" },
  { parent: "strRef", type: "c:CT_StrRef" },
  { parent: "multiLvlStrRef", type: "c:CT_MultiLvlStrRef" },
  { parent: "numCache", type: "c:CT_NumData" },
  { parent: "strCache", type: "c:CT_StrData" },
  { parent: "multiLvlStrCache", type: "c:CT_MultiLvlStrData" },
  { parent: "numLit", type: "c:CT_NumData" },
  { parent: "strLit", type: "c:CT_StrData" },
  { parent: "layout", type: "c:CT_Layout" },
  { parent: "manualLayout", type: "c:CT_ManualLayout" },
  { parent: "dTable", type: "c:CT_DTable" },
  { parent: "bandFmt", type: "c:CT_BandFmt" },
  { parent: "pictureOptions", type: "c:CT_PictureOptions" },
  { parent: "protection", type: "c:CT_Protection" },
  { parent: "printSettings", type: "c:CT_PrintSettings" },
  // DrawingML inside the chart: shape properties and text. Each local name below is DrawingML's alone in a
  // chart part, so matching by local name cannot reach a chart element by mistake.
  { parent: "spPr", type: "a:CT_ShapeProperties" },
  { parent: "xfrm", type: "a:CT_Transform2D" },
  { parent: "ln", type: "a:CT_LineProperties" },
  { parent: "txPr", type: "a:CT_TextBody" },
  { parent: "rich", type: "a:CT_TextBody" },
  { parent: "bodyPr", type: "a:CT_TextBodyProperties" },
  { parent: "lstStyle", type: "a:CT_TextListStyle" },
  { parent: "p", type: "a:CT_TextParagraph" },
  { parent: "pPr", type: "a:CT_TextParagraphProperties" },
  { parent: "r", type: "a:CT_RegularTextRun" },
  { parent: "rPr", type: "a:CT_TextCharacterProperties" },
  { parent: "defRPr", type: "a:CT_TextCharacterProperties" },
  { parent: "endParaRPr", type: "a:CT_TextCharacterProperties" }
];

/**
 * Series (`c:ser`) child order depends on which chart-type element
 * encloses it. The same tag maps to CT_BarSer / CT_LineSer /
 * CT_PieSer / CT_AreaSer / CT_ScatterSer / CT_BubbleSer / CT_RadarSer
 * / CT_SurfaceSer, each with its own sequence.
 */
export interface ChartSeriesOrderRule {
  /** Chart-type parents this rule applies to. */
  chartTypes: readonly string[];
  /** The content model of `c:ser` under these parents. */
  type: ChartSequenceType;
}

export const CHART_SERIES_ORDER_RULES: readonly ChartSeriesOrderRule[] = [
  { chartTypes: ["barChart", "bar3DChart"], type: "c:CT_BarSer" },
  { chartTypes: ["lineChart", "line3DChart", "stockChart"], type: "c:CT_LineSer" },
  { chartTypes: ["pieChart", "pie3DChart", "doughnutChart", "ofPieChart"], type: "c:CT_PieSer" },
  { chartTypes: ["areaChart", "area3DChart"], type: "c:CT_AreaSer" },
  { chartTypes: ["scatterChart"], type: "c:CT_ScatterSer" },
  { chartTypes: ["bubbleChart"], type: "c:CT_BubbleSer" },
  { chartTypes: ["radarChart"], type: "c:CT_RadarSer" },
  { chartTypes: ["surfaceChart", "surface3DChart"], type: "c:CT_SurfaceSer" }
];

const BY_PARENT = new Map(CHART_ORDER_RULES.map(rule => [rule.parent, rule.type]));

const SERIES_BY_CHART_TYPE = new Map(
  CHART_SERIES_ORDER_RULES.flatMap(rule =>
    rule.chartTypes.map(chartType => [chartType, rule.type] as const)
  )
);

/** The content model of a chart element, by its local name and its parent's — or `undefined` when it has none here. */
export function chartContentModel(
  local: string,
  parentLocal: string | undefined
): ChartSequenceType | undefined {
  if (local === "ser") {
    return parentLocal === undefined ? undefined : SERIES_BY_CHART_TYPE.get(parentLocal);
  }
  return BY_PARENT.get(local);
}
