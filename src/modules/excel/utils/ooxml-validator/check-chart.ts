/**
 * Chart / chartEx internal structure check.
 *
 * Classic charts (`xl/charts/chartN.xml`) must contain both `c:chart` and
 * `c:plotArea`. ChartEx charts must contain `cx:chart`, `cx:plotArea`
 * and at least one `cx:series`; series must carry `layoutId` and their
 * `dataId`/`axisId` back-references must resolve inside the chart.
 * `cx:externalData` nodes must refer to a declared relationship.
 *
 * In addition we catch five chartEx schema-violation patterns that
 * cause Excel 2016+ to drop the whole chartEx part with "Removed Part:
 * /xl/drawings/drawingN.xml (Drawing shape)":
 *
 *   - `<cx:series>` with more than one `<cx:dataId>` child
 *     (`CT_Series/dataId` has `maxOccurs="1"`; multi-dim series must
 *     use a single `<cx:data>` wrapper with multiple strDim/numDim).
 *   - `<cx:axisId>N</cx:axisId>`, `<cx:dataId>N</cx:dataId>`,
 *     `<cx:binCount>N</cx:binCount>`, `<cx:binSize>N</cx:binSize>`
 *     emitted as text content instead of `val="N"` attribute. The
 *     underlying types are `CT_UnsignedInteger`/`CT_Double`, which
 *     Excel's strict loader only accepts via the attribute form.
 *   - `<cx:auto/>` element anywhere in the chartEx. Auto binning is
 *     expressed by the ABSENCE of `binSize`/`binCount`, not a
 *     dedicated `<cx:auto/>` tag. The tag is schema-invalid.
 *   - `<cx:paretoLine>` child of `<cx:layoutPr>`. Not in the
 *     CT_SeriesLayoutProperties schema. A real pareto chart adds a
 *     second series with `layoutId="paretoLine"`.
 *   - `<cx:title>` with a direct `<cx:layout>` child. Title layout
 *     belongs in `extLst`-based extensions.
 */

import type { XmlElement } from "@xml/types";

import type { ValidationContext } from "./context";
import {
  attrByLocalName,
  collectDescendantsLocal,
  findChildLocal,
  findChildrenLocal,
  hasDescendantLocal,
  matchesLocal
} from "./xml-utils";

const CHART_PATH_RE = /^xl\/charts\/chart\d+\.xml$/;
const CHARTEX_PATH_RE = /^xl\/charts\/chartEx\d+\.xml$/;

export function checkChart(ctx: ValidationContext): void {
  for (const [path, entry] of ctx.files()) {
    if (ctx.reporter.capped) {
      return;
    }
    if (entry.type === "directory") {
      continue;
    }
    if (CHART_PATH_RE.test(path)) {
      checkClassicChart(ctx, path);
    } else if (CHARTEX_PATH_RE.test(path)) {
      checkChartEx(ctx, path);
    }
  }
}

function checkClassicChart(ctx: ValidationContext, path: string): void {
  const dom = ctx.readDom(path);
  if (!dom) {
    return;
  }
  const root = dom.root;
  if (!hasDescendantLocal(root, "chart")) {
    ctx.reporter.error("chart-missing-chart", `${path}: missing c:chart`, path);
  }
  if (!hasDescendantLocal(root, "plotArea")) {
    ctx.reporter.error("chart-missing-plotArea", `${path}: missing c:plotArea`, path);
  }
}

function checkChartEx(ctx: ValidationContext, path: string): void {
  const dom = ctx.readDom(path);
  if (!dom) {
    return;
  }
  const root = dom.root;
  if (!hasDescendantLocal(root, "chart")) {
    ctx.reporter.error("chartEx-missing-chart", `${path}: missing cx:chart`, path);
  }
  if (!hasDescendantLocal(root, "plotArea")) {
    ctx.reporter.error("chartEx-missing-plotArea", `${path}: missing cx:plotArea`, path);
  }
  const seriesList = collectDescendantsLocal(root, "series");
  if (seriesList.length === 0) {
    ctx.reporter.error("chartEx-missing-series", `${path}: missing cx:series`, path);
    return;
  }

  const dataIds = new Set(
    collectDescendantsLocal(root, "data")
      .map(el => parseInt(attrByLocalName(el, "id") ?? "", 10))
      .filter(Number.isFinite)
  );
  const axisIds = new Set(
    collectDescendantsLocal(root, "axis")
      .map(el => parseInt(attrByLocalName(el, "id") ?? "", 10))
      .filter(Number.isFinite)
  );
  for (const series of seriesList) {
    if (!attrByLocalName(series, "layoutId")) {
      ctx.reporter.error(
        "chartEx-series-missing-layoutId",
        `${path}: cx:series missing layoutId`,
        path
      );
    }
    const dataIdChildren = findChildrenLocal(series, "dataId");
    // Schema cardinality: `CT_Series/dataId` has `maxOccurs="1"`.
    // Multi-dimensional series (box-whisker, sunburst, treemap) must
    // point at a single `<cx:data>` entry that holds every strDim /
    // numDim they need; emitting multiple dataIds tells Excel the
    // series references multiple data entries which is the "Removed
    // Part: drawingN.xml" trigger.
    if (dataIdChildren.length > 1) {
      ctx.reporter.error(
        "chartEx-series-too-many-dataId",
        `${path}: cx:series has ${dataIdChildren.length} <cx:dataId> children; schema permits at most 1. ` +
          `Consolidate the referenced <cx:data> entries into a single entry.`,
        path
      );
    }
    for (const dataId of dataIdChildren) {
      const id = parseInt(attrByLocalName(dataId, "val") ?? "", 10);
      if (!dataIds.has(id)) {
        ctx.reporter.error(
          "chartEx-series-missing-data-id",
          `${path}: cx:series references missing cx:data id ${attrByLocalName(dataId, "val")}`,
          path
        );
      }
    }
    for (const axisId of findChildrenLocal(series, "axisId")) {
      const id = parseInt(attrByLocalName(axisId, "val") ?? "", 10);
      if (!axisIds.has(id)) {
        ctx.reporter.error(
          "chartEx-series-missing-axis-id",
          `${path}: cx:series references missing cx:axis id ${attrByLocalName(axisId, "val")}`,
          path
        );
      }
    }
  }

  // externalData (e.g. cx:externalData r:id="...") must resolve in the chart's rels.
  const externalDataRids = collectDescendantsLocal(root, "externalData")
    .map(el => attrByLocalName(el, "id"))
    .filter((id): id is string => !!id);
  if (externalDataRids.length > 0) {
    const relsPath = chartRelsPath(path);
    const rels = ctx.readRels(relsPath);
    for (const rid of externalDataRids) {
      if (!rels.byId.has(rid)) {
        ctx.reporter.error(
          "chartEx-externalData-missing-rel",
          `${path}: cx:externalData references missing relationship ${rid}`,
          path
        );
      }
    }
  }

  // Schema-conformance checks: text-form violations of typed elements,
  // invalid `<cx:auto/>` element, `<cx:paretoLine>` in layoutPr, and
  // direct `<cx:layout>` child of `<cx:title>`.
  checkTypedElementAttrForm(ctx, path, root);
  checkInvalidAutoElement(ctx, path, root);
  checkParetoLineInLayoutPr(ctx, path, root);
  checkTitleDirectLayoutChild(ctx, path, root);

  // Tier-2 semantic checks.
  checkAxisPosAndType(ctx, path, root);
  checkSeriesFDefinedName(ctx, path, root);
  checkWaterfallSubtotals(ctx, path, root);
}

/**
 * `<cx:axisId>`, `<cx:dataId>`, `<cx:binCount>`, `<cx:binSize>` and
 * their siblings use the `val="N"` attribute form. Earlier writer
 * revisions serialised them as text content (`<cx:axisId>2</cx:axisId>`),
 * which Excel's strict loader rejects. Flag every occurrence so the
 * output never regresses to the broken shape.
 */
const TYPED_ATTR_ONLY_ELEMENTS: readonly string[] = ["axisId", "dataId", "binCount", "binSize"];

function checkTypedElementAttrForm(ctx: ValidationContext, path: string, root: XmlElement): void {
  for (const name of TYPED_ATTR_ONLY_ELEMENTS) {
    for (const el of collectDescendantsLocal(root, name)) {
      if (ctx.reporter.capped) {
        return;
      }
      const val = attrByLocalName(el, "val");
      const text = directTextContent(el).trim();
      // Missing `val` AND present non-empty text = the broken text-form.
      if (val === undefined && text.length > 0) {
        ctx.reporter.error(
          "chartEx-typed-element-text-form",
          `${path}: <cx:${name}>${text}</cx:${name}> uses text-content form; schema requires val="${text}" attribute.`,
          path
        );
      }
    }
  }
}

/**
 * `<cx:auto/>` is NOT a valid element — auto binning is expressed by
 * the absence of both `binSize` and `binCount` inside `<cx:binning>`.
 * A literal `<cx:auto/>` tag anywhere in the chartEx makes Excel drop
 * the part on load.
 */
function checkInvalidAutoElement(ctx: ValidationContext, path: string, root: XmlElement): void {
  for (const _el of collectDescendantsLocal(root, "auto")) {
    if (ctx.reporter.capped) {
      return;
    }
    ctx.reporter.error(
      "chartEx-invalid-auto-element",
      `${path}: <cx:auto/> element is not in the chartEx schema. ` +
        `Auto binning is expressed by omitting both <cx:binSize> and <cx:binCount>.`,
      path
    );
  }
}

/**
 * `<cx:paretoLine>` is not a child of `<cx:layoutPr>` in the schema.
 * A real pareto chart expresses the line as a second `<cx:series>` with
 * `layoutId="paretoLine"`. The mis-placed child made earlier Excel
 * builds reject the chartEx.
 */
function checkParetoLineInLayoutPr(ctx: ValidationContext, path: string, root: XmlElement): void {
  for (const lp of collectDescendantsLocal(root, "layoutPr")) {
    if (ctx.reporter.capped) {
      return;
    }
    const pl = findChildLocal(lp, "paretoLine");
    if (pl) {
      ctx.reporter.error(
        "chartEx-paretoLine-in-layoutPr",
        `${path}: <cx:paretoLine> is not a valid child of <cx:layoutPr>. ` +
          `Add a second <cx:series layoutId="paretoLine"/> instead.`,
        path
      );
    }
  }
}

/**
 * `<cx:title><cx:layout/></cx:title>` is schema-invalid. Title layout
 * lives in `extLst`-based extensions or (in some clients) `<cx:offset>`
 * — never as a direct `<cx:layout>` child.
 */
function checkTitleDirectLayoutChild(ctx: ValidationContext, path: string, root: XmlElement): void {
  for (const title of collectDescendantsLocal(root, "title")) {
    if (ctx.reporter.capped) {
      return;
    }
    const layout = findChildLocal(title, "layout");
    if (layout) {
      ctx.reporter.error(
        "chartEx-title-direct-layout",
        `${path}: <cx:title> has a direct <cx:layout> child. Title layout ` +
          `information belongs in extLst-based extensions.`,
        path
      );
    }
  }
}

/**
 * Concatenate direct text/cdata children of an element, ignoring any
 * nested elements. Useful for "typed element with stray text content"
 * detection where a nested element's text should NOT count as the
 * offending text form.
 */
function directTextContent(el: XmlElement): string {
  let out = "";
  for (const child of el.children) {
    if (child.type === "text" || child.type === "cdata") {
      out += child.value;
    }
  }
  return out;
}

// -----------------------------------------------------------------------------
// Tier-2 semantic checks
// -----------------------------------------------------------------------------

/**
 * `<cx:axis>` must declare its axis role — either via a structural
 * `<cx:catScaling>` / `<cx:valScaling>` CHILD element (the form Excel
 * itself emits; the role is inferred from which scaling child is
 * present) OR via legacy `pos` / `type` attributes. When NONE of
 * these are present, Excel's loader cannot disambiguate the axis
 * role and drops the whole `<cx:chartSpace>` on open, cascading
 * into "Removed Part: /xl/charts/chartExN.xml".
 *
 * Verified against Excel 2021's own output (`tmp/aaaaa.xlsx`,
 * `tmp/ttttt.xlsx`): every `<cx:axis>` it emits omits the
 * `pos` / `type` attributes and relies on the scaling child.
 */
function checkAxisPosAndType(ctx: ValidationContext, path: string, root: XmlElement): void {
  for (const axis of collectDescendantsLocal(root, "axis")) {
    if (ctx.reporter.capped) {
      return;
    }
    const pos = attrByLocalName(axis, "pos");
    const type = attrByLocalName(axis, "type");
    const id = attrByLocalName(axis, "id") ?? "?";
    // Accept either legacy attribute form OR the schema-native
    // `<cx:catScaling>` / `<cx:valScaling>` child.
    const hasCatScaling = axis.children.some(
      c => c.type === "element" && matchesLocal(c.name, "catScaling")
    );
    const hasValScaling = axis.children.some(
      c => c.type === "element" && matchesLocal(c.name, "valScaling")
    );
    if (pos === undefined && type === undefined && !hasCatScaling && !hasValScaling) {
      ctx.reporter.error(
        "chartEx-axis-missing-pos-and-type",
        `${path}: <cx:axis id="${id}"> has no role marker — emit either a ` +
          `<cx:catScaling>/<cx:valScaling> child (preferred, matches Excel) ` +
          `or a pos/type attribute. Excel's loader drops the chartEx otherwise.`,
        path
      );
    }
  }
}

/**
 * `<cx:f>` formulas must point at hidden defined names (the
 * `_xlchart.v1.0`, `_xlchart.v1.1`, … convention Excel itself uses),
 * NOT directly at worksheet ranges. A bare `<cx:f>Sheet1!$A$1:$A$3</cx:f>`
 * is rejected on open with "Removed Part: /xl/drawings/drawingN.xml".
 *
 * Detection heuristic: a formula body that contains `!$` or `!` followed
 * by absolute cell references is a direct sheet reference. Defined-name
 * references are bare identifiers like `_xlchart.v1.0` (no `!`).
 */
function checkSeriesFDefinedName(ctx: ValidationContext, path: string, root: XmlElement): void {
  for (const f of collectDescendantsLocal(root, "f")) {
    if (ctx.reporter.capped) {
      return;
    }
    const formula = directTextContent(f).trim();
    if (formula === "") {
      continue;
    }
    // Heuristic for a direct sheet-qualified range:
    //   - Contains "!" (sheet qualifier), AND
    //   - Does NOT start with `_xl` or other defined-name prefix.
    if (!formula.includes("!")) {
      continue; // bare defined name like `_xlchart.v1.0`
    }
    if (formula.startsWith("_xlchart.") || formula.startsWith("_xlfn.")) {
      continue; // defined-name-qualified alias
    }
    // Looks like `Sheet1!$A$1:$A$3` or `'Some Name'!$A$1`.
    ctx.reporter.error(
      "chartEx-f-uses-direct-range-not-defined-name",
      `${path}: <cx:f>${formula}</cx:f> points at a worksheet range directly. ` +
        `ChartEx requires an indirection through hidden defined names ` +
        `(e.g. _xlchart.v1.0) — otherwise Excel 2016+ drops the chartEx part on load.`,
      path
    );
  }
}

/**
 * Waterfall charts must have `layoutPr.subtotals` on their series —
 * even empty `<cx:subtotals/>` is meaningful: it marks the plot as
 * subtotals-aware. Without it Excel falls back to generic series
 * rendering and has been observed to reject the chartEx as malformed
 * at load time.
 */
function checkWaterfallSubtotals(ctx: ValidationContext, path: string, root: XmlElement): void {
  for (const series of collectDescendantsLocal(root, "series")) {
    if (ctx.reporter.capped) {
      return;
    }
    const layoutId = attrByLocalName(series, "layoutId");
    if (layoutId !== "waterfall") {
      continue;
    }
    const layoutPr = findChildLocal(series, "layoutPr");
    const hasSubtotals = !!layoutPr && findChildLocal(layoutPr, "subtotals") !== undefined;
    if (!hasSubtotals) {
      ctx.reporter.error(
        "chartEx-waterfall-missing-subtotals",
        `${path}: waterfall series has no <cx:layoutPr><cx:subtotals/> marker. ` +
          `Emit the element (even empty) so Excel renders the series as ` +
          `waterfall-aware instead of rejecting the chartEx.`,
        path
      );
    }
  }
}

function chartRelsPath(chartPath: string): string {
  const slash = chartPath.lastIndexOf("/");
  const dir = slash >= 0 ? chartPath.slice(0, slash) : "";
  const name = slash >= 0 ? chartPath.slice(slash + 1) : chartPath;
  return dir ? `${dir}/_rels/${name}.rels` : `_rels/${name}.rels`;
}
