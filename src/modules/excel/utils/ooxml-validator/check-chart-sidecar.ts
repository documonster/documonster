/**
 * ChartEx sidecar completeness check.
 *
 * Every ChartEx `xl/charts/chartExN.xml` MUST ship with a chartStyle
 * sidecar (`xl/charts/styleN.xml` or `xl/charts/styleExN.xml`) and a
 * chartColorStyle sidecar (`xl/charts/colorsN.xml` or
 * `xl/charts/colorsExN.xml`), linked from `chartExN.xml.rels`. Without
 * them Excel 2016+ discards the chartEx on load ("Removed Part:
 * /xl/drawings/drawingN.xml (Drawing shape)").
 *
 * Both sidecars are STUB-REJECTED **for chartEx** — an id-only
 * skeleton such as `<cs:chartStyle id="395"/>` (no child elements)
 * makes Excel treat the sidecar as malformed and drop the parent
 * chartEx + drawing. Excel authoring tools emit the FULL content: ~40
 * child elements for chartStyle (axisTitle, dataPoint, dataLabel, …)
 * and 6 `schemeClr` entries each with 9 `variation` siblings for
 * chartColorStyle.
 *
 * CLASSIC charts (`xl/charts/chartN.xml`) can share the chartStyle
 * namespace but Excel is lenient there — classic chart style stubs
 * are accepted. We therefore only enforce the no-stub rule on
 * sidecars **referenced by a `chartExN.xml` part** via its rels.
 *
 * The precise content is too intricate to validate without baking
 * ~1500 lines of template XML into this module. Instead we detect the
 * two high-signal stub patterns:
 *
 *   - chartStyle root element has zero element children (pure stub).
 *   - chartColorStyle root element has zero element children (pure stub).
 *
 * A rich sidecar trivially has dozens of children so the zero-child
 * case is a strong writer-bug signature.
 */

import type { ValidationContext } from "./context";
import { resolveRelTarget } from "./path-utils";

const CHARTEX_PART_RE = /^xl\/charts\/chartEx\d+\.xml$/;

const CHARTSTYLE_REL_TYPE = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
const CHARTCOLORS_REL_TYPE =
  "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";

export function checkChartSidecars(ctx: ValidationContext): void {
  // Phase 1: collect all sidecar paths that are referenced from a
  // chartEx part. Only these warrant the "stub is rejection" rule.
  const chartExStylePaths = new Set<string>();
  const chartExColorsPaths = new Set<string>();
  for (const [path, entry] of ctx.files()) {
    if (entry.type === "directory" || !CHARTEX_PART_RE.test(path)) {
      continue;
    }
    const relsPath = chartRelsPath(path);
    if (!ctx.has(relsPath)) {
      continue;
    }
    const rels = ctx.readRels(relsPath);
    for (const rel of rels.rels) {
      if (rel.targetMode === "External") {
        continue;
      }
      const resolved = resolveRelTarget(relsPath, rel.target);
      if (rel.type === CHARTSTYLE_REL_TYPE) {
        chartExStylePaths.add(resolved);
      } else if (rel.type === CHARTCOLORS_REL_TYPE) {
        chartExColorsPaths.add(resolved);
      }
    }
  }

  // Phase 2: for every sidecar referenced by a chartEx, flag stub form.
  for (const path of chartExStylePaths) {
    if (ctx.reporter.capped) {
      return;
    }
    checkStubRoot(ctx, path, "chartEx-chartStyle-stub-form", "chartStyle");
  }
  for (const path of chartExColorsPaths) {
    if (ctx.reporter.capped) {
      return;
    }
    checkStubRoot(ctx, path, "chartEx-chartColorStyle-stub-form", "chartColorStyle");
  }
}

function chartRelsPath(chartPath: string): string {
  const slash = chartPath.lastIndexOf("/");
  const dir = slash >= 0 ? chartPath.slice(0, slash) : "";
  const name = slash >= 0 ? chartPath.slice(slash + 1) : chartPath;
  return dir ? `${dir}/_rels/${name}.rels` : `_rels/${name}.rels`;
}

function checkStubRoot(
  ctx: ValidationContext,
  path: string,
  kind: "chartEx-chartStyle-stub-form" | "chartEx-chartColorStyle-stub-form",
  label: string
): void {
  if (!ctx.has(path)) {
    return; // missing-target already reported by the rels checker.
  }
  const dom = ctx.readDom(path);
  if (!dom) {
    return;
  }
  const root = dom.root;
  const hasElementChildren = root.children.some(c => c.type === "element");
  if (!hasElementChildren) {
    ctx.reporter.error(
      kind,
      `${path}: ${label} sidecar referenced by a chartEx part is an id-only stub ` +
        `(root element has no children). Excel 2016+ rejects chartEx stub sidecars and ` +
        `drops the parent chartEx + drawing. Emit the full Microsoft-authored default content.`,
      path
    );
  }
}
