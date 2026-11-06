/**
 * Checks for the `Chart` surface's arguments that the chart builder's own validation does not reach: the placement
 * every `add*` takes last, a series' name and categories, and the legend. Kept out of `chart.ts`, because everything
 * that file exports is public API.
 */
import type { ArgumentChecks } from "@excel/utils/arguments";
import { argumentChecks } from "@excel/utils/arguments";

/** A series' `name` (text, or `{ formula }`) and `categories` (a reference, or a structured data source). */
export function checkChartSeries(where: string, prefix: string, series: unknown): void {
  const check = argumentChecks(where);
  const at = (field: string) => (prefix === "" ? field : `${prefix}.${field}`);
  if (!check.record(prefix === "" ? "series" : prefix, series)) {
    return;
  }
  const name = series.name;
  if (name !== undefined && typeof name !== "string") {
    if (check.record(at("name"), name)) {
      check.required(at("name.formula"), name.formula);
      check.string(at("name.formula"), name.formula, { nonEmpty: true });
    }
  }
  const categories = series.categories;
  if (categories !== undefined) {
    if (typeof categories === "string") {
      check.string(at("categories"), categories, { nonEmpty: true });
    } else {
      check.record(at("categories"), categories);
    }
  }
}

/** A corner of a chart anchor: an `A1` cell, or 0-based `{ col, row }` (fractions are offsets into the cell). */
function checkCorner(check: ArgumentChecks, field: string, corner: unknown): void {
  check.required(field, corner);
  if (typeof corner === "string") {
    check.reference(field, corner);
  } else if (check.record(field, corner)) {
    check.required(`${field}.col`, corner.col);
    check.number(`${field}.col`, corner.col, { min: 0, max: 16384 });
    check.required(`${field}.row`, corner.row);
    check.number(`${field}.row`, corner.row, { min: 0, max: 1048576 });
  }
}

/** A chart's extent, in EMU — `{ cx, cy }`, not an image's `{ width, height }` in pixels. */
function checkChartExtent(check: ArgumentChecks, value: unknown): void {
  if (check.record("range.ext", value)) {
    check.required("range.ext.cx", value.cx);
    check.number("range.ext.cx", value.cx, { min: 0, integer: true });
    check.required("range.ext.cy", value.cy);
    check.number("range.ext.cy", value.cy, { min: 0, integer: true });
  }
}

/** A chart's placement: `"E1:L14"`, a two- or one-cell anchor, or an absolute position, all in the grid. */
export function checkChartRange(where: string, range: unknown): void {
  const check = argumentChecks(where);
  check.required("range", range);
  if (typeof range === "string") {
    check.reference("range", range, { range: true });
    return;
  }
  if (!check.record("range", range)) {
    return;
  }
  check.oneOf("range.editAs", range.editAs, ["oneCell", "twoCell", "absolute"]);
  // An extent without `cx` / `cy` — an image's `{ width, height }` — was accepted and written as
  // `<xdr:ext cx="undefined" cy="undefined"/>`, a part Excel refuses to open.
  if ("pos" in range) {
    if (check.record("range.pos", range.pos)) {
      check.required("range.pos.x", range.pos.x);
      check.number("range.pos.x", range.pos.x, { min: 0, integer: true });
      check.required("range.pos.y", range.pos.y);
      check.number("range.pos.y", range.pos.y, { min: 0, integer: true });
    }
    check.required("range.ext", range.ext);
    checkChartExtent(check, range.ext);
    return;
  }
  checkCorner(check, "range.tl", range.tl);
  if (range.br !== undefined) {
    checkCorner(check, "range.br", range.br);
  } else {
    // Neither corner nor extent: it was read as a two-cell anchor and failed inside the anchor constructor.
    check.required("range.ext", range.ext);
  }
  if (range.ext !== undefined) {
    checkChartExtent(check, range.ext);
  }
}

/**
 * The arguments of an `add*` call: the placement, which every one takes last, and — where the options carry their
 * series — each series' name and categories.
 */
export function checkPlacement(where: string, args: readonly unknown[]): void {
  checkChartRange(where, args[args.length - 1]);
  const options = args[args.length - 2] as { series?: unknown } | undefined;
  if (typeof options === "object" && options !== null && Array.isArray(options.series)) {
    options.series.forEach((series, i) => checkChartSeries(where, `options.series[${i}]`, series));
  }
}

/** A legend: its position and alignment are the format's enumerations, `overlay` a boolean. */
export function checkLegend(where: string, field: string, legend: unknown): void {
  const check = argumentChecks(where);
  if (!check.record(field, legend)) {
    return;
  }
  check.oneOf(`${field}.legendPos`, legend.legendPos, ["b", "l", "r", "t", "tr"]);
  check.oneOf(`${field}.align`, legend.align, ["ctr", "l", "r", "t", "b"]);
  check.boolean(`${field}.overlay`, legend.overlay);
}
