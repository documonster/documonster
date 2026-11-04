import type { AddSparklineGroupOptions, SparklineGroup } from "@excel/core/sparkline";
import { addSparklineGroup, getSparklineGroups, removeSparklineGroup } from "@excel/core/worksheet";
/**
 * `Sparkline` namespace surface.
 * `import { Sparkline } from "documonster/excel"` → `Sparkline.add(ws, opts)`.
 */
import { checkWorksheet } from "@excel/core/worksheet-core";
import type { WorksheetData } from "@excel/core/worksheet-core";
import { checkListTarget } from "@excel/surface/argument-checks";
import { argumentChecks } from "@excel/utils/arguments";

/** The sheet's sparkline groups, in the order they were added. */
export function list(ws: WorksheetData): SparklineGroup[] {
  checkWorksheet(ws);
  return getSparklineGroups(ws);
}

/** Remove a sparkline group — one {@link list} listed, or its index in that list. Returns `true` if one was removed. */
export function remove(ws: WorksheetData, groupOrIndex: SparklineGroup | number): boolean {
  checkWorksheet(ws);
  checkListTarget("Sparkline.remove", "groupOrIndex", groupOrIndex, "a sparkline group");
  return removeSparklineGroup(ws, groupOrIndex);
}

const COLOR_OPTIONS = [
  "lineColor",
  "negativeColor",
  "axisColor",
  "markerColor",
  "highColor",
  "lowColor",
  "firstColor",
  "lastColor"
] as const;
const AXIS_TYPES = ["individual", "group", "custom"];
const HEX_COLOR = /^#?(?:[0-9A-Fa-f]{2})?[0-9A-Fa-f]{6}$/;

/**
 * Add a group of sparklines to the sheet, and return it.
 *
 * Each sparkline pairs a data range (`"Sheet1!B2:G2"`) with the single cell it is drawn in (`"H2"`); colours are hex
 * (`"#FF0000"` or `"FFFF0000"`). The options are checked here, before anything is added: a type, axis type or
 * colour outside the format's vocabulary was written as given and Excel repaired the file, and a group without
 * `sparklines` failed only when the workbook was written.
 *
 * The group keeps copies of the `sparklines` given. The returned group is the one the sheet writes: editing it changes
 * the sheet without these checks.
 */
export function add(ws: WorksheetData, options: AddSparklineGroupOptions): SparklineGroup {
  checkWorksheet(ws);
  const check = argumentChecks("Sparkline.add");
  check.required("options", options);
  if (check.record("options", options)) {
    check.required("options.type", options.type);
    check.oneOf("options.type", options.type, ["line", "column", "stacked"]);
    check.required("options.sparklines", options.sparklines);
    if (check.array("options.sparklines", options.sparklines, { min: 1 })) {
      options.sparklines.forEach((item, i) => {
        check.required(`options.sparklines[${i}]`, item);
        if (check.record(`options.sparklines[${i}]`, item)) {
          check.required(`options.sparklines[${i}].dataRef`, item.dataRef);
          // A range or a defined name, so only its type is checked.
          check.string(`options.sparklines[${i}].dataRef`, item.dataRef, { nonEmpty: true });
          check.required(`options.sparklines[${i}].cellRef`, item.cellRef);
          check.reference(`options.sparklines[${i}].cellRef`, item.cellRef);
        }
      });
    }
    check.number("options.lineWeight", options.lineWeight, { min: 0 });
    for (const flag of [
      "markers",
      "high",
      "low",
      "first",
      "last",
      "negative",
      "displayXAxis",
      "rightToLeft"
    ] as const) {
      check.boolean(`options.${flag}`, options[flag]);
    }
    for (const color of COLOR_OPTIONS) {
      const value = options[color];
      // An empty string has always meant "no colour" here.
      if (
        value !== undefined &&
        value !== "" &&
        (typeof value !== "string" || !HEX_COLOR.test(value))
      ) {
        check.fail(`options.${color}`, 'a hex colour such as "#FF0000" or "FFFF0000"', value);
      }
    }
    check.oneOf("options.minAxisType", options.minAxisType, AXIS_TYPES);
    check.oneOf("options.maxAxisType", options.maxAxisType, AXIS_TYPES);
    check.number("options.manualMin", options.manualMin);
    check.number("options.manualMax", options.manualMax);
    check.oneOf("options.displayEmptyCellsAs", options.displayEmptyCellsAs, [
      "gap",
      "zero",
      "span"
    ]);
    check.string("options.dateAxis", options.dateAxis);
  }
  return addSparklineGroup(ws, options);
}

export type { AddSparklineGroupOptions, SparklineGroup } from "@excel/core/sparkline";
