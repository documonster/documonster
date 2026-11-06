import type {
  AddSparklineGroupOptions as SparklineAddOptions,
  SparklineGroup as SparklineGroupModel
} from "@excel/core/sparkline";
import { addSparklineGroup, getSparklineGroups, removeSparklineGroup } from "@excel/core/worksheet";
/**
 * `Sparkline` namespace surface.
 * `import { Sparkline } from "documonster/excel"` → `Sparkline.add(ws, opts)`.
 */
import { checkWorksheet } from "@excel/core/worksheet-core";
import type { WorksheetData } from "@excel/core/worksheet-core";
import { checkListTarget } from "@excel/surface/argument-checks";
import { argumentChecks, checkSchema } from "@excel/utils/arguments";
import { s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";

/**
 * The sheet's sparkline groups, in the order they were added: the groups themselves, not copies, which {@link remove}
 * takes. An edit made to one directly is not checked as {@link add} checks a group.
 */
export function list(ws: WorksheetData): SparklineGroupModel[] {
  checkWorksheet(ws);
  return getSparklineGroups(ws);
}

/** Remove a sparkline group — one {@link list} listed, or its index in that list. Returns `true` if one was removed. */
export function remove(ws: WorksheetData, groupOrIndex: SparklineGroupModel | number): boolean {
  checkWorksheet(ws);
  checkListTarget("Sparkline.remove", "groupOrIndex", groupOrIndex, "a sparkline group");
  return removeSparklineGroup(ws, groupOrIndex);
}

const flag = () => s.boolean().optional();
// An empty string has always meant "no colour" here.
const color = () =>
  s
    .custom<string>(
      v => v === "" || (typeof v === "string" && /^#?(?:[0-9A-Fa-f]{2})?[0-9A-Fa-f]{6}$/.test(v)),
      'a hex colour such as "#FF0000" or "FFFF0000"'
    )
    .optional();
const axisType = () => s.literal("individual", "group", "custom").optional();

const addOptionsSchema = s.object({
  type: s.literal("line", "column", "stacked"),
  sparklines: s.array(
    s.object({
      // A range or a defined name, so only its type is checked.
      dataRef: s.string({ nonEmpty: true }),
      // Checked as a reference below, refused with an `ExcelInvalidAddressError`.
      cellRef: s.custom<string>(v => v !== undefined && v !== null, "given")
    }),
    { minLength: 1, expected: "an array of at least 1 item(s)" }
  ),
  lineWeight: s.number({ finite: true, min: 0 }).optional(),
  markers: flag(),
  high: flag(),
  low: flag(),
  first: flag(),
  last: flag(),
  negative: flag(),
  lineColor: color(),
  negativeColor: color(),
  axisColor: color(),
  markerColor: color(),
  highColor: color(),
  lowColor: color(),
  firstColor: color(),
  lastColor: color(),
  minAxisType: axisType(),
  maxAxisType: axisType(),
  manualMin: s.number({ finite: true }).optional(),
  manualMax: s.number({ finite: true }).optional(),
  displayXAxis: flag(),
  rightToLeft: flag(),
  displayEmptyCellsAs: s.literal("gap", "zero", "span").optional(),
  dateAxis: s.string().optional()
});
const _addOptionsMatch: SchemaMatches<typeof addOptionsSchema, SparklineAddOptions> = true;
void _addOptionsMatch;

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
export function add(ws: WorksheetData, options: SparklineAddOptions): SparklineGroupModel {
  checkWorksheet(ws);
  checkSchema("Sparkline.add", "options", addOptionsSchema, options);
  const check = argumentChecks("Sparkline.add");
  options.sparklines.forEach((item, i) =>
    check.reference(`options.sparklines[${i}].cellRef`, item.cellRef)
  );
  return addSparklineGroup(ws, options);
}

// The flat names are the public ones; these namespace aliases repeated them (and `SparklineGroup` its namespace noun).
/** @deprecated Use the flat `AddSparklineGroupOptions` type. Removed in 1.0. */
export type AddSparklineGroupOptions = SparklineAddOptions;
/** @deprecated Use the flat `SparklineGroup` type. Removed in 1.0. */
export type SparklineGroup = SparklineGroupModel;
