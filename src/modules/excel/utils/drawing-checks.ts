/**
 * Checks for what a drawing places on a sheet — an image's anchor and a shape's options — so a value the drawing part
 * cannot hold is refused where it was given rather than written as `NaN`, a negative extent or an unknown geometry.
 */
import { ExcelArgumentError } from "@excel/errors";
import type { AddImageRange, AddShapeOptions, ShapeType } from "@excel/types";
import type { ArgumentChecks } from "@excel/utils/arguments";
import { argumentChecks, checkSchema } from "@excel/utils/arguments";
import { colCache } from "@excel/utils/col-cache";
import { s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";

/** ST_ShapeType of ECMA-376 `dml-main.xsd` (Transitional), in schema order. */
const SHAPE_TYPES: ReadonlySet<string> = new Set([
  "line",
  "lineInv",
  "triangle",
  "rtTriangle",
  "rect",
  "diamond",
  "parallelogram",
  "trapezoid",
  "nonIsoscelesTrapezoid",
  "pentagon",
  "hexagon",
  "heptagon",
  "octagon",
  "decagon",
  "dodecagon",
  "star4",
  "star5",
  "star6",
  "star7",
  "star8",
  "star10",
  "star12",
  "star16",
  "star24",
  "star32",
  "roundRect",
  "round1Rect",
  "round2SameRect",
  "round2DiagRect",
  "snipRoundRect",
  "snip1Rect",
  "snip2SameRect",
  "snip2DiagRect",
  "plaque",
  "ellipse",
  "teardrop",
  "homePlate",
  "chevron",
  "pieWedge",
  "pie",
  "blockArc",
  "donut",
  "noSmoking",
  "rightArrow",
  "leftArrow",
  "upArrow",
  "downArrow",
  "stripedRightArrow",
  "notchedRightArrow",
  "bentUpArrow",
  "leftRightArrow",
  "upDownArrow",
  "leftUpArrow",
  "leftRightUpArrow",
  "quadArrow",
  "leftArrowCallout",
  "rightArrowCallout",
  "upArrowCallout",
  "downArrowCallout",
  "leftRightArrowCallout",
  "upDownArrowCallout",
  "quadArrowCallout",
  "bentArrow",
  "uturnArrow",
  "circularArrow",
  "leftCircularArrow",
  "leftRightCircularArrow",
  "curvedRightArrow",
  "curvedLeftArrow",
  "curvedUpArrow",
  "curvedDownArrow",
  "swooshArrow",
  "cube",
  "can",
  "lightningBolt",
  "heart",
  "sun",
  "moon",
  "smileyFace",
  "irregularSeal1",
  "irregularSeal2",
  "foldedCorner",
  "bevel",
  "frame",
  "halfFrame",
  "corner",
  "diagStripe",
  "chord",
  "arc",
  "leftBracket",
  "rightBracket",
  "leftBrace",
  "rightBrace",
  "bracketPair",
  "bracePair",
  "straightConnector1",
  "bentConnector2",
  "bentConnector3",
  "bentConnector4",
  "bentConnector5",
  "curvedConnector2",
  "curvedConnector3",
  "curvedConnector4",
  "curvedConnector5",
  "callout1",
  "callout2",
  "callout3",
  "accentCallout1",
  "accentCallout2",
  "accentCallout3",
  "borderCallout1",
  "borderCallout2",
  "borderCallout3",
  "accentBorderCallout1",
  "accentBorderCallout2",
  "accentBorderCallout3",
  "wedgeRectCallout",
  "wedgeRoundRectCallout",
  "wedgeEllipseCallout",
  "cloudCallout",
  "cloud",
  "ribbon",
  "ribbon2",
  "ellipseRibbon",
  "ellipseRibbon2",
  "leftRightRibbon",
  "verticalScroll",
  "horizontalScroll",
  "wave",
  "doubleWave",
  "plus",
  "flowChartProcess",
  "flowChartDecision",
  "flowChartInputOutput",
  "flowChartPredefinedProcess",
  "flowChartInternalStorage",
  "flowChartDocument",
  "flowChartMultidocument",
  "flowChartTerminator",
  "flowChartPreparation",
  "flowChartManualInput",
  "flowChartManualOperation",
  "flowChartConnector",
  "flowChartPunchedCard",
  "flowChartPunchedTape",
  "flowChartSummingJunction",
  "flowChartOr",
  "flowChartCollate",
  "flowChartSort",
  "flowChartExtract",
  "flowChartMerge",
  "flowChartOfflineStorage",
  "flowChartOnlineStorage",
  "flowChartMagneticTape",
  "flowChartMagneticDisk",
  "flowChartMagneticDrum",
  "flowChartDisplay",
  "flowChartDelay",
  "flowChartAlternateProcess",
  "flowChartOffpageConnector",
  "actionButtonBlank",
  "actionButtonHome",
  "actionButtonHelp",
  "actionButtonInformation",
  "actionButtonForwardNext",
  "actionButtonBackPrevious",
  "actionButtonEnd",
  "actionButtonBeginning",
  "actionButtonReturn",
  "actionButtonDocument",
  "actionButtonSound",
  "actionButtonMovie",
  "gear6",
  "gear9",
  "funnel",
  "mathPlus",
  "mathMinus",
  "mathMultiply",
  "mathDivide",
  "mathEqual",
  "mathNotEqual",
  "cornerTabs",
  "squareTabs",
  "plaqueTabs",
  "chartX",
  "chartStar",
  "chartPlus"
]);

/** ST_LineWidth: 0 to 20116800 EMU, which is 1584 points. */
const MAX_LINE_WIDTH_PT = 1584;

const GRID = { col: 16384, row: 1048576 };
// `srgbClr` takes six hex digits; an eight-digit ARGB has its alpha dropped by the writer.
const HEX_COLOR = /^#?(?:[0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$/;

/**
 * A corner, returned as 0-based grid coordinates: `{ col, row }` (0-based; a fraction places it inside the cell), an
 * anchor model (`nativeCol` / `nativeRow` with EMU offsets), or an A1 cell — the cell's top-left corner for `tl`, its
 * bottom-right corner for `br`, as the image code reads them.
 */
function corner(
  check: ArgumentChecks,
  field: string,
  value: unknown,
  edge: "tl" | "br"
): { col: number; row: number } | undefined {
  check.required(field, value);
  if (typeof value === "string") {
    check.reference(field, value);
    const { col, row } = colCache.decodeAddress(value);
    return edge === "tl" ? { col: col - 1, row: row - 1 } : { col, row };
  }
  if (!check.record(field, value)) {
    return undefined;
  }
  if (value.col === undefined && value.row === undefined && "nativeCol" in value) {
    check.number(`${field}.nativeCol`, value.nativeCol, { min: 0, max: GRID.col, integer: true });
    check.number(`${field}.nativeRow`, value.nativeRow, { min: 0, max: GRID.row, integer: true });
    check.number(`${field}.nativeColOff`, value.nativeColOff, { min: 0, integer: true });
    check.number(`${field}.nativeRowOff`, value.nativeRowOff, { min: 0, integer: true });
    return { col: (value.nativeCol as number) ?? 0, row: (value.nativeRow as number) ?? 0 };
  }
  check.required(`${field}.col`, value.col);
  check.number(`${field}.col`, value.col, { min: 0, max: GRID.col });
  check.required(`${field}.row`, value.row);
  check.number(`${field}.row`, value.row, { min: 0, max: GRID.row });
  return { col: value.col as number, row: value.row as number };
}

function extent(check: ArgumentChecks, field: string, value: unknown): void {
  if (check.record(field, value)) {
    check.required(`${field}.width`, value.width);
    check.number(`${field}.width`, value.width, { min: 0 });
    check.required(`${field}.height`, value.height);
    check.number(`${field}.height`, value.height, { min: 0 });
  }
}

/**
 * Check where an image or shape goes: a range (`"B2:D5"`), `{ tl, br }` with the bottom-right corner not above or
 * left of the top-left one, `{ tl, ext }`, or `{ pos, ext }` in pixels.
 */
export function checkDrawingRange(where: string, field: string, range: unknown): void {
  const check = argumentChecks(where);
  check.required(field, range);
  if (typeof range === "string") {
    check.reference(field, range, { range: true });
    return;
  }
  if (!check.record(field, range)) {
    return;
  }
  if ("pos" in range) {
    if (check.record(`${field}.pos`, range.pos)) {
      check.required(`${field}.pos.x`, range.pos.x);
      check.number(`${field}.pos.x`, range.pos.x, { min: 0 });
      check.required(`${field}.pos.y`, range.pos.y);
      check.number(`${field}.pos.y`, range.pos.y, { min: 0 });
    }
    check.required(`${field}.ext`, range.ext);
    extent(check, `${field}.ext`, range.ext);
    return;
  }
  const tl = corner(check, `${field}.tl`, range.tl, "tl");
  if (range.br !== undefined) {
    const br = corner(check, `${field}.br`, range.br, "br");
    if (tl && br && (br.col < tl.col || br.row < tl.row)) {
      throw new ExcelArgumentError(
        `${field}.br`,
        `${where}: ${field}.br must not be above or left of ${field}.tl.`
      );
    }
  } else {
    check.required(`${field}.ext`, range.ext);
  }
  extent(check, `${field}.ext`, range.ext);
  check.oneOf(`${field}.editAs`, range.editAs, ["oneCell", "twoCell", "absolute"]);
}

const hexColor = () =>
  s
    .custom<string>(
      v => typeof v === "string" && HEX_COLOR.test(v),
      'a hex colour such as "FF0000"'
    )
    .optional();

const shapeOptionsSchema = s.object({
  type: s
    .custom<ShapeType>(
      v => typeof v === "string" && SHAPE_TYPES.has(v),
      'a DrawingML preset geometry such as "rect" or "ellipse"'
    )
    .optional(),
  // Checked by `checkDrawingRange`, which an image's anchor shares.
  range: s.unknown<AddImageRange>(),
  fillColor: hexColor(),
  lineColor: hexColor(),
  lineWidth: s.number({ finite: true, min: 0, max: MAX_LINE_WIDTH_PT }).optional(),
  text: s.string().optional(),
  name: s.string().optional()
});
const _shapeOptionsMatch: SchemaMatches<typeof shapeOptionsSchema, AddShapeOptions> = true;
void _shapeOptionsMatch;

/**
 * Check a shape's options: a geometry from the DrawingML preset vocabulary, hex colours, a line width the format can
 * hold, and an area to sit in.
 */
export function checkShapeOptions(where: string, options: unknown): void {
  checkSchema(where, "options", shapeOptionsSchema, options);
  checkDrawingRange(where, "options.range", (options as AddShapeOptions).range);
}
