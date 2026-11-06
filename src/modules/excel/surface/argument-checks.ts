/**
 * Checks shared by several surface namespaces. Kept out of the namespace files themselves, because everything those
 * export is public API.
 */
import { isNoteData } from "@excel/core/note";
import { getImage } from "@excel/core/workbook-core";
import type { HeaderFooterImageOptions } from "@excel/core/worksheet";
import type { WorksheetData } from "@excel/core/worksheet-core";
import { getSheetWorkbook } from "@excel/core/worksheet-core";
import { ExcelImageError } from "@excel/errors";
import type {
  DataValidationRule,
  DataValidationWithFormulae,
  WatermarkOptions
} from "@excel/types";
import type { ArgumentChecks } from "@excel/utils/arguments";
import { argumentChecks, checkSchema } from "@excel/utils/arguments";
import { checkStyleFacet } from "@excel/utils/style-checks";
import { s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";

const MAX_ROW = 1048576;
const MAX_COL = 16384;

/** The workbook must have image `imageId` — a picture that names no image is written with nothing to show. */
export function checkImageId(where: string, ws: WorksheetData, imageId: unknown): void {
  if (
    (typeof imageId !== "string" && typeof imageId !== "number") ||
    getImage(getSheetWorkbook(ws), imageId) === undefined
  ) {
    throw new ExcelImageError(
      `${where}: the workbook has no image ${JSON.stringify(imageId)}; add it with Image.add first.`
    );
  }
}

const DATA_VALIDATION_TYPES = [
  "any",
  "list",
  "whole",
  "decimal",
  "date",
  "time",
  "textLength",
  "custom"
] as const;

/** The keys every rule takes, in the order they are checked. */
const dataValidationBaseShape = {
  // The file format's enumeration; the published type says `string`, which is what it reads back.
  errorStyle: s.member<string>(["stop", "warning", "information"]).optional(),
  allowBlank: s.boolean().optional(),
  showErrorMessage: s.boolean().optional(),
  showInputMessage: s.boolean().optional(),
  showDropDown: s.boolean().optional(),
  error: s.string().optional(),
  errorTitle: s.string().optional(),
  prompt: s.string().optional(),
  promptTitle: s.string().optional()
};

const dataValidationWithFormulaeSchema = s.object({
  // Settled by the variant's tag. `time` is not in the published union, but readers return it and both writers
  // write it back (see `DataValidationWithFormulae.type`), so it is a tag too.
  type: s.unknown<DataValidationWithFormulae["type"]>(),
  operator: s
    .literal(
      "between",
      "notBetween",
      "equal",
      "notEqual",
      "greaterThan",
      "lessThan",
      "greaterThanOrEqual",
      "lessThanOrEqual"
    )
    .optional(),
  ...dataValidationBaseShape,
  // A rule with no bounds is accepted, as Excel accepts it.
  formulae: s
    .array(
      s.custom<string | number | Date>(
        formula =>
          (typeof formula === "string" && formula.trim() !== "") ||
          (typeof formula === "number" && Number.isFinite(formula)) ||
          (formula instanceof Date && !Number.isNaN(formula.getTime())),
        "a formula, a finite number or a valid Date"
      )
    )
    .refine(formulae => formulae.length <= 2, "an array of 0 to 2 items")
    .orUndefined()
});

const dataValidationAnySchema = s.object({
  type: s.literal("any"),
  ...dataValidationBaseShape
});

const dataValidationRuleSchema = s.variant(
  "type",
  {
    any: dataValidationAnySchema,
    list: dataValidationWithFormulaeSchema,
    whole: dataValidationWithFormulaeSchema,
    decimal: dataValidationWithFormulaeSchema,
    date: dataValidationWithFormulaeSchema,
    time: dataValidationWithFormulaeSchema,
    textLength: dataValidationWithFormulaeSchema,
    custom: dataValidationWithFormulaeSchema
  },
  { unknownTag: `one of ${DATA_VALIDATION_TYPES.map(t => JSON.stringify(t)).join(", ")}` }
);
const _dataValidationRuleMatch: SchemaMatches<typeof dataValidationRuleSchema, DataValidationRule> =
  true;
void _dataValidationRuleMatch;

/**
 * Check a validation rule's vocabulary: the type, operator and error style are the file format's enumerations, the
 * flags are booleans, the messages strings, and the bounds at most two formulas, numbers or dates. A rule with no
 * bounds is accepted, as Excel accepts it. An `any` rule has no operator or bounds, and any it is given are not
 * written. `undefined` removes the cell's rule, as it always has.
 */
export function checkDataValidationRule(where: string, rule: unknown): void {
  checkSchema(where, "rule", dataValidationRuleSchema.optional(), rule);
}

/**
 * Check a comment: `undefined`, a note (text or configuration), or `{ note, author }` — as `Cell.setComment` takes it.
 */
export function checkComment(where: string, argument: string, value: unknown): void {
  if (value !== undefined && (typeof value !== "object" || value === null)) {
    argumentChecks(where).fail(argument, "a note, a note configuration or undefined", value);
  }
  if (isNoteData(value)) {
    checkNote(where, `${argument}.note`, value.note);
    if (value.author !== undefined && value.author !== null) {
      argumentChecks(where).string(`${argument}.author`, value.author);
    }
  } else {
    checkNote(where, argument, value);
  }
}

const POSITIONING = ["twoCells", "oneCells", "absolute"];
// ST_TrueFalseBlank, which `<x:Locked>` and `<x:LockText>` hold.
const VML_BOOLEAN = ["True", "False", "true", "false", "t", "f", ""];

/**
 * Check a note's text or configuration: plain text, or a {@link NoteConfig} whose texts are strings (with an
 * optional font, checked as a cell font is), whose size is a non-negative number of points and whose positioning
 * and protection are the VML vocabulary. A note that is neither failed only when the workbook was written, with a
 * `TypeError` from inside the comment writer.
 */
export function checkNote(where: string, argument: string, note: unknown): void {
  const check = argumentChecks(where);
  if (note === undefined || typeof note === "string") {
    return;
  }
  if (!check.record(argument, note)) {
    return;
  }
  if (check.array(`${argument}.texts`, note.texts)) {
    note.texts.forEach((text, i) => {
      const field = `${argument}.texts[${i}]`;
      check.required(field, text);
      if (check.record(field, text)) {
        check.required(`${field}.text`, text.text);
        check.string(`${field}.text`, text.text);
        checkStyleFacet(where, "font", text.font, `${field}.font`);
      }
    });
  }
  if (check.record(`${argument}.margins`, note.margins)) {
    check.oneOf(`${argument}.margins.insetmode`, note.margins.insetmode, ["auto", "custom"]);
    if (check.array(`${argument}.margins.inset`, note.margins.inset, { max: 4 })) {
      note.margins.inset.forEach((inset, i) =>
        check.number(`${argument}.margins.inset[${i}]`, inset, { min: 0 })
      );
    }
  }
  if (check.record(`${argument}.protection`, note.protection)) {
    check.oneOf(`${argument}.protection.locked`, note.protection.locked, VML_BOOLEAN);
    check.oneOf(`${argument}.protection.lockText`, note.protection.lockText, VML_BOOLEAN);
  }
  check.oneOf(`${argument}.editAs`, note.editAs, POSITIONING);
  check.string(`${argument}.anchor`, note.anchor);
  check.number(`${argument}.width`, note.width, { min: 0 });
  check.number(`${argument}.height`, note.height, { min: 0 });
}

/** A corner of an auto-filter: an `A1` cell, or 1-based `{ row, col }`. */
function checkFilterCorner(check: ArgumentChecks, field: string, corner: unknown): void {
  check.required(field, corner);
  if (typeof corner === "string") {
    check.reference(field, corner);
  } else if (check.record(field, corner)) {
    check.number(`${field}.row`, corner.row, { min: 1, max: MAX_ROW, integer: true });
    check.number(`${field}.col`, corner.col, { min: 1, max: MAX_COL, integer: true });
  }
}

/** Check an auto-filter range: an `A1:C4` string or a `{ from, to }` pair of this sheet's cells. */
export function checkAutoFilter(where: string, field: string, range: unknown): void {
  const check = argumentChecks(where);
  check.required(field, range);
  if (typeof range === "string") {
    check.reference(field, range, { range: true });
  } else if (check.record(field, range)) {
    checkFilterCorner(check, `${field}.from`, range.from);
    checkFilterCorner(check, `${field}.to`, range.to);
  }
}

const POSITIONS = ["LH", "CH", "RH", "LF", "CF", "RF"] as const;
const position = () => s.literal(...POSITIONS);
// Points, written into the VML shape's `style`, where `NaN` and a negative size are not lengths.
const points = () => s.number({ finite: true, min: 0 }).optional();
const applyTo = () => s.literal("all", "odd", "even", "first").optional();
// Checked against the workbook's images by `checkImageId`, refused with an `ExcelImageError`.
const imageId = () => s.unknown<string | number>();

/** `HeaderFooterImage.set`'s options. */
export const headerFooterImageOptionsSchema = s.object({
  imageId: imageId(),
  position: position(),
  width: points(),
  height: points(),
  applyTo: applyTo()
});
const _headerFooterImageOptionsMatch: SchemaMatches<
  typeof headerFooterImageOptionsSchema,
  HeaderFooterImageOptions
> = true;
void _headerFooterImageOptionsMatch;

/** `Watermark.set`'s options. */
export const watermarkOptionsSchema = s.object({
  imageId: imageId(),
  mode: s.literal("overlay", "header").optional(),
  opacity: s.number({ finite: true, min: 0, max: 1 }).optional(),
  headerWidth: points(),
  headerHeight: points(),
  position: position().optional(),
  applyTo: applyTo()
});
const _watermarkOptionsMatch: SchemaMatches<typeof watermarkOptionsSchema, WatermarkOptions> = true;
void _watermarkOptionsMatch;

/** A header/footer section (`"LH"` … `"RF"`), as `HeaderFooterImage.remove` takes it. */
export function checkHeaderPosition(where: string, argument: string, value: unknown): void {
  checkSchema(where, argument, position(), value);
}

/**
 * What a collection's `remove(ws, target)` takes: an item its `list` returned, or a whole-number position in that
 * list. Anything else — a name, a reference string, `1.5`, `null` — named nothing the call can find, and used to
 * answer `false` as if the item were simply absent. An object that is not in the list (another sheet's, or one
 * already removed) still answers `false`.
 */
export function checkListTarget(where: string, field: string, target: unknown, item: string): void {
  if (typeof target === "number" ? Number.isInteger(target) : isPlainHandle(target)) {
    return;
  }
  argumentChecks(where).fail(field, `${item} or its position (a whole number) in the list`, target);
}

function isPlainHandle(value: unknown): boolean {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}
