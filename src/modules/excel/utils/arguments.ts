/**
 * Argument checks for the public surface: a value that means nothing to the call is refused where it was passed,
 * naming the function and the argument, instead of being stored and then dropped or refused when the workbook is
 * written.
 *
 * An options object is checked against a schema (`@utils/schema`), pinned to its public interface with
 * `SchemaMatches` — see {@link at}. The single-value checks below stay hand-written predicates, because several run
 * once per row or column (`Stream.setRowHeight`, `Row.setHeight` in a loop) where building or walking a schema per
 * call measured several times slower; they word a refusal through the schema's own `fail`, so both read alike.
 */
import type { ProtectWorkbookOptions } from "@excel/core/workbook.browser";
import type { ProtectSheetOptions } from "@excel/core/worksheet";
import { ExcelArgumentError, ExcelInvalidAddressError } from "@excel/errors";
import { isRecordArgument } from "@utils/object";
import { check as schemaCheck, expectedOf, fail as schemaFail, s } from "@utils/schema";
import type { Schema, SchemaCheckOptions, SchemaMatches } from "@utils/schema";

const MAX_ROW = 1048576;
const MAX_COL = 16384;

function argumentError(field: string, message: string): ExcelArgumentError {
  return new ExcelArgumentError(field, message);
}

/** Where a checked value sits: the public call and the parameter path (`options`, `rule`, `options.zip`). */
export function at(where: string, field: string): SchemaCheckOptions {
  return { where, field, error: argumentError };
}

/** Check `value` against `schema` at `field` of `where`, throwing an {@link ExcelArgumentError}. */
export function checkSchema<S extends Schema<unknown, boolean>>(
  where: string,
  field: string,
  schema: S,
  value: unknown
): void {
  schemaCheck(schema, value, at(where, field));
}

/** The {@link ExcelArgumentError} for `value` at `field`, worded as a schema words it. */
export function argumentFailure(
  where: string,
  field: string,
  expected: string,
  value: unknown
): Error {
  return schemaFail(at(where, field), "", expected, value);
}

type NumberRange = { min?: number; max?: number; integer?: boolean };

function isNumberIn(value: unknown, { min, max, integer }: NumberRange): value is number {
  return (
    typeof value === "number" &&
    Number.isFinite(value) &&
    (!integer || Number.isInteger(value)) &&
    (min === undefined || value >= min) &&
    (max === undefined || value <= max)
  );
}

/** A finite number in `[min, max]` (and a whole number when `integer`). */
export function checkNumber(
  where: string,
  argument: string,
  value: unknown,
  range: NumberRange = {}
): asserts value is number {
  if (!isNumberIn(value, range)) {
    // The predicate above is the fast path; a refusal takes the schema's wording.
    throw argumentFailure(where, argument, expectedOf(s.number({ finite: true, ...range })), value);
  }
}

// The schemas whose wording the hand-written checks below refuse with.
const BOOLEAN = s.boolean();
const STRING = s.string();
const NON_EMPTY_STRING = s.string({ nonEmpty: true });
const OBJECT = s.object({});

/** What a reference argument may be, beyond a single cell. */
interface ReferenceShape {
  /** Allow `A1:B2`. */
  range?: boolean;
  /** Allow whole columns and rows (`A:C`, `2:5`). */
  lines?: boolean;
  /** Allow several references separated by spaces, as a `sqref` lists them. */
  list?: boolean;
  /** Allow (`true`) or forbid (the default) a `Sheet!` qualifier. */
  sheet?: boolean;
}

/**
 * Checks for one public function's arguments. Each check accepts `undefined` — an omitted optional field — and
 * refuses anything else that is not of the documented form; {@link ArgumentChecks.required} makes a field
 * mandatory.
 */
export interface ArgumentChecks {
  fail(argument: string, expected: string, value: unknown): never;
  required(argument: string, value: unknown): void;
  number(
    argument: string,
    value: unknown,
    range?: { min?: number; max?: number; integer?: boolean }
  ): void;
  boolean(argument: string, value: unknown): void;
  string(argument: string, value: unknown, options?: { nonEmpty?: boolean }): void;
  oneOf(argument: string, value: unknown, allowed: readonly string[]): void;
  /** Returns whether `value` is present, so the caller can go on to check its fields. */
  record(argument: string, value: unknown): value is Record<string, unknown>;
  /** Returns whether `value` is present, so the caller can go on to check its items. */
  array(
    argument: string,
    value: unknown,
    length?: { min?: number; max?: number }
  ): value is unknown[];
  /** An A1 reference: a cell, and whatever else `shape` allows. Refused with an {@link ExcelInvalidAddressError}. */
  reference(argument: string, value: unknown, shape?: ReferenceShape): void;
}

/** The checks for `where`, the name the error messages give the function (`"Sparkline.add"`). */
export function argumentChecks(where: string): ArgumentChecks {
  const checks: ArgumentChecks = {
    fail(argument, expected, value) {
      throw argumentFailure(where, argument, expected, value);
    },
    required(argument, value) {
      if (value === undefined || value === null) {
        checks.fail(argument, "given", value);
      }
    },
    number(argument, value, range) {
      if (value !== undefined) {
        checkNumber(where, argument, value, range);
      }
    },
    boolean(argument, value) {
      if (value !== undefined && typeof value !== "boolean") {
        checks.fail(argument, expectedOf(BOOLEAN), value);
      }
    },
    string(argument, value, options) {
      if (
        value !== undefined &&
        (typeof value !== "string" || (options?.nonEmpty && value === ""))
      ) {
        checks.fail(argument, expectedOf(options?.nonEmpty ? NON_EMPTY_STRING : STRING), value);
      }
    },
    oneOf(argument, value, allowed) {
      if (value !== undefined && (typeof value !== "string" || !allowed.includes(value))) {
        checks.fail(argument, expectedOf(s.oneOf(allowed)), value);
      }
    },
    record(argument, value): value is Record<string, unknown> {
      if (value === undefined) {
        return false;
      }
      if (typeof value !== "object" || value === null || Array.isArray(value)) {
        checks.fail(argument, expectedOf(OBJECT), value);
      }
      return true;
    },
    array(argument, value, length): value is unknown[] {
      if (value === undefined) {
        return false;
      }
      const { min = 0, max = Infinity } = length ?? {};
      if (!Array.isArray(value) || value.length < min || value.length > max) {
        const count =
          max === Infinity
            ? min > 0
              ? ` of at least ${min} item(s)`
              : ""
            : min === max
              ? ` of ${min} item(s)`
              : ` of ${min} to ${max} items`;
        checks.fail(argument, `an array${count}`, value);
      }
      return true;
    },
    reference(argument, value, shape = {}) {
      if (value === undefined) {
        return;
      }
      if (typeof value !== "string" || !isReference(value, shape)) {
        const kinds = [
          "a cell such as B2",
          shape.range ? "a range such as B2:D5" : "",
          shape.lines ? "whole columns or rows such as A:C or 2:5" : ""
        ].filter(Boolean);
        throw new ExcelInvalidAddressError(
          String(value),
          `${argument} must be ${kinds.join(" or ")}` +
            `${shape.list ? ", or several separated by spaces" : ""}` +
            `${shape.sheet ? ", optionally qualified by a sheet name" : ", without a sheet name"}` +
            `, inside A1:XFD${MAX_ROW}`,
          { field: argument, where }
        );
      }
    }
  };
  return checks;
}

const flag = () => s.boolean().optional();
// `spinCount` is left to its documented coercion to a whole number.
const coercedCount = () => s.unknown<number>().optional();

const sheetProtectionShape = {
  sheet: flag(),
  objects: flag(),
  scenarios: flag(),
  selectLockedCells: flag(),
  selectUnlockedCells: flag(),
  formatCells: flag(),
  formatColumns: flag(),
  formatRows: flag(),
  insertColumns: flag(),
  insertRows: flag(),
  insertHyperlinks: flag(),
  deleteColumns: flag(),
  deleteRows: flag(),
  sort: flag(),
  autoFilter: flag(),
  pivotTables: flag(),
  algorithmName: s.string().optional(),
  hashValue: s.string().optional(),
  saltValue: s.string().optional(),
  spinCount: coercedCount()
};
const sheetProtectionSchema = s.object({
  ...sheetProtectionShape,
  password: s.string().optional()
});
const _sheetProtectionMatch: SchemaMatches<typeof sheetProtectionSchema, ProtectSheetOptions> =
  true;
void _sheetProtectionMatch;

const workbookProtectionShape = {
  lockStructure: flag(),
  lockWindows: flag(),
  lockRevision: flag(),
  spinCount: coercedCount()
};
const workbookProtectionSchema = s.object({
  ...workbookProtectionShape,
  password: s.string().optional()
});
const _workbookProtectionMatch: SchemaMatches<
  typeof workbookProtectionSchema,
  ProtectWorkbookOptions
> = true;
void _workbookProtectionMatch;

/**
 * Check a `protect(ws | wb, options)` call, or its deprecated `protect(…, password, options)` form: each flag is a
 * boolean — a string such as `"yes"` was written as the opposite of what it meant — and the password and hash fields
 * are strings.
 */
export function checkProtectCall(
  where: string,
  passwordOrOptions: unknown,
  legacyOptions: unknown,
  kind: "sheet" | "workbook"
): void {
  if (isRecordArgument(passwordOrOptions)) {
    checkSchema(
      where,
      "options",
      kind === "sheet" ? sheetProtectionSchema : workbookProtectionSchema,
      passwordOrOptions
    );
    return;
  }
  // `null` is a placeholder the positional form has always passed through, for both arguments.
  checkSchema(where, "password", legacyPasswordSchema, passwordOrOptions);
  // Its options are the protection alone; the password is the argument before them.
  checkSchema(
    where,
    "options",
    kind === "sheet" ? legacySheetProtectionSchema : legacyWorkbookProtectionSchema,
    legacyOptions
  );
}

const legacyPasswordSchema = s.string().nullable().optional();
const legacySheetProtectionSchema = s.object(sheetProtectionShape).nullable().optional();
const legacyWorkbookProtectionSchema = s.object(workbookProtectionShape).nullable().optional();

/**
 * `text` — a reference {@link ArgumentChecks.reference} accepted — with its column letters in capitals: `b2:c3` is
 * `B2:C3`, and `'my sheet'!a1` is `'my sheet'!A1` (a sheet name keeps its case). A reference is case-insensitive, as
 * Excel's UI reads it; a stored or written one is canonical. An entry point that keeps the text it was given calls
 * this after the check, so what it stores is what every decoder and writer reads.
 */
export function normalizeReference(text: string): string {
  return text
    .split(" ")
    .map(part => {
      const split = withoutSheet(part);
      if (!split) {
        return part;
      }
      return part.slice(0, part.length - split.local.length) + split.local.toUpperCase();
    })
    .join(" ");
}

const CELL = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/;
const COLUMN = /^\$?([A-Za-z]{1,3})$/;
const ROW = /^\$?(\d{1,7})$/;

function columnNumber(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n;
}

function isCell(text: string): boolean {
  const m = CELL.exec(text);
  if (!m) {
    return false;
  }
  const row = Number(m[2]);
  return columnNumber(m[1]) <= MAX_COL && row >= 1 && row <= MAX_ROW;
}

function isLine(a: string, b: string): boolean {
  const ca = COLUMN.exec(a);
  const cb = COLUMN.exec(b);
  if (ca && cb) {
    return columnNumber(ca[1]) <= MAX_COL && columnNumber(cb[1]) <= MAX_COL;
  }
  const ra = ROW.exec(a);
  const rb = ROW.exec(b);
  if (ra && rb) {
    return [ra[1], rb[1]].every(r => Number(r) >= 1 && Number(r) <= MAX_ROW);
  }
  return false;
}

/** The text after a `Sheet!` / `'My Sheet'!` qualifier, or `undefined` when the qualifier is malformed. */
function withoutSheet(text: string): { local: string; qualified: boolean } | undefined {
  if (text.startsWith("'")) {
    let i = 1;
    while (i < text.length) {
      if (text[i] === "'") {
        if (text[i + 1] === "'") {
          i += 2;
          continue;
        }
        break;
      }
      i++;
    }
    return i > 1 && text[i + 1] === "!" ? { local: text.slice(i + 2), qualified: true } : undefined;
  }
  const bang = text.indexOf("!");
  if (bang < 0) {
    return { local: text, qualified: false };
  }
  return bang > 0 ? { local: text.slice(bang + 1), qualified: true } : undefined;
}

function isReference(text: string, shape: ReferenceShape): boolean {
  const parts = shape.list ? text.split(" ") : [text];
  return (
    parts.length > 0 &&
    parts.every(part => {
      const split = withoutSheet(part);
      if (!split || (split.qualified && !shape.sheet)) {
        return false;
      }
      const ends = split.local.split(":");
      if (ends.length === 1) {
        return isCell(ends[0]);
      }
      if (ends.length !== 2) {
        return false;
      }
      return (
        (!!shape.range && isCell(ends[0]) && isCell(ends[1])) ||
        (!!shape.lines && isLine(ends[0], ends[1]))
      );
    })
  );
}
