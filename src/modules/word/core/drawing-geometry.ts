/**
 * The checks that keep a drawing's numbers writable. A size, offset, wrap
 * distance or angle ends up as an integer attribute or text node, so a value
 * that is not a finite number — an object passed where a number belongs,
 * `undefined` for a required size, `NaN` from arithmetic on one — has no
 * valid form and used to be written as `cx="[object Object]"`, which Word
 * refuses to open.
 *
 * Shared by the public builders (which reject a bad value where it is
 * supplied, naming the property) and the writers (which refuse to emit one a
 * hand-built model carries). Builders also enforce the schema's sign; the
 * writers only enforce representability, because a document that was read
 * may legitimately carry values a builder would not accept.
 */

import { describeValue } from "@utils/describe-value";
import {
  MODEL_CHECK,
  checkAnchorFields,
  messagePrefix,
  sameField
} from "@word/core/argument-checks";
import type { FieldName } from "@word/core/argument-checks";
import { DocxInvalidDrawingError } from "@word/errors";

/** Whether a drawing number may be negative. */
export type DrawingNumberSign = "any" | "nonNegative";

/**
 * Throw {@link DocxInvalidDrawingError} unless `value` is a finite number
 * (and not negative, for `"nonNegative"`). `undefined` passes when `optional`.
 */
export function assertDrawingNumber(
  value: unknown,
  field: string,
  where: string,
  options: { sign?: DrawingNumberSign; optional?: boolean } = {}
): void {
  if (value === undefined && options.optional) {
    return;
  }
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new DocxInvalidDrawingError(
      field,
      `${messagePrefix(where)}${field} must be a finite number, got ${describeValue(value)}`
    );
  }
  if (options.sign === "nonNegative" && value < 0) {
    throw new DocxInvalidDrawingError(
      field,
      `${messagePrefix(where)}${field} must not be negative, got ${value}`
    );
  }
}

/** A position as the builders accept it: `{ relativeTo?, offset?, align? }`. */
interface PositionLike {
  readonly offset?: unknown;
}

/** The numeric fields a drawing frame carries, as a builder receives them. */
export interface DrawingNumbers {
  readonly width?: unknown;
  readonly height?: unknown;
  readonly horizontalPosition?: PositionLike;
  readonly verticalPosition?: PositionLike;
  readonly wrap?: {
    readonly style?: unknown;
    readonly margins?: {
      readonly top?: unknown;
      readonly bottom?: unknown;
      readonly left?: unknown;
      readonly right?: unknown;
    };
  };
  readonly simplePos?: { readonly x?: unknown; readonly y?: unknown };
  readonly distT?: unknown;
  readonly distB?: unknown;
  readonly distL?: unknown;
  readonly distR?: unknown;
  readonly rotation?: unknown;
  readonly relativeHeight?: unknown;
}

/**
 * Check every numeric field of a drawing a builder was given: the size is
 * required and not negative, offsets and angles are any finite number, wrap
 * distances are not negative — and its anchor enumerations (see
 * `checkAnchorFields`).
 */
export function assertDrawingNumbers(
  where: string,
  fields: DrawingNumbers,
  field: FieldName = sameField
): void {
  assertDrawingNumber(fields.width, field("width"), where, { sign: "nonNegative" });
  assertDrawingNumber(fields.height, field("height"), where, { sign: "nonNegative" });
  assertAnchorNumbers(where, fields, field);
}

/** {@link assertDrawingNumbers} without the size, for an anchor whose size is checked apart. */
export function assertAnchorNumbers(
  where: string,
  fields: Omit<DrawingNumbers, "width" | "height">,
  field: FieldName = sameField
): void {
  checkAnchorFields(fields, where, field);
  assertPosition(fields.horizontalPosition, field("horizontalPosition"), where);
  assertPosition(fields.verticalPosition, field("verticalPosition"), where);
  const margins = fields.wrap?.margins;
  if (margins !== undefined) {
    for (const side of ["top", "bottom", "left", "right"] as const) {
      assertDrawingNumber(margins[side], field(`wrap.margins.${side}`), where, {
        sign: "nonNegative",
        optional: true
      });
    }
  }
  if (fields.simplePos !== undefined) {
    assertDrawingNumber(fields.simplePos.x, field("simplePos.x"), where);
    assertDrawingNumber(fields.simplePos.y, field("simplePos.y"), where);
  }
  for (const name of ["distT", "distB", "distL", "distR", "relativeHeight"] as const) {
    assertDrawingNumber(fields[name], field(name), where, { sign: "nonNegative", optional: true });
  }
  assertDrawingNumber(fields.rotation, field("rotation"), where, { optional: true });
}

function assertPosition(pos: PositionLike | undefined, field: string, where: string): void {
  if (pos !== undefined) {
    assertDrawingNumber(pos.offset, `${field}.offset`, where, { optional: true });
  }
}

/**
 * The attribute text of a drawing number: rounded to the integer the schema
 * types require (`ST_Coordinate`, `ST_PositionOffset`, `ST_Angle`, …), with
 * `fallback` for an absent value. Throws {@link DocxInvalidDrawingError} for
 * anything that is not a finite number.
 */
export function drawingInteger(value: unknown, field: string, fallback?: number): string {
  const v = value === undefined ? fallback : value;
  assertDrawingNumber(v, field, MODEL_CHECK);
  return String(Math.round(v as number));
}
