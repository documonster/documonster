/**
 * Word's vocabulary for argument schemas (`@utils/schema`). Each builder gives the `must be …`
 * text the hand-written checks in `argument-checks.ts` use, so an option checked by a schema
 * reads as every other Word argument error does: `Document.toBuffer: options.compressionLevel
 * must be an integer 0–9, got 12`.
 *
 * Word's public interfaces are `readonly` throughout, so {@link w.record} types its keys
 * `readonly` and each schema is pinned with `SchemaMatches<typeof schema, Interface>`.
 */

import { check, s } from "@utils/schema";
import type { Infer, Schema, SchemaCheckOptions } from "@utils/schema";
import { DocxArgumentError } from "@word/errors";

function argumentError(field: string, message: string): DocxArgumentError {
  return new DocxArgumentError(field, message);
}

/** How a parameter of the public call `where` is reported. */
export function schemaAt(where: string, field = "options"): SchemaCheckOptions {
  return { where, field, error: argumentError };
}

/** Check `value`, the parameter `field` of `where`, throwing a {@link DocxArgumentError}. */
export function checkSchema<S extends Schema<unknown, boolean>>(
  schema: S,
  value: unknown,
  where: string,
  field = "options"
): asserts value is Infer<S> {
  check(schema, value, schemaAt(where, field));
}

interface Range {
  readonly min?: number;
  readonly max?: number;
}

function rangeText({ min, max }: Range): string {
  if (min !== undefined && max !== undefined) {
    return ` ${min}–${max}`;
  }
  if (min !== undefined) {
    return ` ≥ ${min}`;
  }
  return max === undefined ? "" : ` ≤ ${max}`;
}

const isValidDate = (v: unknown): boolean => v instanceof Date && !Number.isNaN(v.getTime());

/**
 * An optional key whose value is checked elsewhere — document model, which the model checks own,
 * or a value a dedicated check reports with its own vocabulary — named here only for its type.
 */
export function model<T>(): Schema<T, true> {
  return s.unknown<T>().optional();
}

/** The builders, worded as Word's checks are. */
export const w = {
  /** A non-null object that is not an array, typed array or `Date` — `checkRecord`. */
  record<S extends Record<string, Schema<unknown, boolean>>>(shape: S) {
    return s.object(shape, { expected: "an object", record: "argument", readonly: true });
  },
  string(): Schema<string> {
    return s.string({ expected: "a string" });
  },
  boolean(): Schema<boolean> {
    return s.boolean();
  },
  integer(range: Range = {}): Schema<number> {
    return s.number({ integer: true, ...range, expected: `an integer${rangeText(range)}` });
  },
  finite(options: { readonly min?: number; readonly positive?: boolean } = {}): Schema<number> {
    const { min, positive } = options;
    return s.custom<number>(
      v =>
        typeof v === "number" &&
        Number.isFinite(v) &&
        (min === undefined || v >= min) &&
        (!positive || v > 0),
      positive
        ? "a finite number > 0"
        : min === undefined
          ? "a finite number"
          : `a finite number ≥ ${min}`
    );
  },
  /** One of the given strings, listed in the message as JSON. */
  oneOf<const V extends readonly string[]>(values: V): Schema<V[number]> {
    return s.oneOf(values, { expected: `one of ${values.map(v => JSON.stringify(v)).join(", ")}` });
  },
  /** A size limit: a number ≥ 0, `Infinity` for none — `checkLimit`. */
  limit(): Schema<number> {
    return s.custom<number>(
      v => typeof v === "number" && v >= 0,
      "a number ≥ 0 (Infinity for no limit)"
    );
  },
  /** A `Date` that holds a time — `checkValidDate`. */
  validDate(): Schema<Date> {
    return s.custom<Date>(isValidDate, "a valid Date");
  },
  function<F = (...args: never[]) => unknown>(): Schema<F> {
    return s.function<F>({ expected: "a function" });
  },
  array<S extends Schema<unknown, boolean>>(item: S) {
    return s.readonlyArray(item, { expected: "an array" });
  }
};
