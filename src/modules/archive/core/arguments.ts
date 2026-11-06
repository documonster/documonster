/**
 * The argument-schema plumbing of the `documonster/archive` entry: refusals are
 * {@link ArchiveArgumentError}s whose message names the field as `field` does
 * (`Archive.packZip: options.level must be …`), and the schemas several calls share.
 */

import { ArchiveArgumentError } from "@archive/core/errors";
import { check, s } from "@utils/schema";
import type { Infer, Schema, SchemaCheckOptions } from "@utils/schema";

function argumentError(field: string, message: string): ArchiveArgumentError {
  return new ArchiveArgumentError(field, message);
}

/** Where a parameter of an `Archive` call sits, for `check`. */
export function archiveArgument(where: string, field: string): SchemaCheckOptions {
  return { where, field, error: argumentError };
}

/** Check one argument of a public call; `field` is its path in the public signature. */
export function checkArchiveArgument<S extends Schema<unknown, boolean>>(
  schema: S,
  value: unknown,
  where: string,
  field: string
): asserts value is Infer<S> {
  check(schema, value, archiveArgument(where, field));
}

/** An options argument as the archive calls read it: any non-null object, arrays refused. */
export const OPTIONS = { expected: "an options object" };

/** A compression level: an integer from 0 (store) to 9. */
export const levelSchema = s.number({ integer: true, min: 0, max: 9 });

/** A valid `Date`. */
export const dateSchema = s.custom<Date>(
  v => v instanceof Date && !Number.isNaN(v.getTime()),
  "a valid Date"
);

/** A size or count limit: an integer from 0, `Infinity` meaning none. */
export const limitSchema = s.custom<number>(
  v =>
    typeof v === "number" &&
    !Number.isNaN(v) &&
    v >= 0 &&
    (!Number.isFinite(v) || Number.isInteger(v)),
  "an integer from 0, or Infinity"
);

/** An integer from 0 to `max`. */
export function uintSchema(max: number): Schema<number> {
  return s.number({ integer: true, min: 0, max });
}

/** The keys every archive stream operation takes (`ArchiveStreamOptions`). */
export function streamOptionShape<P>() {
  return {
    signal: s.abortSignal().optional(),
    onProgress: s.function<(p: P) => void>().optional(),
    progressIntervalMs: s.number({ min: 0 }).optional()
  };
}
