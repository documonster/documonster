/**
 * Argument schemas of the `documonster/stream` entry. A refusal is a {@link StreamArgumentError},
 * whose message keeps this module's own form, `${where}: expected ${expected}, got ${actual}`,
 * with `field` naming the argument.
 */

import { StreamArgumentError } from "@stream/errors";
import { isEncoding } from "@stream/kernel/state";
import type { BufferedStreamOptions, PullStreamOptions } from "@stream/types";
import { check, s } from "@utils/schema";
import type { Schema, SchemaCheckOptions, SchemaMatches } from "@utils/schema";

/** Where a parameter of a `Streams` call sits, for `check`. */
export function streamArgument(where: string, field: string): SchemaCheckOptions {
  return {
    where,
    field,
    error: (f, _message, refusal) =>
      new StreamArgumentError(refusal.expected, refusal.got, { field: f, where })
  };
}

/** Check one argument of a public call. */
export function checkStreamArgument(
  schema: Schema<unknown, boolean>,
  value: unknown,
  where: string,
  field: string
): void {
  check(schema, value, streamArgument(where, field));
}

/** A byte array, recognised structurally so one from another realm is accepted too. */
export const byteArraySchema = s.custom<Uint8Array>(
  v => ArrayBuffer.isView(v) && (v as { BYTES_PER_ELEMENT?: number }).BYTES_PER_ELEMENT === 1,
  "a Uint8Array"
);

/**
 * An options object as the stream classes have always read it: anything but `null` that is an
 * object, arrays refused.
 */
const OPTIONS = { expected: "an options object" };

export const bufferedStreamOptionsSchema = s.object(
  {
    batchSize: s
      .number({ integer: true, min: 1, expected: "a positive integer batchSize" })
      .optional()
  },
  OPTIONS
);
const _bufferedStreamOptionsMatch: SchemaMatches<
  typeof bufferedStreamOptionsSchema,
  BufferedStreamOptions
> = true;
void _bufferedStreamOptionsMatch;

export const pullStreamOptionsSchema = s.object(
  { highWaterMark: s.number({ min: 0, expected: "a non-negative highWaterMark" }).optional() },
  OPTIONS
);
const _pullStreamOptionsMatch: SchemaMatches<typeof pullStreamOptionsSchema, PullStreamOptions> =
  true;
void _pullStreamOptionsMatch;

const highWaterMark = () =>
  s.number({ integer: true, min: 0, expected: "a non-negative integer highWaterMark" }).optional();

export const webReadableOptionsSchema = s.object(
  {
    objectMode: s.boolean().optional(),
    highWaterMark: highWaterMark(),
    encoding: s.custom<string>(isEncoding, "a Buffer encoding name").optional(),
    signal: s.abortSignal().optional()
  },
  OPTIONS
);

export const webWritableOptionsSchema = s.object(
  {
    objectMode: s.boolean().optional(),
    highWaterMark: highWaterMark(),
    decodeStrings: s.boolean().optional(),
    signal: s.abortSignal().optional()
  },
  OPTIONS
);

export const toWebReadableOptionsSchema = s.object(
  {
    strategy: s
      .custom<QueuingStrategy>(v => typeof v === "object" && v !== null, "a QueuingStrategy")
      .optional()
  },
  OPTIONS
);
