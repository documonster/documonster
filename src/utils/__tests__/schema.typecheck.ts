// This file is typechecked by `pnpm type` but is NOT executed by Vitest.
//
// It pins `SchemaMatches`, the compile-time link between an argument schema and the hand-written
// options interface that carries the TSDoc: a key on one side only, or a type that differs, has to
// fail to compile. Each `@ts-expect-error` below becomes unused — and therefore an error — if the
// check ever stops seeing the drift.

import { s } from "@utils/schema";
import type { Flatten, SchemaMatches } from "@utils/schema";

interface Options {
  name?: string;
  count?: number;
  mode: "a" | "b";
  columns?: readonly string[];
  map?: Record<string, boolean>;
  onRow?: (row: string[]) => void;
}

const exact = s.object({
  name: s.string().optional(),
  count: s.number().optional(),
  mode: s.literal("a", "b"),
  columns: s.readonlyArray(s.string()).optional(),
  map: s.record(s.boolean()).optional(),
  onRow: s.function<(row: string[]) => void>().optional()
});
const matches: SchemaMatches<typeof exact, Options> = true;

const missingKey = s.object({
  name: s.string().optional(),
  count: s.number().optional(),
  mode: s.literal("a", "b"),
  columns: s.readonlyArray(s.string()).optional(),
  map: s.record(s.boolean()).optional()
});
// @ts-expect-error -- `onRow` is in the interface and not in the schema.
const missing: SchemaMatches<typeof missingKey, Options> = true;

const extraKey = s.object({
  name: s.string().optional(),
  count: s.number().optional(),
  mode: s.literal("a", "b"),
  columns: s.readonlyArray(s.string()).optional(),
  map: s.record(s.boolean()).optional(),
  onRow: s.function<(row: string[]) => void>().optional(),
  extra: s.boolean().optional()
});
// @ts-expect-error -- `extra` is in the schema and not in the interface.
const extra: SchemaMatches<typeof extraKey, Options> = true;

const wrongType = s.object({
  name: s.string().optional(),
  count: s.string().optional(),
  mode: s.literal("a", "b"),
  columns: s.readonlyArray(s.string()).optional(),
  map: s.record(s.boolean()).optional(),
  onRow: s.function<(row: string[]) => void>().optional()
});
// @ts-expect-error -- `count` is a number in the interface.
const wrong: SchemaMatches<typeof wrongType, Options> = true;

const optionalRequired = s.object({
  name: s.string().optional(),
  count: s.number().optional(),
  mode: s.literal("a", "b").optional(),
  columns: s.readonlyArray(s.string()).optional(),
  map: s.record(s.boolean()).optional(),
  onRow: s.function<(row: string[]) => void>().optional()
});
// @ts-expect-error -- `mode` is required in the interface.
const optional: SchemaMatches<typeof optionalRequired, Options> = true;

const mutableArray = s.object({
  name: s.string().optional(),
  count: s.number().optional(),
  mode: s.literal("a", "b"),
  columns: s.array(s.string()).optional(),
  map: s.record(s.boolean()).optional(),
  onRow: s.function<(row: string[]) => void>().optional()
});
// @ts-expect-error -- `columns` is a readonly array in the interface.
const mutable: SchemaMatches<typeof mutableArray, Options> = true;

void [matches, missing, extra, wrong, optional, mutable];

interface ReadonlyOptions {
  readonly dpi?: number;
}
const readonlySchema = s.object({ dpi: s.number().optional() }, { readonly: true });
const readonlyMatches: SchemaMatches<typeof readonlySchema, ReadonlyOptions> = true;
// @ts-expect-error -- `dpi` is readonly in the interface.
const notReadonly: SchemaMatches<typeof mutableDpi, ReadonlyOptions> = true;
const mutableDpi = s.object({ dpi: s.number().optional() });
void [readonlyMatches, notReadonly];

interface Base {
  a?: number;
}
interface Extra {
  b: string;
}
const intersected = s.object({ a: s.number().optional(), b: s.string() });
const flattened: SchemaMatches<typeof intersected, Flatten<Base & Extra>> = true;
// @ts-expect-error -- an intersection is not identical to the object it describes.
const unflattened: SchemaMatches<typeof intersected, Base & Extra> = true;
void [flattened, unflattened];

interface BaseOptions {
  a?: number;
}
interface ExtendedOptions extends BaseOptions {
  readonly b?: string;
}
const baseOptions = s.object({ a: s.number().optional() });
const extendedOptions = s.extend(baseOptions, { b: s.string().optional() }, { readonly: true });
const extendedMatches: SchemaMatches<typeof extendedOptions, ExtendedOptions> = true;
const mutableExtension = s.extend(baseOptions, { b: s.string().optional() });
// @ts-expect-error -- `b` is readonly in the interface.
const extendedMismatch: SchemaMatches<typeof mutableExtension, ExtendedOptions> = true;
void [extendedMatches, extendedMismatch];
