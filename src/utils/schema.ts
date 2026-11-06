/**
 * Argument schemas: one declaration gives a public function its runtime check, the `field` path
 * of the error it throws and the message, so the three cannot drift apart.
 *
 * A refusal reads `${where}: ${path} must be ${expected}, got ${describeValue(value)}` and is
 * built by the module's own argument-error constructor, which the caller passes in — this file
 * is Layer 0 and cannot name a module's error class. Paths are written the way the rest of the
 * library writes them: `options.columns[2].width` for a key or an index, `options.headers["a"]`
 * for an entry of a map.
 *
 * Unknown keys of an object are allowed unless the object asks to reject them: an options bag
 * has always ignored a key it does not know, and a closed map (`typeTransform`) says so.
 *
 * The public options type stays a hand-written interface, because that is where its TSDoc lives.
 * {@link SchemaMatches} pins the two together at compile time.
 */

import { describeValue } from "./describe-value";
import { isRecordArgument } from "./object";

/** What was refused, for a module whose error class records more than `field` and message. */
export interface SchemaRefusal {
  /** The public call. */
  where: string;
  /** The path as the message names it: `delimiter`, `options.level`. */
  label: string;
  /** The `must be …` text. */
  expected: string;
  /** The refused value, described as the message shows it. */
  got: string;
}

/**
 * The error a module throws for a refused argument, given its `field`, the message, and the
 * parts the message was built from.
 */
export type SchemaErrorFactory = (field: string, message: string, refusal: SchemaRefusal) => Error;

/** Where a value sits and how to report it. */
export interface SchemaCheckOptions {
  /** The public call, as the caller wrote it: `Csv.parse`. Starts the message. */
  where: string;
  /** The parameter the value was passed as, in the public signature: `options`, `headers`. */
  field: string;
  /** Builds the module's argument error. */
  error: SchemaErrorFactory;
  /**
   * How the parameter is named in the message. Defaults to `field`; `""` leaves it out, so a
   * refused `options.delimiter` reads `delimiter must be …` while `field` stays
   * `options.delimiter`. The parameter itself is then still named by `field`.
   */
  label?: string;
}

interface Issue {
  /** Path below the checked value: `.a`, `[0]`, `["k"]`, concatenated. */
  path: string;
  expected: string;
  got: unknown;
}

/**
 * A schema node. Build one with {@link s}; read its type with {@link Infer}.
 *
 * Every node answers two questions. `is` says whether a value passes, allocates nothing and is
 * what a successful check runs; `test` finds the first problem and its path, and runs only once
 * `is` has said there is one. That split is what lets a schema check a display list node by
 * node at the cost of the hand-written checks it replaces: a path is built only for the value
 * that fails, segment by segment as `test` unwinds, never for the values that pass.
 */
export interface Schema<T, Optional extends boolean = false> {
  /** @internal Phantom: the inferred type. */
  readonly _type?: T;
  /** @internal Phantom: whether an object key holding this schema is optional. */
  readonly _optional?: Optional;
  /** @internal Whether `value` passes. Allocates nothing. */
  readonly is: (value: unknown) => boolean;
  /** @internal The first problem with `value`, or `undefined`. */
  readonly test: (value: unknown) => Issue | undefined;
  /** @internal Whether `value` is of this schema's kind, so a union reports this member's issue. */
  readonly claims: (value: unknown) => boolean;
  /** @internal For `x.optional()`: `x`, so a container can test `undefined` itself. */
  readonly optionalOf?: Schema<unknown, boolean>;
  /** @internal For a leaf a container can test inline, what it tests. */
  readonly inline?: InlineLeaf;
  /** Also accept `undefined`, and make an object key holding it optional. */
  optional(): Schema<T, true>;
  /** Also accept `undefined`, keeping the object key required in the type. */
  orUndefined(): Schema<T, Optional>;
  /** Also accept `null`. */
  nullable(): Schema<T | null, Optional>;
  /**
   * Add a condition checked once the value is of this schema's type. It is not asked about
   * `undefined` or `null`, which only reach it through `optional`/`nullable`.
   */
  refine(predicate: (value: T) => boolean, expected: string): Schema<T, Optional>;
}

/** The TypeScript type a schema accepts. */
export type Infer<S> = S extends Schema<infer T, boolean> ? T : never;

type Simplify<T> = { [K in keyof T]: T[K] };
type Shape = Record<string, Schema<unknown, boolean>>;
type OptionalKeys<S extends Shape> = {
  [K in keyof S]: S[K] extends Schema<unknown, true> ? K : never;
}[keyof S];
type InferShape<S extends Shape> = Simplify<
  { -readonly [K in Exclude<keyof S, OptionalKeys<S>>]: Infer<S[K]> } & {
    -readonly [K in OptionalKeys<S>]?: Infer<S[K]>;
  }
>;

/**
 * An intersection flattened into one object type, modifiers kept — for pinning a schema to
 * `A & B`, which `SchemaMatches` would otherwise compare as an intersection and never match:
 * `SchemaMatches<typeof schema, Flatten<PdfTextWatermark & PdfWatermarkFilter>>`.
 */
export type Flatten<T> = { [K in keyof T]: T[K] };

/** `true` when `A` and `B` are the same type, modifiers included. */
export type IsExact<A, B> =
  (<X>() => X extends A ? 1 : 2) extends <X>() => X extends B ? 1 : 2 ? true : false;

/**
 * Resolves only when a schema's type is exactly `T`: `const _: SchemaMatches<typeof s, T> = true`.
 * A key added to one side and not the other, or a type that differs, fails to compile.
 */
export type SchemaMatches<S, T> = IsExact<Infer<S>, T>;

/**
 * A leaf an object tests inline, without calling the leaf's predicate: a call site that reaches
 * a different function for every key cannot be inlined by the engine, and on a node with a
 * dozen numeric keys that call was most of the cost.
 */
interface InlineLeaf {
  readonly code: InlineCode;
  /** Range of a number leaf. */
  readonly lo: number;
  readonly hi: number;
  /** Values of a `oneOf` leaf. */
  readonly values: readonly unknown[];
}

const InlineCode = {
  FiniteNumber: 1,
  IntegerNumber: 2,
  BoundedNumber: 3,
  AnyNumber: 4,
  String: 5,
  NonEmptyString: 6,
  Boolean: 7,
  Function: 8,
  OneOf: 9
} as const;
type InlineCode = (typeof InlineCode)[keyof typeof InlineCode];

function inline(code: InlineCode, lo = -Infinity, hi = Infinity, values: readonly unknown[] = []) {
  return { code, lo, hi, values };
}

/** {@link InlineLeaf}'s test: the same predicate the leaf's `is` runs. */
function inlineIs(leaf: InlineLeaf, v: unknown): boolean {
  switch (leaf.code) {
    case InlineCode.FiniteNumber:
      return typeof v === "number" && Number.isFinite(v) && v >= leaf.lo && v <= leaf.hi;
    case InlineCode.IntegerNumber:
      return typeof v === "number" && Number.isInteger(v) && v >= leaf.lo && v <= leaf.hi;
    case InlineCode.BoundedNumber:
      return typeof v === "number" && v >= leaf.lo && v <= leaf.hi;
    case InlineCode.AnyNumber:
      return typeof v === "number";
    case InlineCode.String:
      return typeof v === "string";
    case InlineCode.NonEmptyString:
      return typeof v === "string" && v !== "";
    case InlineCode.Boolean:
      return typeof v === "boolean";
    case InlineCode.Function:
      return typeof v === "function";
    default:
      return leaf.values.includes(v);
  }
}

const always = (): boolean => true;
const never = (): boolean => false;

/**
 * What each schema says it must be, for {@link expectedOf}: kept beside the schema rather than
 * on it, so a schema node keeps one shape whatever built it.
 */
const wordings = new WeakMap<object, string | (() => string)>();

/** The shape and options each `s.object` was built from, for {@link s.extend}. */
const objectShapes = new WeakMap<object, { shape: Shape; options: ObjectOptions }>();

function worded<S extends object>(schema: S, text: string | (() => string) | undefined): S {
  if (text !== undefined) {
    wordings.set(schema, text);
  }
  return schema;
}

/**
 * The `must be …` text a schema refuses a value with at its root — what a hand-written fast
 * path passes to `fail()` when its own predicate refuses, so its message is the schema's
 * without building or running the schema: `fail(at, "", expectedOf(LEVEL), value)`.
 */
export function expectedOf(schema: Schema<unknown, boolean>): string {
  const text = wordings.get(schema);
  if (text === undefined) {
    return "valid";
  }
  return typeof text === "string" ? text : text();
}

function make<T>(
  is: (value: unknown) => boolean,
  test: (value: unknown) => Issue | undefined,
  claims: (value: unknown) => boolean,
  optionalOf?: Schema<unknown, boolean>,
  inlineLeaf?: InlineLeaf
): Schema<T> {
  const schema: Schema<T> = {
    is,
    test,
    claims,
    optionalOf,
    inline: inlineLeaf,
    optional: () =>
      worded(
        make<T>(
          v => v === undefined || is(v),
          v => (v === undefined ? undefined : test(v)),
          v => v === undefined || claims(v),
          schema
        ),
        () => expectedOf(schema)
      ) as unknown as Schema<T, true>,
    orUndefined: () =>
      worded(
        make<T>(
          v => v === undefined || is(v),
          v => (v === undefined ? undefined : test(v)),
          v => v === undefined || claims(v)
        ),
        () => expectedOf(schema)
      ),
    nullable: () =>
      worded(
        make<T | null>(
          v => v === null || is(v),
          v => (v === null ? undefined : test(v)),
          v => v === null || claims(v)
        ),
        () => expectedOf(schema)
      ),
    refine: (predicate, expected) =>
      worded(
        make<T>(
          v => is(v) && (v === undefined || v === null || predicate(v as T)),
          v =>
            test(v) ??
            (v === undefined || v === null || predicate(v as T) ? undefined : issue(expected, v)),
          claims
        ),
        () => expectedOf(schema)
      )
  };
  return schema;
}

/** A leaf: one predicate, one message. `test` is derived from `is`. */
function leaf<T>(
  is: (value: unknown) => boolean,
  expected: string,
  claims: (value: unknown) => boolean = is,
  inlineLeaf?: InlineLeaf
): Schema<T> {
  return worded(
    make<T>(is, v => (is(v) ? undefined : issue(expected, v)), claims, undefined, inlineLeaf),
    expected
  );
}

function issue(expected: string, got: unknown, path = ""): Issue {
  return { path, expected, got };
}

/** Prefix a failing child's issue with the segment that leads to it; runs on failure only. */
function under(segment: string, found: Issue | undefined): Issue | undefined {
  if (found) {
    found.path = segment + found.path;
  }
  return found;
}

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function tagOf(value: unknown): string {
  return Object.prototype.toString.call(value);
}

const IDENTIFIER = /^[A-Za-z_$][\w$]*$/;

/** A key as a path segment: `.name` when it is an identifier, `["a b"]` otherwise. */
function keySegment(key: string): string {
  return IDENTIFIER.test(key) ? `.${key}` : `[${JSON.stringify(key)}]`;
}

interface CommonOptions {
  /** The `must be …` text, replacing the default. */
  expected?: string;
}

interface NumberOptions extends CommonOptions {
  integer?: boolean;
  /** Refuse `NaN` and `±Infinity`. Implied by `integer`. */
  finite?: boolean;
  min?: number;
  max?: number;
}

function numberExpected(o: NumberOptions): string {
  const noun = o.integer ? "an integer" : o.finite ? "a finite number" : "a number";
  if (o.min !== undefined && o.max !== undefined) {
    return `${noun} from ${o.min} to ${o.max}`;
  }
  if (o.min !== undefined) {
    return `${noun} >= ${o.min}`;
  }
  return o.max === undefined ? noun : `${noun} <= ${o.max}`;
}

/**
 * The predicate of `s.number`, specialised for the common option sets so the hot call is a
 * single comparison chain rather than a run of flag tests.
 */
function numberPredicate(
  integer: boolean,
  finite: boolean,
  min: number | undefined,
  max: number | undefined
): (v: unknown) => boolean {
  const lo = min ?? -Infinity;
  const hi = max ?? Infinity;
  const bounded = min !== undefined || max !== undefined;
  if (integer) {
    return v => typeof v === "number" && Number.isInteger(v) && v >= lo && v <= hi;
  }
  if (finite) {
    return v => typeof v === "number" && Number.isFinite(v) && v >= lo && v <= hi;
  }
  if (bounded) {
    // `NaN` fails both comparisons; `±Infinity` passes when in range.
    return v => typeof v === "number" && v >= lo && v <= hi;
  }
  return v => typeof v === "number";
}

interface ObjectOptions<R extends boolean = boolean> extends CommonOptions {
  /** Refuse an own key the shape does not name, reported as `["key"]` with the key as the value. */
  unknownKeys?: "allow" | "reject";
  /** Type every key `readonly`, for an interface that declares them so. No runtime effect. */
  readonly?: R;
  /**
   * Which values count as an object. `"plain"` (the default) is any non-null object that is not
   * an array. `"argument"` is what `isRecordArgument` accepts: it also refuses a `Date`, a typed
   * array, an `ArrayBuffer`, a `Map`, a `Set` and any other iterable — the test an overload uses
   * to tell an options object from a positional value.
   */
  record?: "plain" | "argument";
}

interface VariantOptions extends CommonOptions {
  /**
   * How a value that is an object but carries no known tag is reported. By default it is
   * reported like any other refused value, against `expected`. Given, it is reported at the tag
   * itself (`node.kind`), with this text and the tag as the value — `expected` is then only for a
   * value that is not an object at all.
   */
  unknownTag?: string;
}

type UnionOf<M extends readonly Schema<unknown, boolean>[]> = Infer<M[number]>;

/** The schema builders. */
export const s = {
  /** Anything; for a key the schema names only for its type. */
  unknown<T = unknown>(): Schema<T> {
    return worded(
      make<T>(always, () => undefined, always),
      "anything"
    );
  },

  string(options: CommonOptions & { nonEmpty?: boolean } = {}): Schema<string> {
    const expected = options.expected ?? (options.nonEmpty ? "a non-empty string" : "a string");
    const is = options.nonEmpty
      ? (v: unknown): boolean => typeof v === "string" && v !== ""
      : (v: unknown): boolean => typeof v === "string";
    return leaf<string>(
      is,
      expected,
      v => typeof v === "string",
      inline(options.nonEmpty ? InlineCode.NonEmptyString : InlineCode.String)
    );
  },

  number(options: NumberOptions = {}): Schema<number> {
    const expected = options.expected ?? numberExpected(options);
    const { integer, finite, min, max } = options;
    const is = numberPredicate(integer === true, finite === true, min, max);
    const code = integer
      ? InlineCode.IntegerNumber
      : finite
        ? InlineCode.FiniteNumber
        : min !== undefined || max !== undefined
          ? InlineCode.BoundedNumber
          : InlineCode.AnyNumber;
    return leaf<number>(
      is,
      expected,
      v => typeof v === "number",
      inline(code, min ?? -Infinity, max ?? Infinity)
    );
  },

  boolean(options: CommonOptions = {}): Schema<boolean> {
    const is = (v: unknown): boolean => typeof v === "boolean";
    return leaf<boolean>(is, options.expected ?? "a boolean", is, inline(InlineCode.Boolean));
  },

  /** A function; name its signature as the type argument. */
  function<F = (...args: never[]) => unknown>(options: CommonOptions = {}): Schema<F> {
    const is = (v: unknown): boolean => typeof v === "function";
    return leaf<F>(is, options.expected ?? "a function", is, inline(InlineCode.Function));
  },

  /** One of the given values, compared with SameValueZero. */
  literal<const V extends readonly unknown[]>(...values: V): Schema<V[number]> {
    return s.oneOf(values);
  },

  /** {@link s.literal} taking an array, and an optional message. */
  oneOf<const V extends readonly unknown[]>(
    values: V,
    options: CommonOptions = {}
  ): Schema<V[number]> {
    const expected = options.expected ?? `one of ${values.map(v => describeValue(v)).join(", ")}`;
    return leaf<V[number]>(
      v => values.includes(v),
      expected,
      never,
      inline(InlineCode.OneOf, -Infinity, Infinity, values)
    );
  },

  /**
   * A string that is a member of `values` — a vocabulary the module already holds as a `Set` or
   * an array — typed `T`. Listed in the message as JSON unless `expected` is given. In a union
   * it claims every string, so a string outside the vocabulary is reported against it.
   */
  member<T extends string = string>(
    values: ReadonlySet<string> | readonly string[],
    options: CommonOptions = {}
  ): Schema<T> {
    const set: ReadonlySet<string> = Array.isArray(values)
      ? new Set(values)
      : (values as ReadonlySet<string>);
    const expected =
      options.expected ?? `one of ${[...set].map(v => JSON.stringify(v)).join(", ")}`;
    return leaf<T>(
      v => typeof v === "string" && set.has(v),
      expected,
      v => typeof v === "string"
    );
  },

  array<S extends Schema<unknown, boolean>>(
    item: S,
    options: CommonOptions & { minLength?: number } = {}
  ): Schema<Infer<S>[]> {
    const expected = options.expected ?? "an array";
    const minLength = options.minLength ?? 0;
    const itemIs = item.is;
    const is = (v: unknown): boolean => {
      if (!Array.isArray(v) || v.length < minLength) {
        return false;
      }
      for (let i = 0; i < v.length; i++) {
        if (!itemIs(v[i])) {
          return false;
        }
      }
      return true;
    };
    return worded(
      make<Infer<S>[]>(
        is,
        v => {
          if (!Array.isArray(v) || v.length < minLength) {
            return issue(expected, v);
          }
          for (let i = 0; i < v.length; i++) {
            if (!itemIs(v[i])) {
              return under(`[${i}]`, item.test(v[i]));
            }
          }
          return undefined;
        },
        Array.isArray
      ),
      expected
    );
  },

  /** {@link s.array} typed `readonly T[]`. */
  readonlyArray<S extends Schema<unknown, boolean>>(
    item: S,
    options: CommonOptions & { minLength?: number } = {}
  ): Schema<readonly Infer<S>[]> {
    return s.array(item, options);
  },

  /** A non-null, non-array object holding the keys of `shape`. */
  object<S extends Shape, R extends boolean = false>(
    shape: S,
    options: ObjectOptions<R> = {}
  ): Schema<R extends true ? Readonly<InferShape<S>> : InferShape<S>> {
    const expected = options.expected ?? "an object";
    const keys = Object.keys(shape);
    const count = keys.length;
    // An optional key is tested as `undefined || base`, inline, rather than through the
    // wrapper `optional()` made: one call per key instead of two.
    const optional = keys.map(key => shape[key].optionalOf !== undefined);
    const checks = keys.map(key => (shape[key].optionalOf ?? shape[key]).is);
    const leaves = keys.map(key => (shape[key].optionalOf ?? shape[key]).inline);
    const segments = keys.map(keySegment);
    const rejectUnknown = options.unknownKeys === "reject";
    const isObject = options.record === "argument" ? isRecordArgument : isPlainRecord;
    const known = (v: Record<string, unknown>): string | undefined => {
      for (const key of Object.keys(v)) {
        if (!Object.hasOwn(shape, key)) {
          return key;
        }
      }
      return undefined;
    };
    const is = (v: unknown): boolean => {
      if (!isObject(v)) {
        return false;
      }
      const record = v as Record<string, unknown>;
      if (rejectUnknown && known(record) !== undefined) {
        return false;
      }
      for (let i = 0; i < count; i++) {
        const item = record[keys[i]];
        if (item === undefined && optional[i]) {
          continue;
        }
        const leafI = leaves[i];
        if (leafI === undefined ? !checks[i](item) : !inlineIs(leafI, item)) {
          return false;
        }
      }
      return true;
    };
    const built = worded(
      make<R extends true ? Readonly<InferShape<S>> : InferShape<S>>(
        is,
        v => {
          if (!isObject(v)) {
            return issue(expected, v);
          }
          const record = v as Record<string, unknown>;
          if (rejectUnknown) {
            const key = known(record);
            if (key !== undefined) {
              return issue(`a key in ${keys.join(", ")}`, key, `[${JSON.stringify(key)}]`);
            }
          }
          for (let i = 0; i < count; i++) {
            if (!shape[keys[i]].is(record[keys[i]])) {
              return under(segments[i], shape[keys[i]].test(record[keys[i]]));
            }
          }
          return undefined;
        },
        isObject
      ),
      expected
    );
    objectShapes.set(built, { shape, options });
    return built;
  },

  /**
   * An object schema extended with more keys, typed as the interface that `extends` it: the
   * base's keys keep their own modifiers and the new ones take `options` (`readonly: true` for
   * an interface that declares its own keys `readonly` over a base whose keys are not). The
   * base must be built with `s.object` or `s.extend`; its `expected`, `unknownKeys` and `record`
   * carry over unless `options` gives them.
   */
  extend<B extends Schema<object, boolean>, S extends Shape, R extends boolean = false>(
    base: B,
    shape: S,
    options: ObjectOptions<R> = {}
  ): Schema<Flatten<Infer<B> & (R extends true ? Readonly<InferShape<S>> : InferShape<S>)>> {
    const from = objectShapes.get(base);
    if (from === undefined) {
      throw new TypeError("s.extend: base must be built with s.object or s.extend");
    }
    return s.object(
      { ...from.shape, ...shape },
      { ...from.options, ...options }
    ) as unknown as Schema<
      Flatten<Infer<B> & (R extends true ? Readonly<InferShape<S>> : InferShape<S>)>
    >;
  },

  /**
   * A non-null, non-array object whose own entries all match `value`, reported as `["key"]`.
   * `record: "argument"` takes objects as {@link ObjectOptions.record} does.
   */
  record<S extends Schema<unknown, boolean>>(
    value: S,
    options: CommonOptions & { record?: "plain" | "argument" } = {}
  ): Schema<Record<string, Infer<S>>> {
    const expected = options.expected ?? "an object";
    const valueIs = value.is;
    const isObject = options.record === "argument" ? isRecordArgument : isPlainRecord;
    return worded(
      make<Record<string, Infer<S>>>(
        v => {
          if (!isObject(v)) {
            return false;
          }
          const record = v as Record<string, unknown>;
          for (const key of Object.keys(record)) {
            if (!valueIs(record[key])) {
              return false;
            }
          }
          return true;
        },
        v => {
          if (!isObject(v)) {
            return issue(expected, v);
          }
          const record = v as Record<string, unknown>;
          for (const key of Object.keys(record)) {
            if (!valueIs(record[key])) {
              return under(`[${JSON.stringify(key)}]`, value.test(record[key]));
            }
          }
          return undefined;
        },
        isObject
      ),
      expected
    );
  },

  /**
   * Any of `members`: a value any member accepts passes. A refused value is reported against
   * the first member whose kind it is (a string against the string member, an object against
   * the object member), so a wrong entry inside a
   * map is named where it is; a value no member claims is reported against the union, with
   * `expected` (default: the members' texts joined), unless a member failed on something inside it.
   */
  union<const M extends readonly Schema<unknown, boolean>[]>(
    members: M,
    options: CommonOptions = {}
  ): Schema<UnionOf<M>> {
    const count = members.length;
    const claimant = (v: unknown): Schema<unknown, boolean> | undefined => {
      for (let i = 0; i < count; i++) {
        if (members[i].claims(v)) {
          return members[i];
        }
      }
      return undefined;
    };
    return worded(
      make<UnionOf<M>>(
        // A value any member accepts passes. Claiming only chooses whose refusal is reported, so
        // a narrow member (`s.literal(Infinity)`) is not shadowed by a wider one of the same kind
        // (`s.number({ integer: true })`) that claims the value first.
        v => {
          for (let i = 0; i < count; i++) {
            if (members[i].is(v)) {
              return true;
            }
          }
          return false;
        },
        v => {
          for (let i = 0; i < count; i++) {
            if (members[i].is(v)) {
              return undefined;
            }
          }
          const member = claimant(v);
          if (member) {
            return member.test(v);
          }
          let inner: Issue | undefined;
          for (const m of members) {
            const found = m.test(v);
            if (!found) {
              return undefined;
            }
            if (found.path !== "" && !inner) {
              inner = found;
            }
          }
          return (
            inner ??
            issue(options.expected ?? members.map(m => m.test(v)!.expected).join(" or "), v)
          );
        },
        v => claimant(v) !== undefined
      ),
      options.expected ?? (() => members.map(m => expectedOf(m)).join(" or "))
    );
  },

  /**
   * A tagged union: an object whose `key` names which of `variants` it is, then checked
   * against that variant's schema. A value with no such tag is reported against `expected`, or
   * — with `unknownTag` — at the tag.
   */
  variant<K extends string, V extends Record<string, Schema<unknown, boolean>>>(
    key: K,
    variants: V,
    options: VariantOptions = {}
  ): Schema<Infer<V[keyof V]>> {
    const tags = Object.keys(variants);
    const expected =
      options.expected ??
      (options.unknownTag === undefined
        ? `an object whose ${key} is one of ${tags.map(t => JSON.stringify(t)).join(", ")}`
        : "an object");
    const segment = keySegment(key);
    const pick = (v: unknown): Schema<unknown, boolean> | undefined => {
      const tag = (v as Record<string, unknown>)[key];
      return typeof tag === "string" && Object.hasOwn(variants, tag) ? variants[tag] : undefined;
    };
    return worded(
      make<Infer<V[keyof V]>>(
        v => {
          if (!isPlainRecord(v)) {
            return false;
          }
          const chosen = pick(v);
          return chosen !== undefined && chosen.is(v);
        },
        v => {
          if (!isPlainRecord(v)) {
            return issue(expected, v);
          }
          const chosen = pick(v);
          if (chosen === undefined) {
            return options.unknownTag === undefined
              ? issue(expected, v)
              : issue(options.unknownTag, v[key], segment);
          }
          return chosen.test(v);
        },
        isPlainRecord
      ),
      expected
    );
  },

  /** Every schema in turn, typed as their intersection: a custom gate before a shape, say. */
  all<A, B>(first: Schema<A, boolean>, second: Schema<B, boolean>): Schema<A & B> {
    return worded(
      make<A & B>(
        v => first.is(v) && second.is(v),
        v => first.test(v) ?? second.test(v),
        v => first.claims(v) && second.claims(v)
      ),
      () => expectedOf(first)
    );
  },

  /**
   * A value `predicate` accepts. `got` chooses what the message shows, when that is not the
   * value itself (the `type` of a column definition, say).
   */
  custom<T>(
    predicate: (value: unknown) => boolean,
    expected: string,
    options: { got?: (value: unknown) => unknown } = {}
  ): Schema<T> {
    const got = options.got;
    return worded(
      make<T>(
        predicate,
        v => (predicate(v) ? undefined : issue(expected, got ? got(v) : v)),
        never
      ),
      expected
    );
  },

  /**
   * An `AbortSignal`, recognised by shape (`aborted` and `addEventListener`), so a signal from
   * another realm or a polyfill is accepted.
   */
  abortSignal(options: CommonOptions = {}): Schema<AbortSignal> {
    const is = (v: unknown): boolean =>
      typeof v === "object" &&
      v !== null &&
      typeof (v as { aborted?: unknown }).aborted === "boolean" &&
      typeof (v as { addEventListener?: unknown }).addEventListener === "function";
    return leaf<AbortSignal>(is, options.expected ?? "an AbortSignal");
  },

  /** A promise or any thenable. */
  thenable<T = unknown>(options: CommonOptions = {}): Schema<PromiseLike<T>> {
    const is = (v: unknown): boolean =>
      (typeof v === "object" || typeof v === "function") &&
      v !== null &&
      typeof (v as { then?: unknown }).then === "function";
    return leaf<PromiseLike<T>>(is, options.expected ?? "a promise or thenable");
  },

  /**
   * An iterable (`Symbol.iterator`) or an async iterable (`Symbol.asyncIterator`). A string is
   * iterable and is accepted only when `strings` is set.
   */
  iterable<T = unknown>(
    options: CommonOptions & { strings?: boolean } = {}
  ): Schema<Iterable<T> | AsyncIterable<T>> {
    const is = (v: unknown): boolean => {
      if (typeof v === "string") {
        return options.strings === true;
      }
      if ((typeof v !== "object" && typeof v !== "function") || v === null) {
        return false;
      }
      return (
        typeof (v as { [Symbol.iterator]?: unknown })[Symbol.iterator] === "function" ||
        typeof (v as { [Symbol.asyncIterator]?: unknown })[Symbol.asyncIterator] === "function"
      );
    };
    return leaf<Iterable<T> | AsyncIterable<T>>(
      is,
      options.expected ?? "an iterable or async iterable"
    );
  },

  /**
   * A built-in object of the given kind — `"URL"`, `"Blob"`, `"ReadableStream"` — recognised by its
   * `Object.prototype.toString` tag, so one from another realm is accepted.
   */
  builtin<T>(tag: string, options: CommonOptions = {}): Schema<T> {
    const expected = options.expected ?? (/^[AEIO]/.test(tag) ? `an ${tag}` : `a ${tag}`);
    const full = `[object ${tag}]`;
    const is = (v: unknown): boolean => tagOf(v) === full;
    return leaf<T>(is, expected, is);
  },

  /**
   * Binary data: an `ArrayBuffer` (or `SharedArrayBuffer`) or any `ArrayBufferView`. Recognised by tag
   * rather than `instanceof`, so bytes from another realm are accepted.
   */
  bytes<T = Uint8Array | ArrayBuffer>(options: CommonOptions = {}): Schema<T> {
    const is = (v: unknown): boolean => {
      if (ArrayBuffer.isView(v)) {
        return true;
      }
      const tag = tagOf(v);
      return tag === "[object ArrayBuffer]" || tag === "[object SharedArrayBuffer]";
    };
    return leaf<T>(is, options.expected ?? "an ArrayBuffer or ArrayBufferView");
  }
};

/** The field and the message label for a value at `path` below the root `field`. */
function render(
  field: string,
  label: string | undefined,
  path: string
): { field: string; label: string } {
  const full = field + path;
  const root = label ?? field;
  if (root !== "") {
    return { field: full, label: root + path };
  }
  return { field: full, label: path === "" ? field : path.replace(/^\./, "") };
}

function build(
  where: string,
  field: string,
  label: string | undefined,
  error: SchemaErrorFactory,
  path: string,
  expected: string,
  got: unknown
): Error {
  const named = render(field, label, path);
  const shown = describeValue(got);
  return error(named.field, `${where}: ${named.label} must be ${expected}, got ${shown}`, {
    where,
    label: named.label,
    expected,
    got: shown
  });
}

/**
 * Check `value` against `schema`, throwing the module's argument error for the first problem.
 * A value that passes costs one call of the schema's predicate and allocates nothing.
 */
export function check<S extends Schema<unknown, boolean>>(
  schema: S,
  value: unknown,
  options: SchemaCheckOptions
): asserts value is Infer<S> {
  if (schema.is(value)) {
    return;
  }
  const found = schema.test(value) ?? issue("valid", value);
  throw build(
    options.where,
    options.field,
    options.label,
    options.error,
    found.path,
    found.expected,
    found.got
  );
}

/**
 * {@link check} with its context spelled out, for a check that runs per node or per call: the
 * caller passes what it already holds rather than an options object built for every call.
 */
export function checkAt<S extends Schema<unknown, boolean>>(
  schema: S,
  value: unknown,
  where: string,
  field: string,
  error: SchemaErrorFactory,
  label?: string
): asserts value is Infer<S> {
  if (schema.is(value)) {
    return;
  }
  const found = schema.test(value) ?? issue("valid", value);
  throw build(where, field, label, error, found.path, found.expected, found.got);
}

/**
 * The error {@link check} throws, for a rule a schema cannot express (one option against
 * another): `path` is below `options.field`, written `.key`, `[i]` or `["key"]`.
 */
export function fail(
  options: SchemaCheckOptions,
  path: string,
  expected: string,
  got: unknown
): Error {
  return build(options.where, options.field, options.label, options.error, path, expected, got);
}
