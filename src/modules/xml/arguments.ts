/**
 * Argument checks for the `Xml` entry points. A value with no reading is an {@link XmlError}
 * naming the call and the option, raised before anything is parsed or written: `maxDepth:
 * NaN` used to switch the nesting limit off without a word, and a `null` input surfaced as a
 * `TypeError` from deep inside the parser.
 *
 * The options are declared as schemas (`@utils/schema`), each pinned to its interface.
 */

import { isUint8Array } from "@utils/binary";
import { check, fail, s } from "@utils/schema";
import type { SchemaCheckOptions, SchemaMatches } from "@utils/schema";
import { isXmlName, isXmlText } from "@xml/chars";
import { contextBytes } from "@xml/context";
import { XmlArgumentError } from "@xml/errors";
import type {
  XmlParseOptions,
  XmlParseToObjectOptions,
  XmlPartitionStructure,
  XmlAttributeDeclaration,
  XmlPrologContext,
  XmlReadRecordsOptions,
  XmlReadRootOptions,
  XmlSaxOptions,
  XmlToPlainObjectOptions
} from "@xml/types";

function argumentError(field: string, message: string): XmlArgumentError {
  return new XmlArgumentError(field, message);
}

/** A parameter of an `Xml` call, named in the message by its path (`options.maxDepth`). */
export function xmlArgument(where: string, field: string): SchemaCheckOptions {
  return { where, field, error: argumentError };
}

const flag = () => s.boolean().optional();
const limit = () =>
  s
    .number({ integer: true, min: 0, expected: "an integer >= 0 (0 switches the limit off)" })
    .optional();
const invalidCharHandling = () =>
  s.oneOf(["error", "skip", "replace"], { expected: '"error", "skip" or "replace"' }).optional();
const OPTIONS = { expected: "an options object" };
// Byte positions are exact only as safe integers: past 2^53 an addition rounds.
const offset = () =>
  s
    .number({
      integer: true,
      min: 0,
      max: Number.MAX_SAFE_INTEGER,
      expected: "an integer from 0 to Number.MAX_SAFE_INTEGER"
    })
    .optional();
const byteCount = () =>
  s
    .number({
      integer: true,
      min: 1,
      max: Number.MAX_SAFE_INTEGER,
      expected: "an integer from 1 to Number.MAX_SAFE_INTEGER"
    })
    .optional();
const xmlName = s.custom<string>(v => typeof v === "string" && isXmlName(v), "an XML name");

const contextSchema = s.object(
  {
    root: xmlName,
    entities: s.custom<Readonly<Record<string, string>>>(
      v =>
        typeof v === "object" &&
        v !== null &&
        !Array.isArray(v) &&
        Object.keys(v).every(name => {
          const value = (v as Record<string, unknown>)[name];
          return isXmlName(name) && typeof value === "string" && isXmlText(value);
        }),
      "an object of replacement text (a string of XML characters) by entity name"
    ),
    encoding: s.literal("ascii").optional(),
    attributes: s
      .custom<Readonly<Record<string, Readonly<Record<string, XmlAttributeDeclaration>>>>>(
        isAttributeDeclarations,
        "an object of attribute declarations ({ tokenized?: true, default?: string }) by element and attribute name"
      )
      .optional()
  },
  { expected: "a prolog context ({ root, entities, encoding?, attributes? })", readonly: true }
);

/** `{ element: { attribute: { tokenized?: true, default?: string } } }`, every map holding one entry or more. */
function isAttributeDeclarations(v: unknown): boolean {
  if (typeof v !== "object" || v === null || Array.isArray(v)) {
    return false;
  }
  return Object.keys(v).every(element => {
    const declared = (v as Record<string, unknown>)[element];
    if (
      !isXmlName(element) ||
      typeof declared !== "object" ||
      declared === null ||
      Array.isArray(declared) ||
      Object.keys(declared).length === 0
    ) {
      return false;
    }
    return Object.keys(declared).every(name => {
      const declaration = (declared as Record<string, unknown>)[name];
      if (!isXmlName(name) || typeof declaration !== "object" || declaration === null) {
        return false;
      }
      const { tokenized, default: value, ...rest } = declaration as Record<string, unknown>;
      return (
        Object.keys(rest).length === 0 &&
        (tokenized === undefined || tokenized === true) &&
        (value === undefined || (typeof value === "string" && isXmlText(value))) &&
        (tokenized !== undefined || value !== undefined)
      );
    });
  });
}
const _contextMatch: SchemaMatches<typeof contextSchema, XmlPrologContext> = true;
void _contextMatch;

const parseShape = {
  comments: flag(),
  processingInstructions: flag(),
  cdataAsNodes: flag(),
  fragment: flag(),
  xmlns: flag(),
  maxDepth: limit(),
  maxEntityExpansions: limit(),
  maxEntityExpansionLength: limit(),
  invalidCharHandling: invalidCharHandling(),
  internalEntities: flag()
};
const parseOptionsSchema = s.object(parseShape, OPTIONS);
const PARSE_KEYS = Object.keys(parseShape) as Array<keyof XmlParseOptions>;
const _parseOptionsMatch: SchemaMatches<typeof parseOptionsSchema, XmlParseOptions> = true;
void _parseOptionsMatch;

const saxShape = {
  position: flag(),
  fileName: s.string().optional(),
  fragment: flag(),
  xmlns: flag(),
  maxDepth: limit(),
  maxEntityExpansions: limit(),
  maxEntityExpansionLength: limit(),
  invalidCharHandling: invalidCharHandling(),
  byteOffset: offset(),
  resume: flag(),
  internalEntities: flag(),
  maxContextBytes: byteCount(),
  context: contextSchema.optional(),
  skipLeadingText: flag()
};
const saxOptionsSchema = s.object(saxShape, OPTIONS);
const SAX_KEYS = Object.keys(saxShape) as Array<keyof XmlSaxOptions>;
const _saxOptionsMatch: SchemaMatches<typeof saxOptionsSchema, XmlSaxOptions> = true;
void _saxOptionsMatch;

const plainObjectShape = {
  ignoreAttributes: flag(),
  attributePrefix: s.string().optional(),
  textKey: s.string().optional(),
  alwaysArray: flag(),
  isArray: s.function<(name: string) => boolean>().optional(),
  preserveCData: flag(),
  ignoreWhitespaceText: flag()
};
const plainObjectOptionsSchema = s.object(plainObjectShape, OPTIONS);
const PLAIN_OBJECT_KEYS = Object.keys(plainObjectShape) as Array<keyof XmlToPlainObjectOptions>;
const _plainObjectOptionsMatch: SchemaMatches<
  typeof plainObjectOptionsSchema,
  XmlToPlainObjectOptions
> = true;
void _plainObjectOptionsMatch;

const parseToObjectShape = {
  ...plainObjectShape,
  fragment: flag(),
  maxDepth: limit(),
  maxEntityExpansions: limit(),
  maxEntityExpansionLength: limit(),
  invalidCharHandling: invalidCharHandling(),
  internalEntities: flag()
};
const parseToObjectOptionsSchema = s.object(parseToObjectShape, OPTIONS);
const PARSE_TO_OBJECT_KEYS = Object.keys(parseToObjectShape) as Array<
  keyof XmlParseToObjectOptions
>;
const _parseToObjectOptionsMatch: SchemaMatches<
  typeof parseToObjectOptionsSchema,
  XmlParseToObjectOptions
> = true;
void _parseToObjectOptionsMatch;
const readRecordsShape = {
  ...plainObjectShape,
  tags: s.readonlyArray(xmlName, { minLength: 1, expected: "a non-empty array of XML names" }),
  onRecord: s.function<XmlReadRecordsOptions["onRecord"]>(),
  offset: offset(),
  size: byteCount(),
  context: contextSchema.optional(),
  rawXml: flag(),
  maxDepth: limit(),
  maxEntityExpansions: limit(),
  maxEntityExpansionLength: limit(),
  invalidCharHandling: invalidCharHandling(),
  maxContextBytes: byteCount(),
  signal: s.abortSignal().optional(),
  skipLeadingText: flag()
};
const readRecordsOptionsSchema = s.object(readRecordsShape, OPTIONS);
const READ_RECORDS_KEYS = Object.keys(readRecordsShape) as Array<keyof XmlReadRecordsOptions>;
const _readRecordsOptionsMatch: SchemaMatches<
  typeof readRecordsOptionsSchema,
  XmlReadRecordsOptions
> = true;
void _readRecordsOptionsMatch;

const readRootShape = {
  maxStartByte: byteCount(),
  maxTagBytes: byteCount(),
  maxContextBytes: byteCount(),
  maxEntityExpansions: limit(),
  maxEntityExpansionLength: limit(),
  signal: s.abortSignal().optional(),
  skipLeadingText: flag()
};
const readRootOptionsSchema = s.object(readRootShape, OPTIONS);
const READ_ROOT_KEYS = Object.keys(readRootShape) as Array<keyof XmlReadRootOptions>;
const _readRootOptionsMatch: SchemaMatches<typeof readRootOptionsSchema, XmlReadRootOptions> = true;
void _readRootOptionsMatch;
const optionalReadRootOptionsSchema = readRootOptionsSchema.optional();

/**
 * A byte source once checked: the bytes, or the method that opens its iterator — looked up once,
 * so the method checked is the method called.
 */
export type CheckedByteSource =
  | { readonly bytes: Uint8Array }
  | { readonly source: object; readonly open: (this: object) => unknown };

const BYTE_SOURCE = "a Uint8Array, or an iterable or async iterable of Uint8Array chunks";

/**
 * The bytes `Xml.readRecords` and `Xml.readRoot` read. An async iterable's method is taken when it
 * is there at all — one that is there and is not a function is refused, not passed over for a
 * sync one — as `for await` takes it. A function can be an iterable, as any object can.
 */
export function checkByteSource(where: string, source: unknown): CheckedByteSource {
  if (isUint8Array(source)) {
    return { bytes: source };
  }
  const argument = xmlArgument(where, "source");
  if ((typeof source === "object" && source !== null) || typeof source === "function") {
    for (const symbol of [Symbol.asyncIterator, Symbol.iterator]) {
      const open: unknown = (source as Record<symbol, unknown>)[symbol];
      if (typeof open === "function") {
        return { source, open: open as (this: object) => unknown };
      }
      if (open !== undefined && open !== null) {
        throw fail(argument, `[${symbol.description}]`, "a function", open);
      }
    }
  }
  throw fail(argument, "", BYTE_SOURCE, source);
}

function isPartitionStructure(v: unknown): boolean {
  if (typeof v !== "object" || v === null) {
    return false;
  }
  const st = v as Record<string, unknown>;
  const start = st.start;
  const end = st.end;
  if (
    !Number.isSafeInteger(start) ||
    (start as number) < 0 ||
    (end !== null && (!Number.isSafeInteger(end) || (end as number) <= (start as number))) ||
    typeof st.final !== "boolean" ||
    !Array.isArray(st.steps) ||
    !Array.isArray(st.open) ||
    !st.open.every(name => typeof name === "string" && isXmlName(name)) ||
    (st.textOnly !== undefined && st.textOnly !== true) ||
    // `eof` is where the file ends, there exactly when the read reached it.
    (st.final
      ? !Number.isSafeInteger(st.eof) || (st.eof as number) < (start as number)
      : st.eof !== undefined) ||
    (st.follow !== undefined &&
      st.follow !== null &&
      (!Number.isSafeInteger(st.follow) ||
        (end !== null && (st.follow as number) < (end as number))))
  ) {
    return false;
  }
  // Steps in file order, each inside the chunk.
  let last = start as number;
  return st.steps.every(step => {
    if (!Array.isArray(step) || !Number.isSafeInteger(step[1])) {
      return false;
    }
    const byte = step[1] as number;
    const inside = byte >= last && (end === null || byte < (end as number));
    last = byte;
    return (
      inside &&
      (step[0] === "/"
        ? typeof step[2] === "string" && isXmlName(step[2])
        : step[0] === "e" || step[0] === "t" || step[0] === "l")
    );
  });
}

/** What `Xml.verifyPartition` joins: the results of a partition's chunks, each with its `structure`. */
export const partitionResultsSchema = s.readonlyArray(
  s.custom<{ readonly structure: XmlPartitionStructure }>(
    v =>
      typeof v === "object" &&
      v !== null &&
      isPartitionStructure((v as { structure?: unknown }).structure),
    "an Xml.readRecords result, with its structure"
  ),
  { minLength: 1, expected: "a non-empty array of Xml.readRecords results" }
);

/**
 * `Xml.verifyPartition`'s results as one read of them: each result's `structure` and the fields of it
 * the join reads, each read once, so what is checked is what is joined. What is not what it should be
 * is kept as it is, for the check to refuse.
 */
export function snapshotPartitionResults(results: unknown): unknown {
  if (!Array.isArray(results)) {
    return results;
  }
  return Array.from({ length: results.length }, (_, i) => {
    const result: unknown = results[i];
    if (typeof result !== "object" || result === null) {
      return result;
    }
    const structure: unknown = (result as { structure?: unknown }).structure;
    if (typeof structure !== "object" || structure === null) {
      return { structure };
    }
    const { start, end, final, eof, steps, open, follow, textOnly } = structure as Record<
      string,
      unknown
    >;
    return {
      structure: {
        start,
        end,
        final,
        ...(eof !== undefined && { eof }),
        steps: Array.isArray(steps)
          ? Array.from({ length: steps.length }, (_, k) => {
              const step: unknown = steps[k];
              return Array.isArray(step)
                ? Array.from({ length: step.length }, (_, n) => step[n])
                : step;
            })
          : steps,
        open: Array.isArray(open) ? Array.from({ length: open.length }, (_, k) => open[k]) : open,
        ...(follow !== undefined && { follow }),
        ...(textOnly !== undefined && { textOnly })
      }
    };
  });
}

/** The options of `Xml.readRecords`, and the rules that relate them. */
export function checkReadRecordsOptions(where: string, options: unknown): void {
  check(readRecordsOptionsSchema, options, xmlArgument(where, "options"));
  checkContextSize(where, options);
  const { offset, size, context } = options;
  if ((offset ?? 0) + (size ?? 0) > Number.MAX_SAFE_INTEGER) {
    throw fail(
      xmlArgument(where, "options"),
      ".size",
      "a size that keeps offset + size within Number.MAX_SAFE_INTEGER",
      size
    );
  }
  if (offset !== undefined && offset > 0 && context === undefined) {
    throw fail(
      xmlArgument(where, "options"),
      ".context",
      "the file's prolog context (Xml.readRoot(...).context) when options.offset is not 0",
      context
    );
  }
}

/**
 * Options as one read of them: each option the call takes — inherited or not enumerable as well —
 * and its `context`'s, read once, so what is checked is what is used. Nothing else is read. What is
 * not an object is left as it is, for the check to refuse.
 */
function snapshotOptions<T>(options: T, keys: readonly string[]): T {
  if (typeof options !== "object" || options === null || Array.isArray(options)) {
    return options;
  }
  const given = options as Record<string, unknown>;
  const copy: Record<string, unknown> = {};
  for (const key of keys) {
    const value = given[key];
    if (value !== undefined) {
      copy[key] = value;
    }
  }
  const context: unknown = copy.context;
  if (typeof context === "object" && context !== null && !Array.isArray(context)) {
    const { root, entities, encoding, attributes } = context as XmlPrologContext;
    copy.context = {
      root,
      ...(encoding !== undefined && { encoding }),
      entities:
        typeof entities === "object" && entities !== null && !Array.isArray(entities)
          ? copyEntities(entities)
          : entities,
      ...(attributes !== undefined && {
        attributes:
          typeof attributes === "object" && attributes !== null && !Array.isArray(attributes)
            ? copyAttributeDeclarations(attributes)
            : attributes
      })
    };
  }
  return copy as T;
}

/** `Xml.readRecords`' options, read once: see {@link snapshotOptions}. Its `tags` too, element by element. */
export function snapshotReadRecordsOptions(options: XmlReadRecordsOptions): XmlReadRecordsOptions {
  const copy = snapshotOptions(options, READ_RECORDS_KEYS);
  const tags: unknown = copy?.tags;
  if (copy !== options && Array.isArray(tags)) {
    (copy as { tags: readonly string[] }).tags = Array.from(
      { length: tags.length },
      (_, i) => tags[i] as string
    );
  }
  return copy;
}

/** `Xml.readRoot`'s options, read once: see {@link snapshotOptions}. */
export function snapshotReadRootOptions(
  options: XmlReadRootOptions | undefined
): XmlReadRootOptions | undefined {
  return snapshotOptions(options, READ_ROOT_KEYS);
}

/** `Xml.parse`'s options, read once: see {@link snapshotOptions}. */
export function snapshotParseOptions(
  options: XmlParseOptions | undefined
): XmlParseOptions | undefined {
  return snapshotOptions(options, PARSE_KEYS);
}

/** `Xml.parseToObject`'s options, read once: see {@link snapshotOptions}. */
export function snapshotParseToObjectOptions(
  options: XmlParseToObjectOptions | undefined
): XmlParseToObjectOptions | undefined {
  return snapshotOptions(options, PARSE_TO_OBJECT_KEYS);
}

/** `Xml.toPlainObject`'s options, read once: see {@link snapshotOptions}. */
export function snapshotPlainObjectOptions(
  options: XmlToPlainObjectOptions | undefined
): XmlToPlainObjectOptions | undefined {
  return snapshotOptions(options, PLAIN_OBJECT_KEYS);
}

/** `new Xml.SaxParser`'s options, read once: see {@link snapshotOptions}. */
export function snapshotSaxOptions(options: XmlSaxOptions | undefined): XmlSaxOptions | undefined {
  return snapshotOptions(options, SAX_KEYS);
}

/** The entities the check reads — its own enumerable keys — each value read once. */
function copyEntities(entities: Readonly<Record<string, string>>): Record<string, string> {
  const copy = Object.create(null) as Record<string, string>;
  for (const name of Object.keys(entities)) {
    copy[name] = entities[name];
  }
  return copy;
}

/** A copy of a context's attribute declarations, each read once: what is not an object is kept, to be refused. */
function copyAttributeDeclarations(
  attributes: Readonly<Record<string, Readonly<Record<string, XmlAttributeDeclaration>>>>
): unknown {
  const copy = Object.create(null) as Record<string, unknown>;
  for (const element of Object.keys(attributes)) {
    const declared: unknown = attributes[element];
    if (typeof declared !== "object" || declared === null || Array.isArray(declared)) {
      copy[element] = declared;
      continue;
    }
    const names = Object.create(null) as Record<string, unknown>;
    for (const name of Object.keys(declared)) {
      const declaration: unknown = (declared as Record<string, unknown>)[name];
      names[name] =
        typeof declaration === "object" && declaration !== null
          ? Object.assign(Object.create(null) as object, declaration)
          : declaration;
    }
    copy[element] = names;
  }
  return copy;
}

/** The options of `Xml.readRoot`. */
export function checkReadRootOptions(where: string, options: unknown): void {
  check(optionalReadRootOptionsSchema, options, xmlArgument(where, "options"));
}

const optionalPlainObjectOptionsSchema = plainObjectOptionsSchema.optional();
const optionalParseToObjectOptionsSchema = parseToObjectOptionsSchema.optional();
const optionalParseOptionsSchema = parseOptionsSchema.optional();
const optionalSaxOptionsSchema = saxOptionsSchema.optional();

/** The options of `Xml.parse`. */
export function checkParseOptions(where: string, options: unknown): void {
  check(optionalParseOptionsSchema, options, xmlArgument(where, "options"));
}

/** The options of `new Xml.SaxParser`. */
export function checkSaxOptions(where: string, options: unknown): void {
  check(optionalSaxOptionsSchema, options, xmlArgument(where, "options"));
  checkContextSize(where, options);
}

/** A `context` handed in may not be larger than one gathered under the same `maxContextBytes`. */
function checkContextSize(
  where: string,
  options: { context?: XmlPrologContext; maxContextBytes?: number } | undefined
): void {
  const context = options?.context;
  if (context === undefined) {
    return;
  }
  const max = options!.maxContextBytes ?? 65536;
  const size = contextBytes(context);
  if (size > max) {
    throw fail(
      xmlArgument(where, "options"),
      ".context",
      `a prolog context of at most maxContextBytes (${max}) bytes as JSON`,
      `${size} bytes`
    );
  }
}

/** The options of `Xml.toPlainObject`. */
export function checkPlainObjectOptions(where: string, options: unknown): void {
  check(optionalPlainObjectOptionsSchema, options, xmlArgument(where, "options"));
}

/** The options of `Xml.parseToObject`. */
export function checkParseToObjectOptions(where: string, options: unknown): void {
  check(optionalParseToObjectOptionsSchema, options, xmlArgument(where, "options"));
}

const stringSchema = s.string();
const functionSchema = s.function();

/** A string parameter: the text a parser reads, a name, a path. */
export function checkStringArgument(where: string, field: string, value: unknown): void {
  check(stringSchema, value, xmlArgument(where, field));
}

/** The text a parser reads. */
export function checkXmlInput(where: string, xml: unknown): void {
  checkStringArgument(where, "input", xml);
}

const elementSchema = s.custom(
  element =>
    typeof element === "object" &&
    element !== null &&
    (element as { type?: unknown }).type === "element" &&
    Array.isArray((element as { children?: unknown }).children),
  "an XmlElement"
);

/** An element argument of a DOM helper. */
export function checkElement(where: string, element: unknown): void {
  check(elementSchema, element, xmlArgument(where, "element"));
}

/**
 * A node a DOM helper reached while walking the tree it was given, refused when it lacks a field
 * the helper reads. Checked per node as it is reached, so a walk costs no second pass. Each check
 * takes the value already read, so the field checked is the field used: a getter read twice could
 * answer otherwise.
 */
/**
 * Where a node is, below the node a helper was given: a path, or the index of a child among its
 * parent's children — kept as a number so that a walk builds no string for a node that is fine.
 */
export type DomPath = string | number;

export function refuseDom(
  where: string,
  field: string,
  path: DomPath,
  suffix: string,
  expected: string,
  value: unknown
): never {
  const at = typeof path === "number" ? `…children[${path}]` : path;
  throw fail(xmlArgument(where, field), at + suffix, expected, value);
}

/** A node's `type`, read once; refuses what is not a node. */
export function readDomType(where: string, field: string, node: unknown, path: DomPath): string {
  if (typeof node !== "object" || node === null) {
    refuseDom(where, field, path, "", "an XmlNode", node);
  }
  const type = (node as { type?: unknown }).type;
  if (typeof type !== "string") {
    refuseDom(where, field, path, ".type", "a node type", type);
  }
  return type;
}

/** A text or CDATA node's `value`, read once. */
export function readDomText(where: string, field: string, node: object, path: DomPath): string {
  const value = (node as { value?: unknown }).value;
  if (typeof value !== "string") {
    refuseDom(where, field, path, ".value", "a string", value);
  }
  return value;
}

/** An element's `name`, read once. */
export function readDomName(where: string, field: string, node: object, path: DomPath): string {
  const name = (node as { name?: unknown }).name;
  if (typeof name !== "string") {
    refuseDom(where, field, path, ".name", "a string", name);
  }
  return name;
}

/** An element's `children`, read once. */
export function readDomChildren(
  where: string,
  field: string,
  node: object,
  path: DomPath
): readonly unknown[] {
  const children = (node as { children?: unknown }).children;
  if (!Array.isArray(children)) {
    refuseDom(where, field, path, ".children", "an array", children);
  }
  return children;
}

/** An element's `attributes`, read once, for a helper that reads them. */
export function readDomAttributes(
  where: string,
  field: string,
  element: object
): Readonly<Record<string, unknown>> {
  const attributes = (element as { attributes?: unknown }).attributes;
  if (typeof attributes !== "object" || attributes === null || Array.isArray(attributes)) {
    refuseDom(where, field, "", ".attributes", "an object", attributes);
  }
  return attributes as Readonly<Record<string, unknown>>;
}

/** An attribute value of an element a helper was given, which is a string. */
export function checkAttributeValue(
  where: string,
  field: string,
  name: string,
  value: unknown
): asserts value is string {
  if (typeof value !== "string") {
    refuseDom(where, field, "", `.attributes[${JSON.stringify(name)}]`, "a string", value);
  }
}

/** A function argument. */
export function checkFunctionArgument(where: string, field: string, value: unknown): void {
  check(functionSchema, value, xmlArgument(where, field));
}
