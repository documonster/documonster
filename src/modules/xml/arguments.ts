/**
 * Argument checks for the `Xml` entry points. A value with no reading is an {@link XmlError}
 * naming the call and the option, raised before anything is parsed or written: `maxDepth:
 * NaN` used to switch the nesting limit off without a word, and a `null` input surfaced as a
 * `TypeError` from deep inside the parser.
 *
 * The options are declared as schemas (`@utils/schema`), each pinned to its interface.
 */

import { check, s } from "@utils/schema";
import type { SchemaCheckOptions, SchemaMatches } from "@utils/schema";
import { XmlArgumentError } from "@xml/errors";
import type {
  XmlParseOptions,
  XmlParseToObjectOptions,
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

const parseOptionsSchema = s.object(
  {
    comments: flag(),
    processingInstructions: flag(),
    cdataAsNodes: flag(),
    fragment: flag(),
    xmlns: flag(),
    maxDepth: limit(),
    maxEntityExpansions: limit(),
    invalidCharHandling: invalidCharHandling()
  },
  OPTIONS
);
const _parseOptionsMatch: SchemaMatches<typeof parseOptionsSchema, XmlParseOptions> = true;
void _parseOptionsMatch;

const saxOptionsSchema = s.object(
  {
    position: flag(),
    fileName: s.string().optional(),
    fragment: flag(),
    xmlns: flag(),
    maxDepth: limit(),
    maxEntityExpansions: limit(),
    invalidCharHandling: invalidCharHandling()
  },
  OPTIONS
);
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
const _plainObjectOptionsMatch: SchemaMatches<
  typeof plainObjectOptionsSchema,
  XmlToPlainObjectOptions
> = true;
void _plainObjectOptionsMatch;

const parseToObjectOptionsSchema = s.object(
  {
    ...plainObjectShape,
    fragment: flag(),
    maxDepth: limit(),
    maxEntityExpansions: limit(),
    invalidCharHandling: invalidCharHandling()
  },
  OPTIONS
);
const _parseToObjectOptionsMatch: SchemaMatches<
  typeof parseToObjectOptionsSchema,
  XmlParseToObjectOptions
> = true;
void _parseToObjectOptionsMatch;
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

/** A function argument. */
export function checkFunctionArgument(where: string, field: string, value: unknown): void {
  check(functionSchema, value, xmlArgument(where, field));
}
