/**
 * Argument checks for the `Xml` entry points. A value with no reading is an {@link XmlError}
 * naming the call and the option, raised before anything is parsed or written: `maxDepth:
 * NaN` used to switch the nesting limit off without a word, and a `null` input surfaced as a
 * `TypeError` from deep inside the parser.
 */

import { describeValue as describe } from "@utils/describe-value";
import { XmlArgumentError } from "@xml/errors";

/** The argument path an option name stands for: `maxDepth` → `options.maxDepth`. */
function fieldOf(option: string): string {
  return option === "options" || option === "input" || option === "element"
    ? option
    : `options.${option}`;
}

function fail(where: string, option: string, expected: string, value: unknown): never {
  throw new XmlArgumentError(
    fieldOf(option),
    `${where}: ${option} must be ${expected}, got ${describe(value)}`
  );
}

/** An options argument: absent, or a plain object. */
export function checkOptionsObject(where: string, options: unknown): void {
  if (
    options !== undefined &&
    (typeof options !== "object" || options === null || Array.isArray(options))
  ) {
    fail(where, "options", "an options object", options);
  }
}

/** The text a parser reads. */
export function checkXmlInput(where: string, xml: unknown): void {
  if (typeof xml !== "string") {
    fail(where, "input", "a string", xml);
  }
}

/** An element argument of a DOM helper. */
export function checkElement(where: string, element: unknown): void {
  if (
    typeof element !== "object" ||
    element === null ||
    (element as { type?: unknown }).type !== "element" ||
    !Array.isArray((element as { children?: unknown }).children)
  ) {
    fail(where, "element", "an XmlElement", element);
  }
}

export function checkBooleanOptions(
  where: string,
  options: Record<string, unknown> | undefined,
  names: readonly string[]
): void {
  if (!options) {
    return;
  }
  for (const name of names) {
    const value = options[name];
    if (value !== undefined && typeof value !== "boolean") {
      fail(where, name, "true or false", value);
    }
  }
}

/**
 * The parser limits and character handling every parse entry shares. A limit is an integer
 * `>= 0` (`0` switches it off, as documented); anything else would have switched it off too.
 */
export function checkParserLimits(
  where: string,
  options: Record<string, unknown> | undefined
): void {
  if (!options) {
    return;
  }
  for (const name of ["maxDepth", "maxEntityExpansions"]) {
    const value = options[name];
    if (
      value !== undefined &&
      (typeof value !== "number" || !Number.isInteger(value) || value < 0)
    ) {
      fail(where, name, "an integer >= 0 (0 switches the limit off)", value);
    }
  }
  const handling = options.invalidCharHandling;
  if (
    handling !== undefined &&
    handling !== "error" &&
    handling !== "skip" &&
    handling !== "replace"
  ) {
    fail(where, "invalidCharHandling", '"error", "skip" or "replace"', handling);
  }
}
