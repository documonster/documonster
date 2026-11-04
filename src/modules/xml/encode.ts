/**
 * XML Encoding / Decoding Utilities
 *
 * The character-level encode/decode primitives now live in
 * `@utils/xml-encode` (Layer 0) so dependency-restricted modules — notably
 * `pdf`, which may not import from the `xml` module — can share the exact
 * same canonical implementation. This file re-exports them under the stable
 * `@xml/encode` surface and adds the name / comment validators that depend on
 * `XmlError`.
 */

import {
  encodeCData,
  stripXmlIllegalChars,
  xmlDecode,
  xmlEncode,
  xmlEncodeAttr
} from "@utils/xml-encode";
import { XmlArgumentError } from "@xml/errors";

export {
  isLegalXmlChar,
  xmlDecode,
  xmlEncode,
  xmlEncodeAttr,
  stripXmlIllegalChars,
  encodeCData
} from "@utils/xml-encode";

// =============================================================================
// Name Validation
// =============================================================================

/**
 * Characters that must NEVER appear in XML element or attribute names.
 * This is a fast security check to prevent markup injection via names,
 * not a full XML NameChar validation (which would require Unicode tables).
 */
const INVALID_NAME_CHARS = /[\s<>"'/=&]/;

/** XML 1.0 (5th ed.) §2.3 NameStartChar, without ":" (Namespaces in XML: an NCName's). */
const NC_NAME_START =
  "A-Z_a-z\\u00C0-\\u00D6\\u00D8-\\u00F6\\u00F8-\\u02FF\\u0370-\\u037D\\u037F-\\u1FFF\\u200C-\\u200D" +
  "\\u2070-\\u218F\\u2C00-\\u2FEF\\u3001-\\uD7FF\\uF900-\\uFDCF\\uFDF0-\\uFFFD\\u{10000}-\\u{EFFFF}";
/** §2.3 NameChar, without ":". */
const NC_NAME_CHAR = `${NC_NAME_START}\\-.0-9\\u00B7\\u0300-\\u036F\\u203F-\\u2040`;
const NC_NAME = `[${NC_NAME_START}][${NC_NAME_CHAR}]*`;
/**
 * A QName (Namespaces in XML 1.0 §4): an NCName, or two joined by one ":". Every name these
 * writers produce is read by a namespace-aware consumer, which refuses `a:b:c` or `:a`.
 */
const QNAME_RE = new RegExp(`^${NC_NAME}(?::${NC_NAME})?$`, "u");

/** Names that have passed {@link validateXmlName}, so the common ones are tested once. */
const VALID_NAMES = new Set<string>();
const VALID_NAMES_CAP = 4096;

/**
 * Validate an XML element or attribute name.
 *
 * Accepts a QName: an XML 1.0 `Name` (§2.3) with at most one `:`, between two non-empty
 * parts, as Namespaces in XML requires. Anything else is refused with an {@link XmlError} —
 * including `!` and `?`, which would let a "name" open a comment, a CDATA section or a
 * processing instruction (`<!--x>`), and control characters.
 */
export function validateXmlName(name: string): void {
  validateXmlNameAt(name, "name");
}

/**
 * {@link validateXmlName}, reporting `field` on refusal — an attribute name is checked as
 * `attributes.<key>` of the call that passed the attribute map.
 */
export function validateXmlNameAt(name: string, field: string): void {
  // A writer names the same few elements and attributes over and over — `c`, `r`, `s`, `t`, `v` for every cell of a
  // sheet — so a name that has passed once is not tested again. Capped, so names a caller invents cannot grow it
  // without bound; past the cap a name is simply checked every time, as before.
  if (VALID_NAMES.has(name)) {
    return;
  }
  checkXmlName(name, field);
  if (VALID_NAMES.size < VALID_NAMES_CAP) {
    VALID_NAMES.add(name);
  }
}

function checkXmlName(name: string, field: string): void {
  if (typeof name !== "string") {
    throw new XmlArgumentError(
      field,
      `XML name must be a string, got ${name === null ? "null" : typeof name}`
    );
  }
  if (!name) {
    throw new XmlArgumentError(field, "XML name must not be empty");
  }
  if (INVALID_NAME_CHARS.test(name)) {
    throw new XmlArgumentError(
      field,
      `Invalid XML name: contains forbidden character in "${name}"`
    );
  }
  // XML names cannot start with a digit, hyphen, or dot
  const first = name.charCodeAt(0);
  if ((first >= 0x30 && first <= 0x39) || first === 0x2d || first === 0x2e) {
    throw new XmlArgumentError(
      field,
      `Invalid XML name: "${name}" starts with forbidden character`
    );
  }
  // The full production: a `!` or `?` let a name open a comment, a CDATA section or a
  // processing instruction (`<!--x>`), and a control character made the document unreadable.
  if (!QNAME_RE.test(name)) {
    throw new XmlArgumentError(
      field,
      `Invalid XML name: "${name}" is not an XML name (XML 1.0 Name, with at most one ":" between two parts)`
    );
  }
}

/**
 * Stringify an attribute value or text node, refusing a non-finite number and
 * anything that is not a string, number or boolean.
 *
 * `String(NaN)` is `"NaN"` and `String(Infinity)` is `"Infinity"`, and neither is a
 * legal lexical form for any numeric XSD type — so a non-finite number reaching a
 * writer always produces a document its consumer rejects. It was reached through
 * `IntegerXform`, whose `if (model || this.zero)` treats `Infinity` as a value worth
 * writing, and it emitted `<sz val="Infinity"/>` into a stylesheet Excel then refused
 * to open. Guarding here rather than in that one xform covers every other raw numeric
 * attribute the same way — a colour's `theme`/`tint`, a gradient's `degree`, a stop's
 * `position` — instead of waiting for each to be reported separately.
 *
 * Throwing, rather than omitting the attribute, matches {@link validateXmlName}: this
 * writer already refuses input it cannot represent instead of quietly emitting
 * something broken. The caller gets a stack trace at the point of the write, which is
 * more use than a file that opens to a repair dialog.
 */
export function xmlValue(
  value: string | number | boolean,
  name: string,
  field: string = name === "#text" ? "text" : name
): string {
  // The type admits only these three, but a model built by hand or from JSON is
  // not checked by the compiler: `null` became `"null"` and an object
  // `"[object Object]"` — Word refused a drawing written with `cx="[object Object]"`.
  switch (typeof value) {
    case "string":
      return value;
    case "boolean":
      return String(value);
    case "number":
      if (Number.isFinite(value)) {
        return String(value);
      }
      throw new XmlArgumentError(
        field,
        `Invalid XML value for "${name}": ${String(value)} has no valid XML representation`
      );
    default:
      throw new XmlArgumentError(
        field,
        `Invalid XML value for "${name}": ${value === null ? "null" : `a value of type ${typeof value}`} has no XML representation`
      );
  }
}

// =============================================================================
// Comment Validation
// =============================================================================

/** Which XML comment rule `text` breaks (XML 1.0 §2.5), if any. */
function commentTextProblem(text: string): string | undefined {
  if (text.includes("--")) {
    return 'it contains "--"';
  }
  if (text.endsWith("-")) {
    return 'it ends with "-"';
  }
  return undefined;
}

/**
 * Validate that text is legal for an XML comment.
 *
 * XML spec: comments must not contain `--` and must not end with `-`.
 * @throws {XmlError} if the text is invalid.
 */
export function validateCommentText(text: string): void {
  if (typeof text !== "string") {
    throw new XmlArgumentError(
      "text",
      `Invalid comment: expected a string, got ${text === null ? "null" : typeof text}`
    );
  }
  const problem = commentTextProblem(text);
  if (problem !== undefined) {
    throw new XmlArgumentError("text", `Invalid comment: ${problem}`);
  }
}

/**
 * The ` name="value"` text of a set of attributes, checking every name and value first.
 *
 * Built as one string so a writer can check everything before it changes anything: writers
 * that validated each attribute as they appended it left a half-written start tag behind
 * when a later one was refused.
 */
export function serializeXmlAttributes(
  attributes: Readonly<Record<string, string | number | boolean | undefined>> | undefined
): string {
  if (attributes === undefined || attributes === null) {
    return "";
  }
  if (typeof attributes !== "object") {
    throw new XmlArgumentError(
      "attributes",
      `Invalid XML attributes: expected an object, got ${typeof attributes}`
    );
  }
  let s = "";
  for (const key in attributes) {
    const value = attributes[key];
    if (value !== undefined) {
      validateXmlNameAt(key, `attributes.${key}`);
      s += ` ${key}="${xmlEncodeAttr(xmlValue(value, key, `attributes.${key}`))}"`;
    }
  }
  return s;
}

/**
 * Throw unless every attribute in `attributes` is new on an open element whose start tag so
 * far is `openTag`. XML forbids a repeated attribute name on one element (well-formedness
 * constraint "Unique Att Spec"), and a parser refuses the document.
 *
 * Searching the written tag is exact: names are validated to contain no space, `=` or quote,
 * and an attribute value is written with `"` escaped, so ` name="` occurs only where the
 * attribute itself does.
 */
export function checkNewAttributes(
  element: string | undefined,
  openTag: string,
  attributes: Readonly<Record<string, unknown>> | undefined,
  field: (key: string) => string = key => `attributes.${key}`
): void {
  if (attributes === undefined || attributes === null) {
    return;
  }
  for (const key in attributes) {
    if (attributes[key] !== undefined && openTag.includes(` ${key}="`)) {
      throw new XmlArgumentError(
        field(key),
        `Duplicate attribute "${key}" on element <${element ?? ""}>`
      );
    }
  }
}

/**
 * A comment's markup. Characters XML cannot carry are removed, as they are from text and
 * CDATA; what remains must not contain `--` or end with `-`.
 */
export function commentMarkup(text: string): string {
  if (typeof text !== "string") {
    throw new XmlArgumentError(
      "text",
      `Invalid comment: expected a string, got ${text === null ? "null" : typeof text}`
    );
  }
  const clean = stripXmlIllegalChars(text);
  const problem = commentTextProblem(clean);
  if (problem !== undefined) {
    // Name the rule that failed, and say when it failed only after characters XML cannot hold
    // were removed (`"a\u0001-"` ends with "-" once the control character is gone).
    throw new XmlArgumentError(
      "text",
      `Invalid comment: ${problem}${clean === text ? "" : " once the characters XML cannot hold are removed"}`
    );
  }
  return `<!--${clean}-->`;
}

/** A CDATA section's markup; `]]>` inside the text is split across two sections. */
export function cdataMarkup(text: string): string {
  if (typeof text !== "string") {
    throw new XmlArgumentError(
      "text",
      `Invalid CDATA: expected a string, got ${text === null ? "null" : typeof text}`
    );
  }
  return encodeCData(text);
}

/** Raw markup for `writeRaw`, refused unless it is a string — `undefined` was written as nothing, `42` as `42`. */
export function rawMarkup(xml: string): string {
  if (typeof xml !== "string") {
    throw new XmlArgumentError(
      "xml",
      `Invalid raw XML: expected a string, got ${xml === null ? "null" : typeof xml}`
    );
  }
  return xml;
}

// =============================================================================
// Standard XML Declaration
// =============================================================================

/** Default XML declaration attributes (`version`, `encoding`, `standalone`). */
export const StdDocAttributes: Readonly<Record<string, string>> = {
  version: "1.0",
  encoding: "UTF-8",
  standalone: "yes"
};

// =============================================================================
// Checked public forms
// =============================================================================

function checkText(where: string, text: unknown): asserts text is string {
  if (typeof text !== "string") {
    throw new XmlArgumentError(
      "text",
      `${where}: text must be a string, got ${text === null ? "null" : typeof text}`
    );
  }
}

/**
 * `Xml.encode`: {@link xmlEncode}, refusing a value that is not a string. The unchecked form
 * returned a number unchanged and threw a `TypeError` on `null`.
 */
export function encodeText(text: string): string {
  checkText("Xml.encode", text);
  return xmlEncode(text);
}

/** `Xml.decode`: {@link xmlDecode}, refusing a value that is not a string. */
export function decodeText(text: string): string {
  checkText("Xml.decode", text);
  return xmlDecode(text);
}

/** `Xml.encodeAttr`: {@link xmlEncodeAttr}, refusing a value that is not a string. */
export function encodeAttributeValue(value: string): string {
  checkText("Xml.encodeAttr", value);
  return xmlEncodeAttr(value);
}

/** `Xml.encodeCData`: {@link encodeCData}, refusing a value that is not a string. */
export function encodeCDataSection(text: string): string {
  checkText("Xml.encodeCData", text);
  return encodeCData(text);
}
