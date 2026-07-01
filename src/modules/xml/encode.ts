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

import { XmlError } from "@xml/errors";

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

/**
 * Validate an XML element or attribute name against injection attacks.
 *
 * Rejects:
 * - Empty names
 * - Names containing whitespace, `<`, `>`, `"`, `'`, `/`, `=`, `&`
 * - Names starting with a digit, `-`, or `.`
 *
 * This is NOT a full XML Name validation (which requires Unicode NameStartChar
 * tables). It is a focused security check to prevent markup injection.
 */
export function validateXmlName(name: string): void {
  if (!name) {
    throw new XmlError("XML name must not be empty");
  }
  if (INVALID_NAME_CHARS.test(name)) {
    throw new XmlError(`Invalid XML name: contains forbidden character in "${name}"`);
  }
  // XML names cannot start with a digit, hyphen, or dot
  const first = name.charCodeAt(0);
  if ((first >= 0x30 && first <= 0x39) || first === 0x2d || first === 0x2e) {
    throw new XmlError(`Invalid XML name: "${name}" starts with forbidden character`);
  }
}

// =============================================================================
// Comment Validation
// =============================================================================

/**
 * Validate that text is legal for an XML comment.
 *
 * XML spec: comments must not contain `--` and must not end with `-`.
 * @throws {XmlError} if the text is invalid.
 */
export function validateCommentText(text: string): void {
  if (text.includes("--") || text.endsWith("-")) {
    throw new XmlError('Invalid comment: must not contain "--" or end with "-"');
  }
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
