/**
 * DOCX Reader - XML parsing of package parts under the security policy.
 *
 * Every part the reader parses goes through {@link parsePartXml} with the
 * policy's `maxXmlDepth`, so the limit reaches the body, headers, notes,
 * styles, charts and the auxiliary parts alike.
 */

import { DocxLimitExceededError } from "@word/errors";
import { parseXml } from "@xml/dom";
import { XmlParseError } from "@xml/errors";
import type { XmlDocument } from "@xml/types";

/**
 * `parseXml` with a nesting-depth limit (`undefined` = the XML module's
 * default). A depth failure is reported as `DocxLimitExceededError("xmlDepth", …)`,
 * like the reader's other resource limits.
 */
export function parsePartXml(xml: string, maxDepth: number | undefined): XmlDocument {
  try {
    return parseXml(xml, maxDepth === undefined ? undefined : { maxDepth });
  } catch (error) {
    if (error instanceof XmlParseError && error.limit === "depth" && maxDepth !== undefined) {
      throw new DocxLimitExceededError(
        "xmlDepth",
        maxDepth,
        maxDepth + 1,
        "XML element nesting exceeds maxXmlDepth"
      );
    }
    throw error;
  }
}
