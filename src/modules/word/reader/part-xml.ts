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
import type { XmlDocument, XmlElement, XmlNode } from "@xml/types";

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

const CUSTOM_XML_TAG = /<(?:w:)?customXml[\s/>]/;

/**
 * {@link parsePartXml} for a part holding document content (the body, a
 * header or footer, notes, comments), with custom XML markup removed by
 * {@link unwrapCustomXml}.
 */
export function parseContentPartXml(xml: string, maxDepth: number | undefined): XmlDocument {
  const doc = parsePartXml(xml, maxDepth);
  // The regex is the fast test here: it is anchored on "<", which V8 scans
  // for far faster than `includes("customXml")` scans for the whole word
  // (measured 0.9 ms against 3.7 ms over a 6.6 MB body).
  if (CUSTOM_XML_TAG.test(xml)) {
    unwrapCustomXml(doc.root);
  }
  return doc;
}

/** Whether `el` is a `w:customXml` element. */
export function isCustomXmlElement(el: XmlElement): boolean {
  return el.name === "w:customXml" || el.name === "customXml";
}

/**
 * Replace every `w:customXml` under `el` — around blocks, rows, cells or runs —
 * by its content, in place, dropping its `w:customXmlPr`.
 *
 * Custom XML markup tags document content with elements of a caller's schema.
 * Word no longer supports it and removes it from a document it opens, keeping
 * the content; the reader does the same, so the paragraphs, tables, runs and
 * range markers inside are read in place instead of being lost with their
 * wrapper. The custom XML data parts (`customXml/item*.xml`) are unaffected.
 */
export function unwrapCustomXml(el: XmlElement): void {
  if (!el.children.some(c => c.type === "element")) {
    return;
  }
  let changed = false;
  const children: XmlNode[] = [];
  for (const child of el.children) {
    if (child.type === "element" && isCustomXmlElement(child)) {
      changed = true;
      children.push(...customXmlContent(child));
      continue;
    }
    if (child.type === "element") {
      unwrapCustomXml(child);
    }
    children.push(child);
  }
  if (changed) {
    el.children = children;
  }
}

/** The content of a `w:customXml` element, itself unwrapped. */
export function customXmlContent(el: XmlElement): XmlNode[] {
  const wrapper: XmlElement = {
    ...el,
    children: el.children.filter(
      c => !(c.type === "element" && (c.name === "w:customXmlPr" || c.name === "customXmlPr"))
    )
  };
  unwrapCustomXml(wrapper);
  return wrapper.children;
}
