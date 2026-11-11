/**
 * DOCX Reader - XML parsing of package parts under the security policy.
 *
 * Every part the reader parses goes through {@link parsePartXml} with the
 * policy's `maxXmlDepth`, so the limit reaches the body, headers, notes,
 * styles, charts and the auxiliary parts alike.
 */

import { trimXmlPartLead } from "@utils/xml-part";
import { DocxLimitExceededError } from "@word/errors";
import { WORD_NAMESPACES } from "@word/reader/namespaces";
import { parseXml } from "@xml/dom";
import { XmlParseError } from "@xml/errors";
import { needsPrefixNormalization, normalizeElementPrefixes } from "@xml/namespace-prefixes";
import type { XmlDocument, XmlElement, XmlNode } from "@xml/types";

/**
 * `parseXml` with a nesting-depth limit (`undefined` = the XML module's
 * default). A depth failure is reported as `DocxLimitExceededError("xmlDepth", …)`,
 * like the reader's other resource limits.
 */
export function parsePartXml(xml: string, maxDepth: number | undefined): XmlDocument {
  return parsePart(xml, maxDepth).doc;
}

/** {@link parsePartXml}, and the text it parsed. */
function parsePart(xml: string, maxDepth: number | undefined): { doc: XmlDocument; text: string } {
  // From its first `<`: white space before the XML declaration is refused by XML, opened by Office.
  const text = trimXmlPartLead(xml);
  let doc: XmlDocument;
  try {
    doc = parseXml(text, maxDepth === undefined ? undefined : { maxDepth });
  } catch (error) {
    throw asPartLimitError(error, maxDepth);
  }
  // Its names in the prefixes the reader matches, whatever prefixes the part binds (`namespaces.ts`),
  // renamed in the tree the one parse built — so a part is read under the policy above whichever
  // prefixes it uses. Each element takes its text view, as the streaming reader's tags do: a subtree
  // the reader keeps is written back from it. A part that already uses them — every part Word
  // writes — is left as it is.
  if (needsPrefixNormalization(WORD_NAMESPACES, text)) {
    const texts = normalizeElementPrefixes(WORD_NAMESPACES, doc.root);
    for (const [element, tag] of texts) {
      element.name = tag.name;
      element.attributes = tag.attributes;
    }
  }
  return { doc, text };
}

/**
 * What a parse of a part threw, as the reader reports it: a nesting depth past `maxDepth` is a
 * `DocxLimitExceededError("xmlDepth", …)`, like the reader's other resource limits; anything else
 * is returned as it is.
 */
export function asPartLimitError(error: unknown, maxDepth: number | undefined): unknown {
  if (error instanceof XmlParseError && error.limit === "depth" && maxDepth !== undefined) {
    return new DocxLimitExceededError(
      "xmlDepth",
      maxDepth,
      maxDepth + 1,
      "XML element nesting exceeds maxXmlDepth"
    );
  }
  return error;
}

/**
 * Where a `customXml` start tag may be, under any prefix: the part's own text is tested, before its
 * prefixes are renamed. A literal start lets V8 skip through the text, where matching the prefix at
 * every `<` took seven times as long; a match that is no tag only costs the walk that finds none.
 */
const CUSTOM_XML_TAG = /customXml[\s/>]/;

/**
 * {@link parsePartXml} for a part holding document content (the body, a
 * header or footer, notes, comments), with custom XML markup removed by
 * {@link unwrapCustomXml}.
 */
export function parseContentPartXml(xml: string, maxDepth: number | undefined): XmlDocument {
  const { doc, text } = parsePart(xml, maxDepth);
  // The regex is the fast test here: measured 0.8 ms over a 9 MB body, against 4.7 ms for
  // `includes("customXml")`.
  if (CUSTOM_XML_TAG.test(text)) {
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
