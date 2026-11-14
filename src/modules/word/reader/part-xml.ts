/**
 * DOCX Reader - XML parsing of package parts under the security policy.
 *
 * Every part the reader parses goes through {@link parsePartXml} with the
 * policy's `maxXmlDepth`, so the limit reaches the body, headers, notes,
 * styles, charts and the auxiliary parts alike.
 */

import { appendAll } from "@utils/array";
import { trimXmlPartLead } from "@utils/xml-part";
import { paceChunks, runSteps, runStepsPaced } from "@utils/yield";
import type { Steps, YieldPoint } from "@utils/yield";
import { DocxLimitExceededError } from "@word/errors";
import { WORD_NAMESPACES } from "@word/reader/namespaces";
import { createDomBuilder } from "@xml/dom";
import { XmlParseError } from "@xml/errors";
import { needsPrefixNormalization, normalizeElementPrefixesSteps } from "@xml/namespace-prefixes";
import type { ParseOpenTag, XmlDocument, XmlElement, XmlNode } from "@xml/types";

/**
 * `parseXml` with a nesting-depth limit (`undefined` = the XML module's
 * default). A depth failure is reported as `DocxLimitExceededError("xmlDepth", …)`,
 * like the reader's other resource limits.
 */
export function parsePartXml(xml: string, maxDepth: number | undefined): XmlDocument {
  return parsePart(xml, maxDepth).doc;
}

/**
 * {@link parsePartXml} a slice of the text at a time, asking `yieldPoint` (a `createYieldBudget`) between the slices —
 * for the streaming reader, which must give the event loop a turn while it parses a large part.
 *
 * Not the form every reader takes: writing the text in slices costs a whole-part read about a tenth more, measured on
 * a 5 MB body, which a reader that does not yield pays for nothing.
 */
export async function parsePartXmlPaced(
  xml: string,
  maxDepth: number | undefined,
  yieldPoint: YieldPoint
): Promise<XmlDocument> {
  return (await parsePartPaced(xml, maxDepth, yieldPoint)).doc;
}

/** {@link parsePartXml}, and the text it parsed. */
function parsePart(xml: string, maxDepth: number | undefined): { doc: XmlDocument; text: string } {
  // From its first `<`: white space before the XML declaration is refused by XML, opened by Office.
  const text = trimXmlPartLead(xml);
  const builder = partBuilder(maxDepth);
  try {
    builder.write(text);
    return runSteps(finishPart(builder.close(), text));
  } catch (error) {
    throw asPartLimitError(error, maxDepth);
  }
}

/** {@link parsePart}, a slice at a time — see {@link parsePartXmlPaced}. */
async function parsePartPaced(
  xml: string,
  maxDepth: number | undefined,
  yieldPoint: YieldPoint
): Promise<{ doc: XmlDocument; text: string }> {
  const text = trimXmlPartLead(xml);
  const builder = partBuilder(maxDepth);
  try {
    for await (const slice of paceChunks([text], yieldPoint)) {
      builder.write(slice);
    }
    return await runStepsPaced(finishPart(builder.close(), text), yieldPoint);
  } catch (error) {
    throw asPartLimitError(error, maxDepth);
  }
}

function partBuilder(maxDepth: number | undefined): ReturnType<typeof createDomBuilder> {
  return createDomBuilder(maxDepth === undefined ? undefined : { maxDepth });
}

/**
 * The part's names in the prefixes the reader matches, whatever prefixes the part binds (`namespaces.ts`), renamed in
 * the tree the one parse built — so a part is read under the policy above whichever prefixes it uses. Each element
 * takes its text view, as the streaming reader's tags do: a subtree the reader keeps is written back from it. A part
 * that already uses them — every part Word writes — is left as it is.
 */
function* finishPart(doc: XmlDocument, text: string): Steps<{ doc: XmlDocument; text: string }> {
  if (needsPrefixNormalization(WORD_NAMESPACES, text)) {
    yield* renamePrefixes(yield* normalizeElementPrefixesSteps(WORD_NAMESPACES, doc.root));
  }
  return { doc, text };
}

/** Give each element the text view {@link finishPart} describes, in steps of {@link RENAME_STEP} elements. */
function* renamePrefixes(
  texts: ReadonlyMap<{ name: string; attributes: Record<string, string> }, ParseOpenTag>
): Steps<void> {
  let renamed = 0;
  for (const [element, tag] of texts) {
    element.name = tag.name;
    element.attributes = tag.attributes;
    if (++renamed % RENAME_STEP === 0) {
      yield;
    }
  }
}

/** How many elements {@link renamePrefixes} renames in one step. */
const RENAME_STEP = 4096;

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
  return withoutCustomXml(parsePart(xml, maxDepth));
}

/** {@link parseContentPartXml} a slice at a time — see {@link parsePartXmlPaced}. */
export async function parseContentPartXmlPaced(
  xml: string,
  maxDepth: number | undefined,
  yieldPoint: YieldPoint
): Promise<XmlDocument> {
  return withoutCustomXml(await parsePartPaced(xml, maxDepth, yieldPoint));
}

function withoutCustomXml({ doc, text }: { doc: XmlDocument; text: string }): XmlDocument {
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
      appendAll(children, customXmlContent(child));
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
