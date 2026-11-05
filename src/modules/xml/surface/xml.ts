/**
 * `Xml` namespace surface — XML encode/decode, DOM + SAX parsing, writers,
 * and query.
 *
 * `import { Xml } from "documonster/xml"` →
 *   `Xml.encode(s)`, `Xml.parse(src)`, `Xml.query(doc, sel)`,
 *   `Xml.findChild(el, name)` / `Xml.listChildren(el, name)`,
 *   `new Xml.Writer()`, `new Xml.StreamWriter(target)`.
 *
 * Two ways to read XML as events, by who drives: `Xml.parseSax(source)` is pull — it
 * yields batches of events for a `for await` loop; `new Xml.SaxParser()` is push — you
 * register `.on(...)` handlers, then `Xml.feedSax(parser, source)` feeds it every chunk of
 * a source (or call `parser.write()` / `parser.close()` yourself).
 *
 * Single flat namespace (xml is a single-purpose module). Re-exported via
 * `export * as Xml`, tree-shaken per-member on rolldown / rspack.
 */
import {
  attr as attrImpl,
  findChildren as findChildrenImpl,
  textContent as textContentImpl,
  walk as walkImpl
} from "@xml/dom";
import { validateXmlName as validateXmlNameImpl } from "@xml/encode";
import { saxStream as saxStreamImpl } from "@xml/sax";
import { StdDocAttributes as StdDocAttributesImpl } from "@xml/writer";

// Encoding / decoding
export {
  encodeText as encode,
  decodeText as decode,
  encodeAttributeValue as encodeAttr,
  encodeCDataSection as encodeCData,
  validateXmlName as validateName,
  validateCommentText
} from "@xml/encode";

// Writers
export { XmlWriter as Writer } from "@xml/writer";
/** The declaration `Writer.openXml()` writes when given no attributes: version 1.0, UTF-8, standalone. */
export { StdDocAttributes as DEFAULT_DECLARATION } from "@xml/writer";
export { XmlStreamWriter as StreamWriter } from "@xml/stream-writer";

// Parsers
export { SaxParser, parseSax, saxStream as feedSax } from "@xml/sax";
export {
  parseXml as parse,
  findChildOf as findChild,
  listChildrenOf as listChildren,
  getTextOf as getText,
  getAttrOf as getAttr,
  eachElementOf as eachElement,
  toPlainObject
} from "@xml/dom";
export { parseXmlToObject as parseToObject } from "@xml/to-object";
export { query, queryAll } from "@xml/query";

/** @deprecated Use {@link listChildren}. Removed in 1.0. */
export const findChildren = findChildrenImpl;
/** @deprecated Use {@link validateName}. Removed in 1.0. */
export const validateXmlName = validateXmlNameImpl;
/** @deprecated Use {@link getText}. Removed in 1.0. */
export const textContent = textContentImpl;
/** @deprecated Use {@link getAttr}. Removed in 1.0. */
export const attr = attrImpl;
/** @deprecated Use {@link eachElement}. Removed in 1.0. */
export const walk = walkImpl;
/** @deprecated Use {@link feedSax}. Removed in 1.0. */
export const saxStream = saxStreamImpl;

/** @deprecated Use {@link DEFAULT_DECLARATION}. Removed in 1.0. */
export const StdDocAttributes = StdDocAttributesImpl;
