/**
 * `Xml` namespace surface — XML encode/decode, DOM + SAX parsing, writers,
 * and query.
 *
 * `import { Xml } from "documonster/xml"` →
 *   `Xml.encode(s)`, `Xml.parse(src)`, `Xml.query(doc, sel)`,
 *   `new Xml.Writer()`, `new Xml.StreamWriter()`, `Xml.parseSax(...)`.
 *
 * Single flat namespace (xml is a single-purpose module). Re-exported via
 * `export * as Xml`, tree-shaken per-member on rolldown / rspack.
 */

// Encoding / decoding
export {
  xmlEncode as encode,
  xmlDecode as decode,
  xmlEncodeAttr as encodeAttr,
  validateXmlName,
  encodeCData,
  validateCommentText
} from "../encode";

// Writers
export { XmlWriter as Writer, StdDocAttributes } from "../writer";
export { XmlStreamWriter as StreamWriter } from "../stream-writer";

// Parsers
export { SaxParser, parseSax, saxStream } from "../sax";
export {
  parseXml as parse,
  findChild,
  findChildren,
  textContent,
  attr,
  walk,
  toPlainObject
} from "../dom";
export { parseXmlToObject as parseToObject } from "../to-object";
export { query, queryAll } from "../query";
