/**
 * The namespaces of the content models in the generated sequence tables, by the prefix their types are keyed with
 * (`x:CT_Worksheet`, `w:CT_PPr`, `x14:CT_DataBar`). A child is ranked by a model only when it is in that model's
 * namespace: matched by local name alone, an extension element that happens to share a name with a standard one
 * (`ext:a` beside `a`) was ranked as if it were the standard one. Each ECMA-376 namespace is listed in its Transitional
 * and its Strict form, which carry the same content models.
 */
export const OOXML_SEQUENCE_NAMESPACES: Readonly<Record<string, readonly string[]>> = {
  x: [
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "http://purl.oclc.org/ooxml/spreadsheetml/main"
  ],
  w: [
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://purl.oclc.org/ooxml/wordprocessingml/main"
  ],
  a: [
    "http://schemas.openxmlformats.org/drawingml/2006/main",
    "http://purl.oclc.org/ooxml/drawingml/main"
  ],
  c: [
    "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "http://purl.oclc.org/ooxml/drawingml/chart"
  ],
  xdr: [
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"
  ],
  wp: [
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing"
  ],
  pic: [
    "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "http://purl.oclc.org/ooxml/drawingml/picture"
  ],
  m: [
    "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "http://purl.oclc.org/ooxml/officeDocument/math"
  ],
  sl: [
    "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
    "http://purl.oclc.org/ooxml/schemaLibrary/main"
  ],
  x14: ["http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"],
  xm: ["http://schemas.microsoft.com/office/excel/2006/main"],
  w14: ["http://schemas.microsoft.com/office/word/2010/wordml"],
  w15: ["http://schemas.microsoft.com/office/word/2012/wordml"],
  wps: ["http://schemas.microsoft.com/office/word/2010/wordprocessingShape"]
};
