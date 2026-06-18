/**
 * `Convert` namespace surface — format conversions (Flat OPC, ODT, semantic IR,
 * OMML ↔ MathML, document mapping).
 *
 * `import { Convert } from "@cj-tech-master/excelts/word"` →
 *   `Convert.parseFlatOpc(xml)`, `Convert.readOdt(buf)`,
 *   `Convert.docxToSemantic(doc)`, `Convert.ommlToMathML(...)`, … — tree-shaken
 *   via `export * as Convert`.
 */
export { parseFlatOpc, isFlatOpc, toFlatOpc } from "@word/convert/flat-opc";
export { readOdt, writeOdt } from "@word/convert/odt/odt";
export { createConversionContext } from "@word/convert/conversion-ir";
export { docxToSemantic } from "@word/convert/docx-to-semantic";
export { ommlToMathML, mathMLToOmml } from "@word/advanced/math-convert";
export { mapDocument } from "@word/core/mapper";
