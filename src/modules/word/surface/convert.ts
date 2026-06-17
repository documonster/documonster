/**
 * `Convert` namespace surface — format conversions (Flat OPC, ODT, semantic IR,
 * OMML ↔ MathML, document mapping).
 *
 * `import { Convert } from "documonster/word"` →
 *   `Convert.parseFlatOpc(xml)`, `Convert.readOdt(buf)`,
 *   `Convert.docxToSemantic(doc)`, `Convert.ommlToMathML(...)`, … — tree-shaken
 *   via `export * as Convert`.
 */
export { parseFlatOpc, isFlatOpc, toFlatOpc } from "../convert/flat-opc";
export { readOdt, writeOdt } from "../convert/odt/odt";
export { createConversionContext } from "../convert/conversion-ir";
export { docxToSemantic } from "../convert/docx-to-semantic";
export { ommlToMathML, mathMLToOmml } from "../advanced/math-convert";
export { mapDocument } from "../core/mapper";
