import {
  mathMLToOmml as mathMLToOmmlImpl,
  ommlToMathML as ommlToMathMLImpl
} from "@word/advanced/math-convert";
/**
 * `Convert` namespace surface — format conversions (Flat OPC, ODT, semantic IR,
 * OMML ↔ MathML).
 *
 * `import { Convert } from "documonster/word"` →
 *   `Convert.flatOpcToParts(xml)`, `Convert.partsToFlatOpc(parts, types)`, `Convert.fromOdt(buf)`, `Convert.toOdt(doc)`,
 *   `Convert.toSemantic(doc)`, `Convert.ommlToMathMl(...)`, … — tree-shaken
 *   via `export * as Convert`.
 */
import { docxToSemantic as docxToSemanticImpl } from "@word/convert/docx-to-semantic";
import {
  parseFlatOpc as parseFlatOpcImpl,
  toFlatOpc as toFlatOpcImpl
} from "@word/convert/flat-opc";
import { readOdt as readOdtImpl, writeOdt as writeOdtImpl } from "@word/convert/odt/odt";
import { mapDocument as mapDocumentImpl } from "@word/core/mapper";

export { isFlatOpc, toFlatOpc as partsToFlatOpc } from "@word/convert/flat-opc";
export { createConversionContext } from "@word/convert/conversion-ir";

export { parseFlatOpc as flatOpcToParts } from "@word/convert/flat-opc";
export { readOdt as fromOdt, writeOdt as toOdt } from "@word/convert/odt/odt";
export { docxToSemantic as toSemantic } from "@word/convert/docx-to-semantic";
export {
  ommlToMathML as ommlToMathMl,
  mathMLToOmml as mathMlToOmml
} from "@word/advanced/math-convert";

/**
 * @deprecated Use {@link flatOpcToParts} for the part map, or `Io.read` for the
 * document a Flat OPC file holds. Removed in 1.0.
 */
export const parseFlatOpc = parseFlatOpcImpl;
/**
 * @deprecated Use {@link partsToFlatOpc} — this wraps a part map, not a
 * document; `Io.toFlatOpc(doc)` is the document form. Removed in 1.0.
 */
export const toFlatOpc = toFlatOpcImpl;
/** @deprecated Use {@link fromOdt}. Removed in 1.0. */
export const readOdt = readOdtImpl;
/** @deprecated Use {@link toOdt}. Removed in 1.0. */
export const writeOdt = writeOdtImpl;
/** @deprecated Use {@link toSemantic}. Removed in 1.0. */
export const docxToSemantic = docxToSemanticImpl;
/** @deprecated Use {@link ommlToMathMl}. Removed in 1.0. */
export const ommlToMathML = ommlToMathMLImpl;
/** @deprecated Use {@link mathMlToOmml}. Removed in 1.0. */
export const mathMLToOmml = mathMLToOmmlImpl;
/** @deprecated Use `Transform.map`. Removed in 1.0. */
export const mapDocument = mapDocumentImpl;
