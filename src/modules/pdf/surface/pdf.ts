/**
 * `Pdf` namespace surface — PDF writing, reading, building, editing,
 * conversion, and digital signatures.
 *
 * `import { Pdf } from "documonster/pdf"` →
 *   `Pdf.create(rows)`, `Pdf.fromExcel(wb)`, `Pdf.fromDocx(doc)`,
 *   `Pdf.read(bytes)`, `new Pdf.Builder()`, `new Pdf.Editor()`,
 *   `Pdf.sign(...)`, `Pdf.verifySignature(...)`.
 *
 * Single flat namespace (pdf is one cohesive engine). Cross-module
 * converters live here as `from<Source>` verbs (the pdf module is the
 * upper layer that depends on excel/word). Re-exported via `export * as Pdf`.
 */

// Writing
export { pdf as create } from "../pdf";

// Conversion bridges (cross-module → target namespace, per API design)
export {
  excelToPdf as fromExcel,
  chartToPdf as fromChart,
  createWordChartPdfRenderer
} from "../excel-bridge";
export { docxToPdf as fromDocx } from "../word-bridge";

// Reading
export { readPdf as read } from "../reader/pdf-reader";

// Building
export {
  PdfDocumentBuilder as Builder,
  PdfPageBuilder as PageBuilder,
  parseSvgPath
} from "../builder/document-builder";

// Editing
export { PdfEditor as Editor, PdfEditorPage as EditorPage } from "../builder/pdf-editor";

// Digital signatures
export {
  verifyPdfSignature as verifySignature,
  signPdf as sign,
  buildSignatureDictPlaceholder,
  asn1Parse
} from "../core/digital-signature";

// Page-size presets
export { PageSizes } from "../types";
