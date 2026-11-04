/**
 * `Template` namespace surface — template filling (variables, conditionals,
 * loops), data sources, chart binding.
 *
 * `Template.fill`, `Template.fillFromSource` and `Template.bindChartData`
 * leave the document they are given unchanged and return a new one, which may
 * share unchanged nodes with it — take `Document.clone(result)` before editing
 * the result in place while you still use the template. (The deprecated
 * `fillTemplate` and `fillTemplateFromSource` fill `doc` in place instead.)
 *
 * Placeholder patching — exact replacement of named placeholders with text,
 * paragraphs, tables or images, no data object — is a different engine and
 * lives in `Package` (`Package.patch`, `Package.patchTemplate`). Use
 * `Template.fill` when the template carries logic, `Package.patch` when you
 * only swap placeholders for ready-made content.
 *
 * `import { Template } from "documonster/word"` →
 *   `Template.fill(doc, data)`, `Template.listTags(doc)`,
 *   `Template.fillFromSource(doc, source)`, … — tree-shaken via
 *   `export * as Template`.
 *
 * Byte-level template operations (`.docx` in, `.docx` out) live in `Package`
 * (`Package.fillTemplate`, `Package.patch`, `Package.compileTemplate`,
 * `Package.patchTemplate`), so importing `Template` does not pull the ZIP
 * reader and writer.
 */
import { TemplateError as TemplateErrorClass } from "@word/errors";
import { fillTemplateFromSource as fillTemplateFromSourceImpl } from "@word/template/template-datasource";
import {
  fill,
  fillTemplate as fillTemplateImpl,
  fillTemplateEnhanced as fillTemplateEnhancedImpl,
  isTemplateChart as isTemplateChartImpl,
  listTemplateTags as listTemplateTagsImpl
} from "@word/template/template-engine";

export { fill };
export {
  createJsonDataSource,
  createXmlDataSource,
  createCsvDataSource,
  createCompositeDataSource
} from "@word/template/template-datasource";
export { bindChartData } from "@word/template/template-chart";

export { isTemplateChart as isChart } from "@word/template/template-engine";
export { listTags } from "@word/template/template-engine";
export { fillFromSource } from "@word/template/template-datasource";

/**
 * Fills variables, conditionals and loops in place (the input is modified and
 * returned); no rich placeholders, no comment bodies.
 *
 * @deprecated Use {@link fill}, which returns a new document. Removed in 1.0.
 */
export const fillTemplate = fillTemplateImpl;
/**
 * Returns a new document with rich placeholders resolved as well; no comment
 * bodies.
 *
 * @deprecated Use {@link fill} with `{ richPlaceholders: true }`. Removed in 1.0.
 */
export const fillTemplateEnhanced = fillTemplateEnhancedImpl;
/** @deprecated Use {@link listTags}. Removed in 1.0. */
export const listTemplateTags = listTemplateTagsImpl;
/** @deprecated Use {@link fillFromSource}. Removed in 1.0. */
export const fillTemplateFromSource = fillTemplateFromSourceImpl;
/** @deprecated Use {@link isChart}. Removed in 1.0. */
export const isTemplateChart = isTemplateChartImpl;
/**
 * @deprecated Import `TemplateError` from `documonster/word` directly. Removed in 1.0.
 */
export const TemplateError = TemplateErrorClass;
/**
 * @deprecated Import `TemplateError` from `documonster/word` directly. Removed in 1.0.
 */
export type TemplateError = TemplateErrorClass;
