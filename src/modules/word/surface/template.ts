/**
 * `Template` namespace surface — template filling, data sources, chart binding.
 *
 * `import { Template } from "documonster/word"` →
 *   `Template.fillTemplate(...)`, `Template.JsonDataSource`,
 *   `Template.fillTemplateFromSource(...)`, … — tree-shaken via
 *   `export * as Template`.
 */
export {
  fillTemplate,
  fillTemplateEnhanced,
  listTemplateTags,
  isTemplateChart
} from "@word/template/template-engine";
export { TemplateError } from "@word/errors";
export {
  JsonDataSource,
  XmlDataSource,
  CsvDataSource,
  CompositeDataSource,
  fillTemplateFromSource
} from "@word/template/template-datasource";
export { bindChartData } from "@word/template/template-chart";
