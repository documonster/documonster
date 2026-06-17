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
  isTemplateChart,
  TemplateError
} from "../template/template-engine";
export {
  JsonDataSource,
  XmlDataSource,
  CsvDataSource,
  CompositeDataSource,
  fillTemplateFromSource
} from "../template/template-datasource";
export { bindChartData } from "../template/template-chart";
