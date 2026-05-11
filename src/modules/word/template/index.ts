/** @module Sub-path entry point for direct imports: `import { ... } from "excelts/word/template"` */

export {
  fillTemplate,
  fillTemplateEnhanced,
  listTemplateTags,
  isTemplateChart
} from "./template-engine";
export type { TemplateOptions, TemplateTag } from "./template-engine";
export { TemplateError } from "./template-engine";
export {
  fillTemplateFromSource,
  JsonDataSource,
  XmlDataSource,
  CsvDataSource,
  CompositeDataSource
} from "./template-datasource";
export type { DataSource, FillFromSourceOptions } from "./template-datasource";
export { bindChartData } from "./template-chart";
export type { ChartBinding, ChartSeriesData, ChartTemplateData } from "./template-chart";
