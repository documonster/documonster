/**
 * `RenderContext` namespace surface — render context / id generator factories.
 *
 * `import { RenderContext } from "@cj-tech-master/excelts/word"` →
 *   `RenderContext.create(opts)`, `RenderContext.createIds(...)` — tree-shaken
 *   via `export * as RenderContext`.
 */
export {
  createRenderContext as create,
  createIdGenerators as createIds
} from "@word/writer/render-context";
