/**
 * `StyleMap` namespace surface — style-mapping DSL parsing / matching.
 *
 * `import { StyleMap } from "documonster/word"` →
 *   `StyleMap.parse(dsl)`, `StyleMap.create(rules)`, `StyleMap.match(...)`,
 *   `StyleMap.DEFAULT`, … — tree-shaken via `export * as StyleMap`.
 */
export {
  parseStyleMap as parse,
  createStyleMap as create,
  mergeStyleMaps as merge,
  matchStyleMap as match,
  DEFAULT_STYLE_MAP as DEFAULT
} from "@word/advanced/style-map";
