/**
 * `Theme` namespace surface — theme color resolution.
 *
 * `import { Theme } from "documonster/word"` →
 *   `Theme.resolveColor(...)` — tree-shaken via `export * as Theme`.
 */
export { resolveThemeColor as resolveColor } from "@word/core/color-utils";
