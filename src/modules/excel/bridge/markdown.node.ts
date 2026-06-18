/**
 * `documonster/excel/markdown` — Node entry.
 *
 * Re-exports the cross-platform Markdown functions plus the Node-only
 * file-path variants.
 */

export {
  readMarkdown,
  readMarkdownAll,
  writeMarkdown,
  writeMarkdownBuffer
} from "./markdown-bridge";
export { readMarkdownFile, readMarkdownAllFile, writeMarkdownFile } from "./markdown-bridge.node";
