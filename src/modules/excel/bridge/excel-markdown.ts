/**
 * `ExcelMarkdown` namespace surface — Markdown tables ↔ Workbook (cross-platform).
 *
 * `import { ExcelMarkdown } from "documonster/excel/markdown"` →
 *   `ExcelMarkdown.read(text)` (a new workbook: the first table, or one sheet per table with
 *   `{ tables: "all" }`; text or UTF-8 bytes), `ExcelMarkdown.addWorksheet(wb, text)` /
 *   `ExcelMarkdown.addWorksheets(wb, text)` (onto `wb`),
 *   `ExcelMarkdown.format(wb)`, `ExcelMarkdown.toBuffer(wb)` / `ExcelMarkdown.toBufferSync(wb)`.
 *
 * Named `ExcelMarkdown` rather than `Markdown` so it can be imported next to
 * `Markdown` from `documonster/markdown` without aliasing.
 */
export {
  readMarkdownWorkbook as read,
  addMarkdownWorksheet as addWorksheet,
  addMarkdownWorksheets as addWorksheets,
  writeMarkdown as format,
  markdownToBuffer as toBuffer,
  writeMarkdownBuffer as toBufferSync
} from "@excel/bridge/markdown-bridge";
