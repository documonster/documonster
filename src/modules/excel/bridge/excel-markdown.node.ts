/**
 * `ExcelMarkdown` namespace surface — Node entry (adds the file-path variants).
 */
export {
  readMarkdownWorkbook as read,
  addMarkdownWorksheet as addWorksheet,
  addMarkdownWorksheets as addWorksheets,
  writeMarkdown as format,
  markdownToBuffer as toBuffer,
  writeMarkdownBuffer as toBufferSync
} from "@excel/bridge/markdown-bridge";
export {
  readMarkdownFileWorkbook as readFile,
  writeMarkdownFile as writeFile
} from "@excel/bridge/markdown-bridge.node";
