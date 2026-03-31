/**
 * Browser utility functions
 * Re-exports shared utilities and adds browser-specific implementations
 */

// Re-export all shared utilities
export {
  delay,
  dateToExcel,
  excelToDate,
  parseOoxmlDate,
  decodeOoxmlEscape,
  encodeOoxmlEscape,
  encodeOoxmlAttr,
  validInt,
  isDateFmt,
  splitFormatSections,
  parseBoolean,
  range,
  toSortedArray,
  bufferToString,
  base64ToUint8Array,
  uint8ArrayToBase64,
  stringToUtf16Le
} from "@utils/utils.base";

// Re-export file system utilities from centralized fs module (browser stubs)
export { fileExists } from "@utils/fs.browser";
