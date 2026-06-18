/**
 * `Address` namespace surface — stateless A1-notation encode/decode helpers.
 *
 * `import { Address } from "documonster/excel"` →
 *   `Address.decodeCell("B3")`, `Address.encodeCol(2)`,
 *   `Address.decodeRange("A1:C5")`, `Address.quoteSheetName("My Sheet")`.
 */
export {
  decodeCol,
  encodeCol,
  decodeRow,
  encodeRow,
  decodeCell,
  encodeCell,
  decodeRange,
  encodeRange,
  quoteSheetName
} from "@excel/utils/address";
