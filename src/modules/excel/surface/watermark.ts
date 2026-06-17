/**
 * `Watermark` namespace surface — worksheet background watermark.
 *
 * `import { Watermark } from "documonster/excel"` →
 *   `Watermark.add(ws, opts)`, `Watermark.get(ws)`, `Watermark.remove(ws)`.
 */
export {
  addWatermark as add,
  getWatermark as get,
  removeWatermark as remove
} from "@excel/worksheet";
