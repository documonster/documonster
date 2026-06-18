/**
 * `Sparkline` namespace surface.
 * `import { Sparkline } from "@cj-tech-master/excelts/excel"` → `Sparkline.add(ws, opts)`.
 */
export {
  addSparklineGroup as add,
  getSparklineGroups as list,
  removeSparklineGroup as remove
} from "@excel/worksheet";

export type { AddSparklineGroupOptions, SparklineGroup } from "@excel/sparkline";
