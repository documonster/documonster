/**
 * Chartsheet structure check.
 *
 * A chartsheet (`xl/chartsheets/sheetN.xml`) must contain a `<drawing>`
 * reference — the chart itself is always in the linked drawing part.
 * Without it, Excel shows an empty chart sheet.
 */

import type { ValidationContext } from "./context";
import { hasDescendantLocal } from "./xml-utils";

const CHARTSHEET_PATH_RE = /^xl\/chartsheets\/sheet\d+\.xml$/;

export function checkChartsheet(ctx: ValidationContext): void {
  for (const [path, entry] of ctx.files()) {
    if (ctx.reporter.capped) {
      return;
    }
    if (entry.type === "directory" || !CHARTSHEET_PATH_RE.test(path)) {
      continue;
    }
    const dom = ctx.readDom(path);
    if (!dom) {
      continue;
    }
    if (!hasDescendantLocal(dom.root, "drawing")) {
      ctx.reporter.error("chartsheet-missing-drawing", `${path}: missing drawing reference`, path);
    }
  }
}
