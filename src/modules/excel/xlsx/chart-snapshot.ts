/**
 * Chart-model snapshots — the one piece of chart state both directions need: the reader records one when
 * it loads a chart, and the writer compares against it to tell an untouched chart from an edited one.
 */
export function snapshotChartModel(model: unknown): string | undefined {
  try {
    return JSON.stringify(model);
  } catch {
    return undefined;
  }
}
