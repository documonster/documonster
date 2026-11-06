/**
 * `ChartClassic.Backend` — the classic chart renderer's building blocks, for a caller writing their own backend:
 * `toScene` (the scene every backend draws from), `applyAxisTransform`, `toSvgFilter` and `resolveTopologyObject`.
 */
import type { ChartModel } from "@excel/chart/model/types";
import { buildChartScene } from "@excel/chart/render/chart-renderer";
import type { ChartRenderOptions, ChartScene } from "@excel/chart/render/chart-renderer";
import { resolveTopologyObject as resolveTopologyObjectImpl } from "@excel/chart/render/topojson";
import type { TopologyLike } from "@excel/chart/render/topojson";

export {
  applyAxisTransform,
  buildEffectFilter as toSvgFilter
} from "@excel/chart/render/chart-renderer";

/**
 * The scene a classic chart model draws as — the geometry every backend paints from.
 *
 * @throws {ExcelChartOptionsError} `model` is not a classic chart model, or an option has no valid
 *   reading.
 */
export function toScene(model: ChartModel, options?: ChartRenderOptions): ChartScene {
  return buildChartScene(model, options, "ChartClassic.Backend.toScene");
}
/**
 * The features of the object `objectName` of a TopoJSON topology, as `{ id, properties,
 * geometry, rings }` records in absolute coordinates — what a region map draws.
 *
 * @throws {ExcelChartOptionsError} the topology cannot be read — no such object, an arc table or a
 *   geometry of the wrong shape — with `field` naming the path (`topology.arcs[3]`).
 */
export function resolveTopologyObject(
  topology: TopologyLike,
  objectName: string
): ReturnType<typeof resolveTopologyObjectImpl> {
  return resolveTopologyObjectImpl(topology, objectName);
}
