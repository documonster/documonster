/**
 * Sparse `idx`-addressed points, shared between the classic and ChartEx readers.
 *
 * OOXML addresses cached points by `idx`, and `ptCount` is optional. Sizing the dense
 * array without consulting the largest `idx` therefore drops every point that sits past
 * the declared count — the ChartEx reader had its own copy of this logic and did exactly
 * that, rendering an empty chart instead of a sparse one. Both readers now share
 * `densifySparsePoints`, so the two cannot disagree again.
 */
import type { ChartExModel } from "@excel/chart";
import { ChartEx } from "@excel/chart";
import { densifySparsePoints } from "@excel/chart/shared/chart-utils";
import { describe, expect, it } from "vitest";

/** A one-series ChartEx model whose only value sits at `idx`, with no `ptCount`. */
function sparseModel(index: number, value: number): ChartExModel {
  return {
    chartSpace: {
      chart: {
        plotArea: {
          plotAreaRegion: {
            plotSurface: {},
            series: [
              { idx: 0, layoutId: "clusteredColumn", dataLabels: {}, dataRefs: [{ dataId: 0 }] }
            ]
          },
          axes: []
        }
      },
      chartData: {
        data: [{ id: 0, numDim: { type: "val", levels: [{ points: [{ index, value }] }] } }]
      }
    }
  } as unknown as ChartExModel;
}

/** Data rects, i.e. everything but the chart area (the one with Office's automatic border). */
function dataRects(svg: string): string[] {
  return [...svg.matchAll(/<rect [^>]*>/g)]
    .map(match => match[0])
    .filter(rect => !/stroke="#d9d9d9"/i.test(rect));
}

describe("densifySparsePoints", () => {
  it("sizes the array from the largest index when ptCount is absent", () => {
    // `points.length` is 1 and `ptCount` is undefined; only `maxIdx + 1` reaches slot 3.
    expect(
      densifySparsePoints([{ index: 3, value: 42 }], undefined, NaN, raw => raw ?? NaN)
    ).toEqual([NaN, NaN, NaN, 42]);
  });

  it("still honours a ptCount larger than the largest index", () => {
    expect(densifySparsePoints([{ index: 0, value: 7 }], 3, NaN, raw => raw ?? NaN)).toEqual([
      7,
      NaN,
      NaN
    ]);
  });
});

describe("a ChartEx series whose points are sparse", () => {
  it("keeps a value addressed past the declared count", () => {
    // Before sharing the densifier this rendered the background and nothing else.
    expect(dataRects(ChartEx.toSvg(sparseModel(3, 42), { width: 300, height: 200 }))).toHaveLength(
      1
    );
  });

  it("still renders a dense series", () => {
    expect(dataRects(ChartEx.toSvg(sparseModel(0, 42), { width: 300, height: 200 }))).toHaveLength(
      1
    );
  });
});
