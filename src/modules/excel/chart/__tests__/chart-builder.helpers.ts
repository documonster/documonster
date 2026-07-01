/**
 * Shared helpers + imports for the chart-builder test suite.
 *
 * The chart-builder tests were originally a single 13,000+ line file.
 * Loading that file in a full-suite context pushed vitest's
 * transform/import pipeline past 40 seconds per run. Splitting the
 * describe blocks into multiple `chart-builder-*.test.ts` files let
 * vitest transform and type-check each half in parallel; this module
 * hosts the bits every split shares: imports, helper utilities, and
 * small fixture builders.
 *
 * The original file's first 255 lines lived here; keeping the pattern
 * (everything every split needs, nothing test-specific) makes the
 * surface extremely boring by design. If you find yourself wanting to
 * add a test-specific helper, add it in the calling file instead.
 */

import type {
  AddChartOptions,
  AddChartSeriesOptions,
  ChartModel,
  ChartTypeGroup,
  PlotArea
} from "@excel/chart";
import { addChart, getCharts } from "@excel/core/worksheet";
import { Cell, Chart, Workbook, Worksheet } from "@excel/index";
import { expect } from "vitest";

const textDecoder = new TextDecoder();

// ---------------------------------------------------------------------------
// PNG inspection utilities
//
// The renderer tests (P2 and the bridge/preview tests) verify PNG output
// by cracking open the 8-byte signature + chunk table. These helpers
// are stable pure functions so they live here; nothing about them
// depends on test state.
// ---------------------------------------------------------------------------

export function stableHash(input: string): string {
  let hash = 2166136261;
  for (let i = 0; i < input.length; i++) {
    hash ^= input.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

export function expectPngDimensions(png: Uint8Array, width: number, height: number): void {
  expect([...png.slice(0, 8)]).toEqual([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  expect(textDecoder.decode(png.slice(12, 16))).toBe("IHDR");
  expect(readU32be(png, 16)).toBe(width);
  expect(readU32be(png, 20)).toBe(height);
  expect(png.length).toBeGreaterThan(100);
}

export function expectPngPhysDpi(png: Uint8Array, dpi: number): void {
  const offset = findPngChunk(png, "pHYs");
  expect(offset).toBeGreaterThan(0);
  const pixelsPerMeter = readU32be(png, offset + 8);
  expect(pixelsPerMeter).toBe(Math.round(dpi / 0.0254));
  expect(png[offset + 16]).toBe(1);
}

export function pngSignature(png: Uint8Array): string {
  const idat = collectPngChunks(png, "IDAT");
  const phys = collectPngChunks(png, "pHYs");
  return stableHash(
    [
      readU32be(png, 16),
      readU32be(png, 20),
      idat.reduce((sum, chunk) => sum + chunk.length, 0),
      stableHashBytes(idat),
      stableHashBytes(phys)
    ].join(":")
  );
}

export function collectPngChunks(png: Uint8Array, type: string): Uint8Array[] {
  const chunks: Uint8Array[] = [];
  let offset = 8;
  while (offset + 12 <= png.length) {
    const length = readU32be(png, offset);
    const chunkType = textDecoder.decode(png.slice(offset + 4, offset + 8));
    if (chunkType === type) {
      chunks.push(png.slice(offset + 8, offset + 8 + length));
    }
    offset += 12 + length;
  }
  return chunks;
}

export function stableHashBytes(chunks: Uint8Array[]): string {
  let hash = 2166136261;
  for (const chunk of chunks) {
    for (const byte of chunk) {
      hash ^= byte;
      hash = Math.imul(hash, 16777619);
    }
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

export function findPngChunk(png: Uint8Array, type: string): number {
  let offset = 8;
  while (offset + 12 <= png.length) {
    const length = readU32be(png, offset);
    if (textDecoder.decode(png.slice(offset + 4, offset + 8)) === type) {
      return offset;
    }
    offset += 12 + length;
  }
  return -1;
}

export function readU32be(bytes: Uint8Array, offset: number): number {
  return (
    bytes[offset] * 0x1000000 +
    ((bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3])
  );
}

// ---------------------------------------------------------------------------
// Workbook builders used by several splits
// ---------------------------------------------------------------------------

/**
 * Build a tiny two-row workbook with a bar chart so renderer tests can
 * exercise `exportRenderedChartModel()` without re-wiring the whole
 * fixture each time.
 */
export function makeRootExportRenderedChartModel(): ChartModel {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "Sheet1");
  Worksheet.addRows(ws, [
    ["A", 10],
    ["B", 20]
  ]);
  addChart(
    ws,
    {
      type: "bar",
      series: [{ name: "S", categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }],
      title: "Sales"
    },
    "D1:J10"
  );
  return Chart.chartModel(getCharts(ws)[0])!;
}

// ---------------------------------------------------------------------------
// Series / model accessor helpers
//
// Every chart-builder test file reaches for these. Keeping them in one
// place means we never fork a subtly-different `baseSeries` between
// splits.
// ---------------------------------------------------------------------------

export const CATEGORIES = "Sheet1!$A$1:$A$4";
export const VALUES_A = "Sheet1!$B$1:$B$4";
export const VALUES_B = "Sheet1!$C$1:$C$4";

export function baseSeries(name: string, values = VALUES_A): AddChartSeriesOptions {
  return { name, categories: CATEGORIES, values };
}

export function scatterSeries(name: string): AddChartSeriesOptions {
  return { name, xValues: "Sheet1!$A$1:$A$4", values: VALUES_A };
}

export function bubbleSeries(name: string): AddChartSeriesOptions {
  return {
    name,
    xValues: "Sheet1!$A$1:$A$3",
    values: "Sheet1!$B$1:$B$3",
    bubbleSize: "Sheet1!$C$1:$C$3"
  };
}

export function pa(m: ChartModel): PlotArea {
  return m.chart.plotArea;
}

export function ctg(m: ChartModel, idx = 0): ChartTypeGroup {
  return pa(m).chartTypes[idx];
}

/** Round-trip: `addChart → write → load → return Chart`. */
export async function roundTripChart(opts: AddChartOptions) {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "Sheet1");
  Cell.setValue(ws, "A1", "x");
  addChart(ws, opts, "C1:J15");
  const buf = await Workbook.toBuffer(wb);
  const wb2 = Workbook.create();
  await Workbook.read(wb2, buf);
  const ws2 = Workbook.getWorksheet(wb2, "Sheet1")!;
  return getCharts(ws2)[0];
}
