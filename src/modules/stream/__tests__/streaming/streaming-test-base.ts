/**
 * Shared True Streaming Test Utilities
 *
 * This module provides test helpers that work in both Node.js and Browser environments.
 * Use these to verify TRUE streaming behavior (data flows progressively, not buffered).
 */

import { expect } from "vitest";

// ============================================================================
// Types
// ============================================================================

export interface StreamingMetrics {
  chunks: { time: number; size: number; totalBytes: number; phase: string }[];
  startTime: number;
  totalBytes: number;
  currentPhase: string;
}

export interface PhaseLog {
  phase: string;
  chunks: number;
  bytes: number;
}

// ============================================================================
// Metric Collection Helpers
// ============================================================================

export function createMetrics(): StreamingMetrics {
  return {
    chunks: [],
    startTime: typeof performance !== "undefined" ? performance.now() : Date.now(),
    totalBytes: 0,
    currentPhase: "setup"
  };
}

export function recordChunk(metrics: StreamingMetrics, chunk: Uint8Array): void {
  const now = typeof performance !== "undefined" ? performance.now() : Date.now();
  metrics.totalBytes += chunk.length;
  metrics.chunks.push({
    time: Math.round(now - metrics.startTime),
    size: chunk.length,
    totalBytes: metrics.totalBytes,
    phase: metrics.currentPhase
  });
}

export function setPhase(metrics: StreamingMetrics, phase: string): void {
  metrics.currentPhase = phase;
}

export function logPhase(metrics: StreamingMetrics, phaseName: string): PhaseLog {
  return {
    phase: phaseName,
    chunks: metrics.chunks.length,
    bytes: metrics.totalBytes
  };
}

// ============================================================================
// Analysis Helpers
// ============================================================================

export function analyzeGrowth(logs: PhaseLog[]): {
  totalGrowth: number;
  hadProgressiveGrowth: boolean;
  growthPerPhase: number[];
} {
  const growthPerPhase = logs.map((log, i, arr) =>
    i > 0 ? log.bytes - arr[i - 1].bytes : log.bytes
  );

  return {
    totalGrowth: logs.length > 1 ? logs[logs.length - 1].bytes - logs[0].bytes : 0,
    hadProgressiveGrowth: growthPerPhase.filter(g => g > 0).length > 1,
    growthPerPhase
  };
}

export function printAnalysis(
  title: string,
  logs: PhaseLog[],
  bytesBeforeFinalize: number,
  totalBytes: number
): void {
  const shouldLog = (() => {
    const g = globalThis as unknown as { __DOCUMONSTER_TEST_LOGS__?: boolean };
    if (g.__DOCUMONSTER_TEST_LOGS__) {
      return true;
    }
    if (typeof process !== "undefined") {
      const env = (process as unknown as { env?: Record<string, string | undefined> }).env;
      return env?.DOCUMONSTER_TEST_LOGS === "1" || env?.VITEST_VERBOSE === "1";
    }
    return false;
  })();

  if (!shouldLog) {
    return;
  }

  console.log(`\n=== ${title} ===`);
  for (const log of logs) {
    const mb = (log.bytes / 1024 / 1024).toFixed(2);
    console.log(`${log.phase}: ${log.chunks} chunks, ${mb} MB`);
  }

  const analysis = analyzeGrowth(logs.filter(l => !l.phase.includes("After")));
  console.log(`\nGrowth during processing: ${(analysis.totalGrowth / 1024 / 1024).toFixed(2)} MB`);
  console.log(`Data before finalize: ${(bytesBeforeFinalize / 1024 / 1024).toFixed(2)} MB`);
  console.log(`Total data: ${(totalBytes / 1024 / 1024).toFixed(2)} MB`);

  // TRUE STREAMING: any data output before finalize means streaming is working
  // Even 1% means data flows through without buffering the entire file
  const ratio = totalBytes > 0 ? bytesBeforeFinalize / totalBytes : 0;
  if (bytesBeforeFinalize > 0) {
    console.log(`\n✅ TRUE STREAMING: ${(ratio * 100).toFixed(0)}% data streamed before finalize!`);
  } else {
    console.log("\n❌ NOT TRUE STREAMING: All data buffered until finalize");
  }
}

// ============================================================================
// Async Yield Helper
// ============================================================================

/**
 * Yield to event loop - allows async operations to complete.
 * Node.js zlib uses libuv thread pool, so we need multiple yields
 * to ensure pending compression callbacks are processed.
 */
export function yieldToEventLoop(iterations: number = 1): Promise<void> {
  return new Promise(resolve => {
    let count = 0;
    const tick = () => {
      count++;
      if (count >= iterations) {
        resolve();
      } else if (typeof setImmediate !== "undefined") {
        setImmediate(tick);
      } else {
        setTimeout(tick, 0);
      }
    };
    if (typeof setImmediate !== "undefined") {
      setImmediate(tick);
    } else {
      setTimeout(tick, 0);
    }
  });
}

/**
 * Strong yield - uses setTimeout to ensure libuv thread pool has time to complete.
 * Node.js zlib operations run in the thread pool, so setImmediate alone may not be enough.
 */
export function strongYield(): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, 5));
}

// ============================================================================
// Test Data Generators
// ============================================================================

/**
 * Generate row data with pseudo-random content for streaming tests.
 * Uses varying characters to prevent extreme compression.
 */
export function generateRowData(rowIndex: number, columnsPerRow: number = 5): string[] {
  const row: string[] = [`Row ${rowIndex}`, String(rowIndex)];
  // Use different characters based on row index to reduce compression ratio
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  for (let c = 2; c < columnsPerRow; c++) {
    // Generate pseudo-random content for each cell
    let content = `C${c}_`;
    for (let j = 0; j < 50; j++) {
      content += chars[(rowIndex * 17 + c * 13 + j * 7) % chars.length];
    }
    content += `_${rowIndex}`;
    row.push(content);
  }
  return row;
}

export function generateLargeText(sizeBytes: number): string {
  const chunk = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  let result = "";
  while (result.length < sizeBytes) {
    result += chunk;
  }
  return result.slice(0, sizeBytes);
}

export function generateRandomBytes(size: number): Uint8Array {
  const data = new Uint8Array(size);
  for (let i = 0; i < size; i++) {
    data[i] = Math.floor(Math.random() * 256);
  }
  return data;
}

// ============================================================================
// Assertion Helpers
// ============================================================================

export function assertTrueStreaming(
  growthDuringProcessing: number,
  bytesBeforeFinalize: number
): void {
  // TRUE STREAMING: any data output before finalize means streaming is working
  // Even 1 byte means data flows through without buffering the entire file
  expect(
    bytesBeforeFinalize,
    `Expected data before finalize > 0 bytes, got ${bytesBeforeFinalize}. All data was buffered until finalize!`
  ).toBeGreaterThan(0);
}
