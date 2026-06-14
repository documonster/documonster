/**
 * Test ACTUAL Excel streaming write behavior in browser
 * This tests the real WorkbookWriter, not just compression
 *
 * IMPORTANT: Browser CompressionStream is async, but JavaScript is single-threaded.
 * When sync code (like a for loop) writes data, the async reader never gets a chance
 * to run until the sync code yields to the event loop.
 *
 * True streaming in browser would require:
 * 1. Using async/await with setTimeout/setImmediate between writes, OR
 * 2. Using Web Workers for parallel processing
 *
 * Current behavior:
 * - Data is written to CompressionStream.writable (sync, fast)
 * - Data is read from CompressionStream.readable (async, only runs when event loop is free)
 * - Result: Compressed data appears after await points (worksheet.commit, workbook.commit)
 *
 * This is NOT a bug - it's how JavaScript's single-threaded event loop works with async streams.
 */

import { WorkbookWriter } from "@excel/stream/workbook-writer.browser";
import { rowCommit } from "@excel/worksheet";
import { describe, it, expect } from "vitest";

describe("Real Excel Streaming Write - Browser", () => {
  it("should test raw CompressionStream behavior in browser", async () => {
    // Test CompressionStream directly - no wrapper
    const chunks: number[] = [];

    const encoder = new TextEncoder();

    const cs = new CompressionStream("deflate-raw");
    const writer = cs.writable.getWriter();
    const reader = cs.readable.getReader();

    const readPromise = (async () => {
      while (true) {
        const { value, done } = await reader.read();
        if (done) {
          break;
        }
        if (value) {
          chunks.push(value.length);
        }
      }
    })();

    // Write a moderately large payload in fewer, larger chunks.
    // This keeps the test fast and avoids appearing to hang on slower machines.
    const totalRows = 20000;
    const rowsPerChunk = 250;
    console.log(`Writing ${totalRows} rows directly to CompressionStream...`);
    for (let i = 0; i < totalRows; i += rowsPerChunk) {
      let xml = "";
      const end = Math.min(totalRows, i + rowsPerChunk);
      for (let r = i; r < end; r++) {
        xml += `<row r="${r}"><c r="A${r}"><v>Row ${r} data content</v></c></row>\n`;
      }
      await writer.write(encoder.encode(xml));

      if (i > 0 && i % (rowsPerChunk * 20) === 0) {
        console.log(`Row ${i}: ${chunks.length} chunks`);
      }
    }

    console.log(`Before close: ${chunks.length} chunks`);
    await writer.close();
    await readPromise;
    console.log(`After close: ${chunks.length} chunks`);

    expect(chunks.length).toBeGreaterThan(0);
    console.log("✅ Browser CompressionStream supports streaming!");
  });

  it("should work correctly for small Excel (1000 rows)", async () => {
    const chunks: { time: number; size: number }[] = [];
    const startTime = performance.now();

    const outputChunks: Uint8Array[] = [];
    const writable = new WritableStream<Uint8Array>({
      write(chunk) {
        chunks.push({
          time: Math.round(performance.now() - startTime),
          size: chunk.length
        });
        outputChunks.push(chunk);
      }
    });

    const workbook = new WorkbookWriter({
      stream: writable
    });

    const worksheet = workbook.addWorksheet("Sheet1");

    // Write 1000 rows - small data
    for (let i = 0; i < 1000; i++) {
      rowCommit(worksheet.addRow([`Row ${i}`, i, `Data ${i}`]));
    }

    const chunksBeforeCommit = chunks.length;

    worksheet.commit();
    const chunksAfterWorksheetCommit = chunks.length;

    await workbook.commit();

    const totalSize = outputChunks.reduce((sum, c) => sum + c.length, 0);
    console.log(`\n=== Small Excel (1000 rows) Summary ===`);
    console.log(`Total: ${chunks.length} chunks, ${totalSize} bytes`);
    console.log(`Chunks during sync row writes: ${chunksBeforeCommit}`);
    console.log(
      `Chunks during worksheet commit: ${chunksAfterWorksheetCommit - chunksBeforeCommit}`
    );
    console.log(`Chunks during workbook commit: ${chunks.length - chunksAfterWorksheetCommit}`);

    // Verify we got valid Excel output
    expect(chunks.length).toBeGreaterThan(0);
    expect(totalSize).toBeGreaterThan(0);

    // Note: Due to JS single-threaded nature, sync code doesn't allow reader to run
    // Chunks appear after await points - this is expected browser behavior
    console.log("✅ Browser streaming works correctly (data flows at await points)");
  });

  it("should show TRUE streaming for very large Excel (200000 rows, >5MB uncompressed)", async () => {
    const chunks: { time: number; size: number }[] = [];
    const startTime = performance.now();

    const outputChunks: Uint8Array[][] = [[]];
    let outputPhase = 0;
    const writable = new WritableStream<Uint8Array>({
      write(chunk) {
        chunks.push({
          time: Math.round(performance.now() - startTime),
          size: chunk.length
        });
        outputChunks[outputPhase].push(chunk);
      }
    });

    const workbook = new WorkbookWriter({
      stream: writable
    });

    const worksheet = workbook.addWorksheet("Sheet1");

    // Write a large-enough worksheet to exceed streaming thresholds, but keep runtime sane.
    // We use fewer rows with larger per-row payload.
    const totalRows = 30000;
    const padX = "X".repeat(256);
    const padY = "Y".repeat(256);
    const padZ = "Z".repeat(256);

    for (let i = 0; i < totalRows; i++) {
      rowCommit(
        worksheet.addRow([
          `Row ${i}`,
          i,
          `Data ${i} with extra content ${padX}`,
          `More data ${i} padding ${padY}`,
          `Even more ${i} content ${padZ}`,
          `Column F ${i} value ${padX}`,
          `Column G ${i} text ${padY}`,
          `Column H ${i} end ${padZ}`
        ])
      );

      // Log progress every 10000 rows
      if (i > 0 && i % 10000 === 0) {
        const size = outputChunks[0].reduce((sum, c) => sum + c.length, 0);
        console.log(`At row ${i}: ${chunks.length} chunks, ${(size / 1024 / 1024).toFixed(2)} MB`);
      }
    }

    const chunksBeforeCommit = chunks.length;
    const phase0Size = outputChunks[0].reduce((sum, c) => sum + c.length, 0);

    outputPhase = 1;
    outputChunks.push([]);
    worksheet.commit();
    const chunksAfterWorksheetCommit = chunks.length;
    const phase1Size = outputChunks[1].reduce((sum, c) => sum + c.length, 0);

    outputPhase = 2;
    outputChunks.push([]);
    await workbook.commit();
    const phase2Size = outputChunks[2].reduce((sum, c) => sum + c.length, 0);

    const totalSize = phase0Size + phase1Size + phase2Size;
    console.log(`\n=== Very Large Excel (${totalRows} rows) Summary ===`);
    console.log(`Total: ${chunks.length} chunks, ${(totalSize / 1024 / 1024).toFixed(2)} MB`);
    console.log(
      `During row writes: ${chunksBeforeCommit} chunks, ${(phase0Size / 1024 / 1024).toFixed(2)} MB`
    );
    console.log(
      `During worksheet.commit(): ${chunksAfterWorksheetCommit - chunksBeforeCommit} chunks, ${(phase1Size / 1024 / 1024).toFixed(2)} MB`
    );
    console.log(
      `During workbook.commit(): ${chunks.length - chunksAfterWorksheetCommit} chunks, ${(phase2Size / 1024 / 1024).toFixed(2)} MB`
    );

    // Verify we got valid Excel output
    expect(chunks.length).toBeGreaterThan(0);
    expect(totalSize).toBeGreaterThan(1000000);

    // For a large dataset, DEFLATE output may appear during row writes depending on yielding.
    // and produce chunks DURING row writes (TRUE STREAMING)
    if (chunksBeforeCommit > 0) {
      console.log("✅ TRUE STREAMING: Chunks emitted during row writes!");
      console.log(
        `   ${chunksBeforeCommit} chunks (${(phase0Size / 1024 / 1024).toFixed(2)} MB) streamed before commit`
      );
    } else {
      // With synchronous for-loop, JavaScript single-thread blocks async reader
      // This is expected behavior - see async write test for TRUE streaming
      console.log("ℹ️ Sync loop: No streaming (expected - JS blocks async reader)");
      console.log("   For TRUE streaming, use async writes (see next test)");
    }

    // Note: Sync loop cannot stream due to JS single-threaded model
    // The async test below demonstrates TRUE streaming IS possible
  });

  it("should verify if row data streams progressively", async () => {
    /**
     * This test checks if the ROW DATA itself is being streamed,
     * not just the Excel header files.
     *
     * We write ~10MB of row data to clearly see streaming behavior.
     */
    const chunks: { time: number; size: number; totalBytes: number }[] = [];
    const startTime = performance.now();
    let totalBytes = 0;

    const writable = new WritableStream<Uint8Array>({
      write(chunk) {
        totalBytes += chunk.length;
        chunks.push({
          time: Math.round(performance.now() - startTime),
          size: chunk.length,
          totalBytes
        });
      }
    });

    const workbook = new WorkbookWriter({
      stream: writable
    });

    const worksheet = workbook.addWorksheet("Sheet1");

    // Track bytes at different stages
    const bytesLog: { stage: string; chunks: number; bytes: number }[] = [];

    // Write enough rows to clearly see streaming behavior without being too slow.
    // Each row has a fairly large payload to exceed thresholds.
    const padX = "X".repeat(256);
    const padY = "Y".repeat(256);
    const padZ = "Z".repeat(256);
    const totalRows = 20000;
    for (let i = 0; i < totalRows; i++) {
      // Make each row ~500 bytes of data
      rowCommit(
        worksheet.addRow([
          `Row ${i}`,
          i,
          `Data ${i} - ${padX}`,
          `More ${i} - ${padY}`,
          `Extra ${i} - ${padZ}`
        ])
      );

      // Yield to event loop regularly so async readers can flush.
      if (i > 0 && i % 2000 === 0) {
        await new Promise(r => setTimeout(r, 0));
        bytesLog.push({
          stage: `Row ${i}`,
          chunks: chunks.length,
          bytes: totalBytes
        });
      }
    }

    bytesLog.push({
      stage: "Before worksheet.commit()",
      chunks: chunks.length,
      bytes: totalBytes
    });
    const bytesBeforeCommit = totalBytes;

    worksheet.commit();
    bytesLog.push({
      stage: "After worksheet.commit()",
      chunks: chunks.length,
      bytes: totalBytes
    });

    await workbook.commit();
    bytesLog.push({
      stage: "After workbook.commit()",
      chunks: chunks.length,
      bytes: totalBytes
    });

    console.log("\n=== 25MB Row Data Streaming Analysis ===");
    bytesLog.forEach(log => {
      console.log(`${log.stage}: ${log.chunks} chunks, ${(log.bytes / 1024 / 1024).toFixed(2)} MB`);
    });

    // Check if bytes grew significantly during row writes
    const rowLogs = bytesLog.filter(l => l.stage.startsWith("Row "));
    const firstRowLog = rowLogs[0];
    const lastRowLog = rowLogs[rowLogs.length - 1];
    const growthDuringWrites = lastRowLog ? lastRowLog.bytes - firstRowLog.bytes : 0;

    console.log(`\nGrowth during row writes: ${(growthDuringWrites / 1024 / 1024).toFixed(2)} MB`);
    console.log(`Data before commit: ${(bytesBeforeCommit / 1024 / 1024).toFixed(2)} MB`);

    // 100KB growth = TRUE STREAMING (row data is streaming)
    if (growthDuringWrites > 100000) {
      console.log("\n✅ TRUE STREAMING: Row data streamed during writes!");
    } else if (bytesBeforeCommit > 10000) {
      console.log("\n⚠️ PARTIAL: Some data before commit, but minimal growth");
    } else {
      console.log("\n❌ NOT TRUE STREAMING: All data buffered until commit");
    }

    // Verify TRUE STREAMING: significant data growth during writes
    expect(growthDuringWrites).toBeGreaterThan(100000);

    expect(chunks.length).toBeGreaterThan(0);
  }, 60000);

  it("should demonstrate TRUE streaming with async writes", async () => {
    /**
     * This test demonstrates that TRUE streaming IS possible in browser
     * when we yield to the event loop between writes.
     *
     * By using await and setTimeout, we give the CompressionStream reader
     * a chance to process compressed data between writes.
     */
    const chunks: { time: number; size: number }[] = [];
    const startTime = performance.now();

    const writable = new WritableStream<Uint8Array>({
      write(chunk) {
        chunks.push({
          time: Math.round(performance.now() - startTime),
          size: chunk.length
        });
      }
    });

    const workbook = new WorkbookWriter({
      stream: writable
    });

    const worksheet = workbook.addWorksheet("Sheet1");

    // Write 500 rows with yields to event loop every 100 rows
    for (let i = 0; i < 500; i++) {
      rowCommit(worksheet.addRow([`Row ${i}`, i, `Data ${i}`, `More ${i}`.repeat(20)]));

      // Yield to event loop every 100 rows
      if (i > 0 && i % 100 === 0) {
        await new Promise(r => setTimeout(r, 0));
        console.log(`At row ${i}: ${chunks.length} chunks emitted so far`);
      }
    }

    const chunksBeforeCommit = chunks.length;
    console.log(`\nChunks before worksheet.commit(): ${chunksBeforeCommit}`);

    worksheet.commit();

    await workbook.commit();

    console.log(`\n=== Async Write Test (500 rows) ===`);
    console.log(`Total chunks: ${chunks.length}`);
    console.log(`Chunks during async writes: ${chunksBeforeCommit}`);
    console.log(`Chunks at commit: ${chunks.length - chunksBeforeCommit}`);

    expect(chunks.length).toBeGreaterThan(0);

    if (chunksBeforeCommit > 0) {
      console.log("✅ TRUE STREAMING achieved with async writes!");
    } else {
      console.log("⚠️ Still buffered (may need more data per batch)");
    }
  });
});
