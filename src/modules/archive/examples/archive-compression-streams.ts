/**
 * Example: Archive Module — Compression Streams, Auto-Detect & TAR Guards
 *
 * Covers archive exports not demonstrated by archive-complete.ts or
 * archive-functional-tar-errors.ts:
 * - Incremental CRC32: crc32Update / crc32Finalize (chunked == one-shot crc32)
 * - Streaming gzip/zlib factories: createGzipStream / createGunzipStream /
 *   createZlibStream / createUnzlibStream (round-trip through Node streams)
 * - Synchronous zlib: zlibSync / unzlibSync
 * - Auto-detect decompression: decompressAuto / decompressAutoSync (feed it
 *   gzip- or zlib-framed data and it picks the right inflater)
 * - TAR entry type guards: Archive.isTarFile / Archive.isTarDirectory /
 *   Archive.isTarSymlink, driven by an Archive.openTar reader
 *
 * The streaming/sync/CRC/auto-detect helpers are Node-only and exported from
 * the package entrypoint (not the platform-agnostic Archive namespace); the TAR
 * guards live on the Archive namespace.
 *
 * Usage: pnpm example --filter archive-compression-streams
 */
import { Readable, pipeline } from "node:stream";
import { promisify } from "node:util";

import { Archive, Compression } from "../index";

const pipelineAsync = promisify(pipeline);
const encoder = new TextEncoder();
const decoder = new TextDecoder();

// =============================================================================
// 1. Incremental CRC32 — Archive.crc32Update / Archive.crc32Finalize
// =============================================================================
//
// Chunked CRC must equal the one-shot Archive.crc32() of the concatenated input.
// Start the running state at 0xffffffff, fold in each chunk, then finalize.

console.log("1. Incremental CRC32 (Archive.crc32Update / Archive.crc32Finalize):");

const fullData = encoder.encode("The quick brown fox jumps over the lazy dog");
const chunkA = fullData.subarray(0, 16);
const chunkB = fullData.subarray(16, 30);
const chunkC = fullData.subarray(30);

let running = 0xffffffff;
running = Archive.crc32Update(running, chunkA);
running = Archive.crc32Update(running, chunkB);
running = Archive.crc32Update(running, chunkC);
const incremental = Archive.crc32Finalize(running);

const oneShot = Archive.crc32(fullData);

console.log(`  incremental: 0x${incremental.toString(16)}`);
console.log(`  one-shot:    0x${oneShot.toString(16)}`);
console.log(`  match:       ${incremental === oneShot}`);

// =============================================================================
// 2. Streaming GZIP round-trip — Compression.createGzipStream / Compression.createGunzipStream
// =============================================================================

console.log("\n2. Streaming gzip (Compression.createGzipStream / Compression.createGunzipStream):");

const gzipPayload = Buffer.from("Compressible gzip stream line. ".repeat(100));

const gzipStream = Compression.createGzipStream({ level: 6 });
const gzipChunks: Buffer[] = [];
gzipStream.on("data", (c: Buffer) => gzipChunks.push(c));
gzipStream.end(gzipPayload);
await new Promise<void>(resolve => gzipStream.on("end", resolve));
const gzipped = Buffer.concat(gzipChunks);
console.log(`  ${gzipPayload.length} bytes → ${gzipped.length} bytes gzipped`);

// Decompress the gzip stream back to the original.
const gunzipStream = Compression.createGunzipStream();
const gunzipChunks: Buffer[] = [];
gunzipStream.on("data", (c: Buffer) => gunzipChunks.push(c));
gunzipStream.end(gzipped);
await new Promise<void>(resolve => gunzipStream.on("end", resolve));
const gunzipped = Buffer.concat(gunzipChunks);
console.log(`  round-trip identical: ${gunzipped.equals(gzipPayload)}`);

// =============================================================================
// 3. Streaming Zlib round-trip — Compression.createZlibStream / Compression.createUnzlibStream
// =============================================================================
//
// Drive these two through node:stream pipeline() to show they compose as real
// Transform streams.

console.log("\n3. Streaming zlib (Compression.createZlibStream / Compression.createUnzlibStream):");

const zlibPayload = Buffer.from("Zlib pipeline payload. ".repeat(80));

const zlibStream = Compression.createZlibStream({ level: 9 });
const zlibChunks: Buffer[] = [];
zlibStream.on("data", (c: Buffer) => zlibChunks.push(c));
zlibStream.end(zlibPayload);
await new Promise<void>(resolve => zlibStream.on("end", resolve));
const zlibbed = Buffer.concat(zlibChunks);
console.log(`  ${zlibPayload.length} bytes → ${zlibbed.length} bytes zlib-deflated`);

const unzlibStream = Compression.createUnzlibStream();
const unzlibChunks: Buffer[] = [];
unzlibStream.on("data", (c: Buffer) => unzlibChunks.push(c));
// Feed a single-chunk readable through pipeline() to show real stream composition.
await pipelineAsync(Readable.from([zlibbed]), unzlibStream);
const unzlibbed = Buffer.concat(unzlibChunks);
console.log(`  round-trip identical: ${unzlibbed.equals(zlibPayload)}`);

// =============================================================================
// 4. Synchronous zlib — Compression.zlibSync / Compression.unzlibSync
// =============================================================================

console.log("\n4. Synchronous zlib (Compression.zlibSync / Compression.unzlibSync):");

const syncPayload = encoder.encode("Synchronous zlib round-trip. ".repeat(50));
const syncZlibbed = Compression.zlibSync(syncPayload, { level: 6 });
const syncRestored = Compression.unzlibSync(syncZlibbed);
console.log(
  `  ${syncPayload.length} bytes → ${syncZlibbed.length} bytes → ${syncRestored.length} bytes`
);
console.log(
  `  round-trip identical: ${decoder.decode(syncRestored) === decoder.decode(syncPayload)}`
);

// =============================================================================
// 5. Auto-detect decompression — Compression.decompressAuto / Compression.decompressAutoSync
// =============================================================================
//
// Both helpers sniff the format (gzip magic 0x1f8b, zlib CMF/FLG, else raw
// deflate) and route to the matching inflater. Feed each a different framing.

console.log(
  "\n5. Auto-detect decompression (Compression.decompressAuto / Compression.decompressAutoSync):"
);

const autoText = "Auto-detected payload. ".repeat(40);
const asGzip = Compression.gzipSync(encoder.encode(autoText));
const asZlib = Compression.zlibSync(encoder.encode(autoText));

// Async: hand it gzip-framed bytes — it detects gzip.
const autoFromGzip = await Compression.decompressAuto(asGzip);
console.log(
  `  Compression.decompressAuto(gzip) identical: ${decoder.decode(autoFromGzip) === autoText}`
);

// Sync: hand it zlib-framed bytes — it detects zlib.
const autoFromZlib = Compression.decompressAutoSync(asZlib);
console.log(
  `  Compression.decompressAutoSync(zlib) identical: ${decoder.decode(autoFromZlib) === autoText}`
);

// =============================================================================
// 6. TAR entry type guards — Archive.isTarFile / isTarDirectory / isTarSymlink
// =============================================================================
//
// Build a TAR containing a regular file and a directory entry, read it back
// with Archive.openTar, then classify each entry by its .info.

console.log("\n6. TAR type guards (isTarFile / isTarDirectory / isTarSymlink):");

const tarBytes = Archive.packTarSync(
  new Map<string, string | Uint8Array>([
    ["docs/", ""], // trailing slash → directory entry
    ["docs/readme.txt", "A regular file inside a directory."],
    ["data.bin", new Uint8Array([1, 2, 3, 4])]
  ])
);

const reader = Archive.openTar(tarBytes);
for await (const entry of reader.entries()) {
  const info = entry.info;
  const kind = Archive.isTarDirectory(info)
    ? "directory"
    : Archive.isTarSymlink(info)
      ? "symlink"
      : Archive.isTarFile(info)
        ? "file"
        : "other";
  console.log(
    `  ${entry.path.padEnd(18)} type=${info.type} → ` +
      `isTarFile=${Archive.isTarFile(info)} ` +
      `isTarDirectory=${Archive.isTarDirectory(info)} ` +
      `isTarSymlink=${Archive.isTarSymlink(info)} (${kind})`
  );
}

console.log("\n=== Archive Compression Streams / Auto-Detect / TAR Guard Examples Complete ===");
