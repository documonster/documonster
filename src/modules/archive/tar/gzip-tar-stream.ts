import { gzipSync } from "@archive/compression/compress";
import { crc32, crc32Finalize, crc32Update } from "@archive/compression/crc32";
import type {
  GunzipStream,
  GzipEncoder,
  GzipStream
} from "@archive/compression/streaming-compress";
import { createGzipEncoder } from "@archive/compression/streaming-compress";
import { concatUint8Arrays } from "@utils/binary";

const GZIP_PREMATURE_CLOSE_MESSAGE = "Gzip stream closed before completing";

/**
 * Push `tarChunks` through a gzip compressor — or, given a gunzip stream, a decompressor —
 * yielding output as it is produced, with backpressure.
 *
 * @internal
 */
export async function* gzipTarChunks(
  tarChunks: AsyncIterable<Uint8Array>,
  gzipStream: GzipStream | GunzipStream | GzipEncoder
): AsyncIterable<Uint8Array> {
  const outputChunks: Uint8Array[] = [];
  let outputHead = 0;
  let gzipError: Error | null = null;
  let gzipClosed = false;
  let completed = false;

  function* drainOutput(): Iterable<Uint8Array> {
    while (outputHead < outputChunks.length) {
      yield outputChunks[outputHead++]!;
    }
    if (outputHead > 0) {
      outputChunks.length = 0;
      outputHead = 0;
    }
  }

  function onData(chunk: Uint8Array): void {
    outputChunks.push(chunk);
  }

  function onError(error: Error): void {
    gzipError ??= error;
  }

  function onClose(): void {
    gzipClosed = true;
  }

  function currentFailure(): Error | null {
    return gzipError ?? (gzipClosed ? new Error(GZIP_PREMATURE_CLOSE_MESSAGE) : null);
  }

  async function waitForDrain(): Promise<void> {
    const failure = currentFailure();
    if (failure) {
      throw failure;
    }

    await new Promise<void>((resolve, reject) => {
      function cleanup(): void {
        gzipStream.off("drain", handleDrain);
        gzipStream.off("error", handleError);
        gzipStream.off("close", handleClose);
      }

      function handleDrain(): void {
        cleanup();
        resolve();
      }

      function handleError(error: Error): void {
        cleanup();
        reject(error);
      }

      function handleClose(): void {
        cleanup();
        reject(gzipError ?? new Error(GZIP_PREMATURE_CLOSE_MESSAGE));
      }

      gzipStream.once("drain", handleDrain);
      gzipStream.once("error", handleError);
      gzipStream.once("close", handleClose);

      const racedFailure = currentFailure();
      if (racedFailure) {
        cleanup();
        reject(racedFailure);
      }
    });
  }

  async function finishGzip(): Promise<void> {
    const failure = currentFailure();
    if (failure) {
      throw failure;
    }

    // Settle on `end`, not on the `end()` callback: Node's zlib streams run their final
    // flush after `finish`, so the callback can fire before the last output — and before
    // the error a truncated input produces
    await new Promise<void>((resolve, reject) => {
      function cleanup(): void {
        gzipStream.off("end", handleEnd);
        gzipStream.off("error", handleError);
        gzipStream.off("close", handleClose);
      }

      function handleEnd(): void {
        cleanup();
        resolve();
      }

      function handleError(error: Error): void {
        cleanup();
        reject(error);
      }

      function handleClose(): void {
        cleanup();
        reject(gzipError ?? new Error(GZIP_PREMATURE_CLOSE_MESSAGE));
      }

      gzipStream.once("end", handleEnd);
      gzipStream.once("error", handleError);
      gzipStream.once("close", handleClose);
      gzipStream.end((error?: Error | null) => {
        if (error) {
          handleError(error);
        }
      });
    });
  }

  /** Settle `step`, handing back its failure instead of throwing it. */
  function attempt(step: Promise<void>): Promise<unknown> {
    return step.then(
      () => undefined,
      (error: unknown) => error
    );
  }

  gzipStream.on("data", onData);
  gzipStream.on("error", onError);
  gzipStream.on("close", onClose);

  // Output produced before a failure is yielded ahead of it, so a consumer sees everything
  // that decoded cleanly before the error
  try {
    for await (const tarChunk of tarChunks) {
      const failure =
        currentFailure() ??
        (gzipStream.write(tarChunk) ? undefined : await attempt(waitForDrain()));
      yield* drainOutput();
      if (failure) {
        throw failure;
      }
    }

    const failure = await attempt(finishGzip());
    yield* drainOutput();
    if (failure) {
      throw failure;
    }
    completed = true;
  } finally {
    if (!completed && !gzipStream.destroyed) {
      gzipStream.destroy();
    }
    gzipStream.off("data", onData);
    gzipStream.off("close", onClose);
    // `onError` stays attached: zlib can still report an error for input it was processing
    // when the consumer stopped, and with no listener left that would crash the process
  }
}

// -----------------------------------------------------------------------------
// Compressing a whole TAR, streamed or in one call, to the same bytes
// -----------------------------------------------------------------------------

/** ID1, ID2, CM=deflate, no flags, MTIME=0, XFL=0, OS=unknown. */
const STORED_GZIP_HEADER = new Uint8Array([0x1f, 0x8b, 8, 0, 0, 0, 0, 0, 0, 255]);

/** The largest payload of one stored deflate block. */
const STORED_BLOCK_MAX = 65535;

/** One stored deflate block: BFINAL/BTYPE=00, then LEN and NLEN, then the bytes. */
function storedBlock(data: Uint8Array, final: boolean): Uint8Array {
  const block = new Uint8Array(5 + data.length);
  block[0] = final ? 1 : 0;
  block[1] = data.length & 0xff;
  block[2] = data.length >>> 8;
  block[3] = ~data.length & 0xff;
  block[4] = (~data.length >>> 8) & 0xff;
  block.set(data, 5);
  return block;
}

function gzipTrailer(crc: number, size: number): Uint8Array {
  const trailer = new Uint8Array(8);
  const view = new DataView(trailer.buffer);
  view.setUint32(0, crc >>> 0, true);
  view.setUint32(4, size >>> 0, true);
  return trailer;
}

/**
 * Level 0 is framed here rather than by zlib, whose stored blocks follow how the input
 * arrived — so streamed and one-call output would differ — and which a browser's
 * `CompressionStream` cannot produce at all.
 */
async function* storedGzipChunks(tarChunks: AsyncIterable<Uint8Array>): AsyncIterable<Uint8Array> {
  yield STORED_GZIP_HEADER;
  let crc = 0xffffffff;
  let size = 0;
  let pending: Uint8Array = new Uint8Array(0);
  for await (const chunk of tarChunks) {
    crc = crc32Update(crc, chunk);
    size += chunk.length;
    pending = pending.length === 0 ? chunk : concatUint8Arrays([pending, chunk]);
    // A full block is held back until more follows: the last one must carry BFINAL
    let offset = 0;
    while (pending.length - offset > STORED_BLOCK_MAX) {
      yield storedBlock(pending.subarray(offset, offset + STORED_BLOCK_MAX), false);
      offset += STORED_BLOCK_MAX;
    }
    pending = offset === 0 ? pending : pending.slice(offset);
  }
  yield storedBlock(pending, true);
  yield gzipTrailer(crc32Finalize(crc), size);
}

function storedGzip(tar: Uint8Array): Uint8Array {
  const parts: Uint8Array[] = [STORED_GZIP_HEADER];
  let offset = 0;
  do {
    const end = Math.min(offset + STORED_BLOCK_MAX, tar.length);
    parts.push(storedBlock(tar.subarray(offset, end), end === tar.length));
    offset = end;
  } while (offset < tar.length);
  parts.push(gzipTrailer(crc32(tar), tar.length));
  return concatUint8Arrays(parts);
}

/**
 * Gzip a TAR stream at `level`. In Node the output is byte-identical to
 * {@link gzipTarSync} of the same TAR; in a browser that holds at level 0, and otherwise
 * only where no native `CompressionStream` is available (see `createGzipEncoder`).
 *
 * @internal
 */
export function gzipTarStream(
  tarChunks: AsyncIterable<Uint8Array>,
  level: number
): AsyncIterable<Uint8Array> {
  return level === 0
    ? storedGzipChunks(tarChunks)
    : gzipTarChunks(tarChunks, createGzipEncoder({ level }));
}

/**
 * Gzip a whole TAR at `level` — the one-call form of {@link gzipTarStream}.
 *
 * @internal
 */
export function gzipTarSync(tar: Uint8Array, level: number): Uint8Array {
  return level === 0 ? storedGzip(tar) : gzipSync(tar, { level });
}
