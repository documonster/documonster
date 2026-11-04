/**
 * Example: stream errors + type guards
 *
 * Demonstrates:
 * - StreamError / StreamStateError / StreamTypeError /
 *   UnsupportedStreamTypeError — the stream error hierarchy
 * - isStreamError()  — narrow an unknown caught value to a StreamError
 * - Streams.isStream / isReadable / isWritable / Streams.isTransform / Streams.isDuplex —
 *   runtime type guards for stream shapes
 *
 * Usage: pnpm example --filter 07-errors-guards
 */

import {
  isReadable,
  isStreamError,
  isWritable,
  StreamError,
  StreamStateError,
  StreamTypeError,
  UnsupportedStreamTypeError,
  Streams
} from "@stream";

/** The error classes form a hierarchy rooted at StreamError. */
export function exampleErrorHierarchy(): void {
  const errors: StreamError[] = [
    new StreamError("generic failure"),
    new StreamStateError("write", "stream is destroyed"),
    new StreamTypeError("Uint8Array", "string"),
    new UnsupportedStreamTypeError("streamToBuffer", "object")
  ];

  for (const err of errors) {
    console.log(`${err.name}: "${err.message}" | isStreamError=${isStreamError(err)}`);
  }

  // A non-stream error is correctly rejected by the guard.
  console.log("plain Error isStreamError:", isStreamError(new Error("nope")));
}

/** Catch a StreamStateError emitted by a destroyed BufferedStream-like flow. */
export async function exampleCatchStreamError(): Promise<void> {
  try {
    // StreamStateError is thrown directly here to show typed-catch narrowing;
    // it is also emitted internally e.g. when writing to an ended stream.
    throw new StreamStateError("read", "stream already finished");
  } catch (err) {
    if (isStreamError(err)) {
      console.log("caught StreamError ->", err.name, "|", err.message);
    } else {
      throw err;
    }
  }
}

/** Runtime type guards distinguish the different stream shapes. */
export function exampleTypeGuards(): void {
  const readable = Streams.createReadableFromArray([1, 2, 3], { objectMode: true });
  const writable = Streams.createWritable<number>({
    objectMode: true,
    write: (_chunk, _encoding, callback) => callback()
  });
  const transformStream = Streams.createTransform<number, number>(n => n, { objectMode: true });
  const duplex = Streams.createDuplex<number, number>({ objectMode: true });

  console.log(
    "readable: Streams.isStream=%s isReadable=%s isWritable=%s",
    Streams.isStream(readable),
    isReadable(readable),
    isWritable(readable)
  );

  console.log(
    "writable: Streams.isStream=%s isReadable=%s isWritable=%s",
    Streams.isStream(writable),
    isReadable(writable),
    isWritable(writable)
  );

  console.log(
    "transform: Streams.isTransform=%s Streams.isDuplex=%s",
    Streams.isTransform(transformStream),
    Streams.isDuplex(transformStream)
  );

  console.log(
    "duplex: Streams.isDuplex=%s Streams.isTransform=%s",
    Streams.isDuplex(duplex),
    Streams.isTransform(duplex)
  );

  // Non-stream values are rejected.
  console.log("plain object isStream:", Streams.isStream({ foo: "bar" }));
}

export async function exampleErrorsGuards(): Promise<void> {
  console.log("=== Error hierarchy ===");
  exampleErrorHierarchy();
  console.log("\n=== Catch StreamError ===");
  await exampleCatchStreamError();
  console.log("\n=== Type guards ===");
  exampleTypeGuards();
}

await exampleErrorsGuards();
