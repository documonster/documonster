/**
 * Stream Module Examples
 *
 * Demonstrates usage of universal stream utilities.
 * All examples work identically in browser and Node.js.
 */

import { Streams } from "@stream";
import { stringToUint8Array, uint8ArrayToString } from "@utils/binary";

/**
 * Example 1: Using Streams.createTransform for custom transformation
 */
export async function exampleTransformBasic(): Promise<void> {
  console.log("\n=== Streams.createTransform Example ===");

  const transform = Streams.createTransform<string, string>(chunk => chunk.toUpperCase(), {
    objectMode: true
  });

  transform.on("data", (chunk: string) => {
    console.log("Transformed:", chunk);
  });

  transform.write("hello world");
  transform.write("this is a test");
  transform.end();

  await new Promise<void>(resolve => transform.on("finish", resolve));
}

/**
 * Example 2: Using BufferedStream for accumulation
 */
export function exampleBufferedStream(): void {
  console.log("\n=== BufferedStream Example ===");

  const buffered = Streams.createBufferedStream({ batchSize: 1024 });

  // Write some data
  buffered.write("Hello ");
  buffered.write("World!");
  buffered.write(stringToUint8Array(" This is a test."));

  console.log("Buffered length:", buffered.bufferedLength);

  // Get all data as Uint8Array
  const allData = buffered.toUint8Array();
  console.log("All data:", uint8ArrayToString(allData));
}

/**
 * Example 3: Using ChunkedBuilder for efficient string building
 */
export function exampleChunkedBuilder(): void {
  console.log("\n=== ChunkedBuilder Example ===");

  const builder = new Streams.ChunkedBuilder({ chunkSize: 10 });

  // Simulate building an XML document
  builder.push('<?xml version="1.0"?>');
  builder.push("<root>");
  builder.pushAll(["<item>1</item>", "<item>2</item>", "<item>3</item>"]);
  builder.push("</root>");

  console.log("Built string:", builder.toString());
  console.log("Cursor position:", builder.cursor);
}

/**
 * Example 4: Using TransactionalChunkedBuilder for rollback support
 */
export function exampleTransactionalBuilder(): void {
  console.log("\n=== TransactionalChunkedBuilder Example ===");

  const builder = new Streams.TransactionalChunkedBuilder();

  builder.push("<record>");

  // Take a snapshot before making potentially invalid changes
  builder.snapshot();

  builder.push("<field>");
  builder.push("invalid data that might fail validation");

  // Simulate validation failure - rollback
  console.log("Before rollback:", builder.toString());
  builder.rollback();
  console.log("After rollback:", builder.toString());

  // Now add valid data
  builder.snapshot();
  builder.push("<name>Valid Name</name>");
  builder.commit();

  builder.push("</record>");

  console.log("Final result:", builder.toString());
}

/**
 * Example 5: Using PullStream for pattern-based reading
 */
export async function examplePullStream(): Promise<void> {
  console.log("\n=== PullStream Example ===");

  const stream = Streams.createPullStream();

  // Write some data
  stream.write(stringToUint8Array("Line1\nLine2\nLine3\n"));
  stream.end();

  // Pull data until newline using pull() with Buffer pattern
  const line1 = await stream.pull(Buffer.from("\n"), false);
  console.log("Line 1:", line1.toString());

  const line2 = await stream.pull(Buffer.from("\n"), false);
  console.log("Line 2:", line2.toString());

  // Pull remaining bytes
  const remaining = await stream.pull(6);
  console.log("Remaining:", remaining.toString());
}

/**
 * Example 6: Using Streams.createTransform factory
 */
export async function exampleCreateTransform(): Promise<void> {
  console.log("\n=== Streams.createTransform Example ===");

  // Create a transform that doubles numbers
  const doubler = Streams.createTransform<number, number>(n => n * 2, { objectMode: true });

  const results: number[] = [];
  doubler.on("data", n => results.push(n));

  doubler.write(1);
  doubler.write(2);
  doubler.write(3);
  doubler.end();

  await new Promise<void>(resolve => doubler.on("finish", resolve));

  console.log("Doubled values:", results);
}

/**
 * Example 7: Piping transforms together
 */
export async function examplePipeTransforms(): Promise<void> {
  console.log("\n=== Pipe Transforms Example ===");

  const upper = Streams.createTransform<string, string>(chunk => chunk.toUpperCase(), {
    objectMode: true
  });
  const prefixer = Streams.createTransform<string, string>(chunk => `[PREFIX] ${chunk}`, {
    objectMode: true
  });

  // Pipe: input -> uppercase -> add prefix -> output
  upper.pipe(prefixer);

  const results: string[] = [];
  prefixer.on("data", (chunk: string) => results.push(chunk));

  upper.write("hello");
  upper.write("world");
  upper.end();

  await new Promise<void>(resolve => prefixer.on("finish", resolve));

  console.log("Piped results:", results);
}

/**
 * Run all examples
 */
export async function runAllExamples(): Promise<void> {
  await exampleTransformBasic();
  exampleBufferedStream();
  exampleChunkedBuilder();
  exampleTransactionalBuilder();
  await examplePullStream();
  await exampleCreateTransform();
  await examplePipeTransforms();

  console.log("\n=== All examples completed ===");
}
