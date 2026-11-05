/**
 * Example: pipeline() (cross-platform, objectMode)
 *
 * Demonstrates:
 * - Using pipeline() to connect streams with proper error handling
 * - objectMode streams (strings in, strings out)
 *
 * pipeline() is usually preferred over manual pipe() chains when you want
 * a single Promise you can await and centralized error/cleanup handling.
 */

import { pipeline, Streams } from "@stream";

export async function examplePipelineObjectMode(): Promise<void> {
  const source = Streams.createReadableFromArray(["hello", "world"], { objectMode: true });

  const upper = Streams.createTransform<string, string>(s => s.toUpperCase(), {
    objectMode: true
  });

  const collector = new Streams.Collector<string>();

  await pipeline(source, upper, collector);

  console.log(collector.chunks);
}
