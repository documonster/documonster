/**
 * Tests for stream.base utilities: bufferStream and PullStream.
 */

import { bufferStream, PullStream } from "@archive/unzip/stream.base";
import { Readable } from "@stream";
import { concatUint8Arrays } from "@utils/binary";
import { describe, it, expect } from "vitest";

const decoder = new TextDecoder();
const encoder = new TextEncoder();

// =============================================================================
// bufferStream
// =============================================================================

describe("bufferStream", () => {
  it("should concatenate all chunks into a single buffer", async () => {
    const input = Readable.from([
      encoder.encode("hello"),
      encoder.encode(" "),
      encoder.encode("world")
    ]);
    const result = await bufferStream(input);
    expect(decoder.decode(result)).toBe("hello world");
  });

  it("should handle empty stream", async () => {
    const input = Readable.from([]);
    const result = await bufferStream(input);
    expect(result.length).toBe(0);
  });

  it("should handle single chunk", async () => {
    const input = Readable.from([encoder.encode("single chunk")]);
    const result = await bufferStream(input);
    expect(decoder.decode(result)).toBe("single chunk");
  });

  it("should handle binary data", async () => {
    const binary = new Uint8Array([0x00, 0x01, 0x02, 0xff, 0xfe, 0xfd]);
    const input = Readable.from([binary]);
    const result = await bufferStream(input);
    expect(result.length).toBe(binary.length);
    expect([...result]).toEqual([...binary]);
  });

  it("should reject on stream error", async () => {
    const input = new Readable({
      read() {
        this.destroy(new Error("Stream error"));
      }
    });
    await expect(bufferStream(input)).rejects.toThrow("Stream error");
  });
});

// =============================================================================
// PullStream
// =============================================================================

describe("PullStream", () => {
  it("should pull exact number of bytes", async () => {
    const pull = new PullStream();
    const input = Readable.from([encoder.encode("hello world")]);
    input.pipe(pull);

    const result = await pull.pull(5);
    expect(decoder.decode(result)).toBe("hello");
  });

  it("should pull remaining bytes after first pull", async () => {
    const pull = new PullStream();
    const input = Readable.from([encoder.encode("hello world")]);
    input.pipe(pull);

    const first = await pull.pull(6);
    expect(decoder.decode(first)).toBe("hello ");

    const second = await pull.pull(5);
    expect(decoder.decode(second)).toBe("world");
  });

  it("should handle chunked input", async () => {
    const pull = new PullStream();
    const input = Readable.from([
      encoder.encode("hel"),
      encoder.encode("lo "),
      encoder.encode("world")
    ]);
    input.pipe(pull);

    const result = await pull.pull(11);
    expect(decoder.decode(result)).toBe("hello world");
  });

  it("should return empty buffer for pull(0)", async () => {
    const pull = new PullStream();
    const input = Readable.from([encoder.encode("hello")]);
    input.pipe(pull);

    const result = await pull.pull(0);
    expect(result.length).toBe(0);
  });

  it("should handle immediate data availability", async () => {
    const pull = new PullStream();
    pull.buffer = encoder.encode("immediate data");

    const result = await pull.pull(9);
    expect(decoder.decode(result)).toBe("immediate");
  });

  it("should emit error on FILE_ENDED when stream finishes before pull completes", async () => {
    const pull = new PullStream();
    const input = Readable.from([encoder.encode("short")]);
    input.pipe(pull);

    await new Promise(resolve => setTimeout(resolve, 10));
    await expect(pull.pull(100)).rejects.toThrow("FILE_ENDED");
  });

  it("should stream data until eof number of bytes", async () => {
    const pull = new PullStream();
    const input = Readable.from([encoder.encode("hello world")]);
    input.pipe(pull);

    const chunks: Uint8Array[] = [];
    const stream = pull.stream(5);

    await new Promise<void>(resolve => {
      stream.on("data", (chunk: Uint8Array) => chunks.push(chunk));
      stream.on("end", () => resolve());
    });

    expect(decoder.decode(concatUint8Arrays(chunks))).toBe("hello");
  });

  it("should stream data until eof pattern found", async () => {
    const pull = new PullStream();
    const input = Readable.from([encoder.encode("hello|world")]);
    input.pipe(pull);

    const chunks: Uint8Array[] = [];
    const eof = encoder.encode("|");
    const stream = pull.stream(eof);

    await new Promise<void>(resolve => {
      stream.on("data", (chunk: Uint8Array) => chunks.push(chunk));
      stream.on("end", () => resolve());
    });

    expect(decoder.decode(concatUint8Arrays(chunks))).toBe("hello");
  });
});
