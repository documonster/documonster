/**
 * Stream.DocumentWriter — the `commit` / `toBuffer` lifecycle, and the
 * deprecated `finalize` it replaces.
 */

import { describe, it, expect } from "vitest";

import { DocxWriteError, Io, Query, Stream } from "../index";

function collectingSink(): { sink: WritableStream<Uint8Array>; bytes: () => Uint8Array } {
  const chunks: Uint8Array[] = [];
  const sink = new WritableStream<Uint8Array>({
    write(chunk) {
      chunks.push(chunk.slice());
    }
  });
  return {
    sink,
    bytes: () => {
      const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
      let at = 0;
      for (const c of chunks) {
        out.set(c, at);
        at += c.length;
      }
      return out;
    }
  };
}

describe("Stream.DocumentWriter — commit / toBuffer", () => {
  it("toBuffer returns a readable document", async () => {
    const writer = new Stream.DocumentWriter();
    writer.addText("hello");
    const bytes = await writer.toBuffer();
    const doc = await Io.read(bytes);
    expect(Query.getText(doc)).toContain("hello");
  });

  it("toBuffer after commit returns the same bytes, and may be repeated", async () => {
    const writer = new Stream.DocumentWriter();
    writer.addText("hello");
    await writer.commit();
    const first = await writer.toBuffer();
    const second = await writer.toBuffer();
    expect(second).toEqual(first);
    expect(Query.getText(await Io.read(first))).toContain("hello");
  });

  it("refuses a second commit", async () => {
    const writer = new Stream.DocumentWriter();
    await writer.commit();
    await expect(writer.commit()).rejects.toThrow(DocxWriteError);
    expect(() => writer.addText("late")).toThrow(DocxWriteError);
  });

  it("commit ends the sink and toBuffer is refused in sink mode", async () => {
    const { sink, bytes } = collectingSink();
    const writer = new Stream.DocumentWriter({ sink });
    writer.addText("streamed");
    await writer.commit();
    expect(Query.getText(await Io.read(bytes()))).toContain("streamed");
    await expect(writer.toBuffer()).rejects.toThrow(/options\.sink/);
    await expect(writer.toBuffer()).rejects.toBeInstanceOf(DocxWriteError);
  });

  it("refuses toBuffer in sink mode before anything is committed, and writes nothing", async () => {
    const { sink, bytes } = collectingSink();
    const writer = new Stream.DocumentWriter({ sink });
    await expect(writer.toBuffer()).rejects.toBeInstanceOf(DocxWriteError);
    await writer.commit();
    expect(bytes().length).toBeGreaterThan(0);
  });
});

describe("Stream.DocumentWriter#finalize (deprecated)", () => {
  it("returns the document bytes in buffered mode", async () => {
    const writer = new Stream.DocumentWriter();
    writer.addText("legacy");
    const bytes = await writer.finalize();
    expect(Query.getText(await Io.read(bytes))).toContain("legacy");
    await expect(writer.finalize()).rejects.toThrow(/already finalized/);
  });

  it("resolves to an empty array in sink mode and ends the sink", async () => {
    const { sink, bytes } = collectingSink();
    const writer = new Stream.DocumentWriter({ sink });
    writer.addText("legacy");
    const result = await writer.finalize();
    expect(result).toBeInstanceOf(Uint8Array);
    expect(result.length).toBe(0);
    expect(Query.getText(await Io.read(bytes()))).toContain("legacy");
  });
});
