/**
 * DOCX Module - Streaming Writer Tests
 */

import { describe, it, expect } from "vitest";

import { createDocxStream, textParagraph } from "../index";
import type { BodyContent } from "../types";

describe("StreamingDocxWriter", () => {
  it("basic usage: add 3 paragraphs, finalize, verify output", async () => {
    const stream = createDocxStream();
    stream.add(textParagraph("First"));
    stream.add(textParagraph("Second"));
    stream.add(textParagraph("Third"));

    const output = await stream.finalize();
    expect(output).toBeInstanceOf(Uint8Array);
    expect(output.length).toBeGreaterThan(0);
    // ZIP magic number
    expect(output[0]).toBe(0x50);
    expect(output[1]).toBe(0x4b);
  });

  it("addMany with 100 elements", async () => {
    const stream = createDocxStream();
    const elements: BodyContent[] = [];
    for (let i = 0; i < 100; i++) {
      elements.push(textParagraph(`Paragraph ${i}`));
    }
    stream.addMany(elements);
    expect(stream.elementCount).toBe(100);

    const output = await stream.finalize();
    expect(output).toBeInstanceOf(Uint8Array);
    expect(output.length).toBeGreaterThan(0);
  });

  it("progress callback reports progress", async () => {
    const stream = createDocxStream({ chunkSize: 5 });
    const reports: Array<{ elementsWritten: number; phase: string }> = [];
    stream.onProgress(info => {
      reports.push({ ...info });
    });

    for (let i = 0; i < 15; i++) {
      stream.add(textParagraph(`Para ${i}`));
    }
    await stream.finalize();

    // Should have reported at least at 5, 10, 15 elements
    expect(reports.some(r => r.elementsWritten === 5)).toBe(true);
    expect(reports.some(r => r.elementsWritten === 10)).toBe(true);
    expect(reports.some(r => r.elementsWritten === 15)).toBe(true);
    // Should have a "finalizing" phase report
    expect(reports.some(r => r.phase === "finalizing")).toBe(true);
  });

  it("throws on add after finalize", async () => {
    const stream = createDocxStream();
    stream.add(textParagraph("test"));
    await stream.finalize();

    expect(() => stream.add(textParagraph("too late"))).toThrow();
  });

  it("reset for reuse", async () => {
    const stream = createDocxStream();
    stream.add(textParagraph("first run"));
    await stream.finalize();

    stream.reset();
    stream.add(textParagraph("second run"));
    const output = await stream.finalize();
    expect(output).toBeInstanceOf(Uint8Array);
    expect(output.length).toBeGreaterThan(0);
    expect(stream.elementCount).toBe(1);
  });
});
