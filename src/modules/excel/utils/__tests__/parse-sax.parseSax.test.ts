/**
 * parseSax async generator function tests
 */

import { parseSax } from "@xml/sax";
import { describe, it, expect } from "vitest";

describe("parseSax", () => {
  // Helper to convert array to async iterable
  async function* toAsyncIterable(chunks: string[]): AsyncGenerator<string> {
    for (const chunk of chunks) {
      yield chunk;
    }
  }

  it("should parse simple XML as async generator", async () => {
    const chunks = ["<root>", "hello", "</root>"];
    const allEvents: any[] = [];

    for await (const events of parseSax(toAsyncIterable(chunks))) {
      allEvents.push(...events);
    }

    expect(allEvents.length).toBe(3);
    expect(allEvents[0].eventType).toBe("opentag");
    expect(allEvents[0].value.name).toBe("root");
    expect(allEvents[1].eventType).toBe("text");
    expect(allEvents[1].value).toBe("hello");
    expect(allEvents[2].eventType).toBe("closetag");
    expect(allEvents[2].value.name).toBe("root");
  });

  it("should handle multiple elements", async () => {
    const chunks = ['<row r="1"><c r="A1"><v>42</v></c></row>'];
    const allEvents: any[] = [];

    for await (const events of parseSax(toAsyncIterable(chunks))) {
      allEvents.push(...events);
    }

    const eventTypes = allEvents.map(e => e.eventType);
    expect(eventTypes).toEqual([
      "opentag",
      "opentag",
      "opentag",
      "text",
      "closetag",
      "closetag",
      "closetag"
    ]);
  });

  it("should handle empty chunks gracefully", async () => {
    const chunks = ["<root>", "", "text", "", "</root>"];
    const allEvents: any[] = [];

    for await (const events of parseSax(toAsyncIterable(chunks))) {
      allEvents.push(...events);
    }

    expect(allEvents.some(e => e.eventType === "text" && e.value === "text")).toBe(true);
  });

  it("should yield events in batches per chunk", async () => {
    const chunks = ["<root><a/>", "<b/></root>"];
    const batchSizes: number[] = [];

    for await (const events of parseSax(toAsyncIterable(chunks))) {
      batchSizes.push(events.length);
    }

    // chunk 1: opentag root + opentag a + closetag a = 3
    // chunk 2: opentag b + closetag b + closetag root = 3
    expect(batchSizes[0]).toBe(3);
    expect(batchSizes[1]).toBe(3);
  });

  it("should work with single chunk containing full XML", async () => {
    const chunks = ['<root attr="value">content</root>'];
    const allEvents: any[] = [];

    for await (const events of parseSax(toAsyncIterable(chunks))) {
      allEvents.push(...events);
    }

    expect(allEvents[0].eventType).toBe("opentag");
    expect(allEvents[0].value.attributes.attr).toBe("value");
    expect(allEvents[1].eventType).toBe("text");
    expect(allEvents[1].value).toBe("content");
  });

  it("should handle Excel-like XML structure", async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>`;
    const chunks = [xml];
    const allEvents: any[] = [];

    for await (const events of parseSax(toAsyncIterable(chunks))) {
      allEvents.push(...events);
    }

    const tagNames = allEvents.filter(e => e.eventType === "opentag").map(e => e.value.name);

    expect(tagNames).toContain("worksheet");
    expect(tagNames).toContain("sheetData");
    expect(tagNames).toContain("row");
    expect(tagNames).toContain("c");
    expect(tagNames).toContain("v");
  });
});
