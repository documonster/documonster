import { SharedStrings } from "@excel/utils/shared-strings";
import { describe, it, expect } from "vitest";

describe("SharedStrings", () => {
  it("Stores and shares string values", () => {
    const ss = new SharedStrings();

    const iHello = ss.add("Hello");
    const iHelloV2 = ss.add("Hello");
    const iGoodbye = ss.add("Goodbye");

    expect(iHello).toBe(iHelloV2);
    expect(iGoodbye).not.toBe(iHelloV2);

    expect(ss.count).toBe(2);
    expect(ss.totalRefs).toBe(3);
  });

  it("Does not escape values", () => {
    // that's the job of the xml utils
    const ss = new SharedStrings();

    const iXml = ss.add("<tag>value</tag>");
    const iAmpersand = ss.add("&");

    expect(ss.getString(iXml)).toBe("<tag>value</tag>");
    expect(ss.getString(iAmpersand)).toBe("&");
  });

  it("De-duplicates rich-text entries by structural equality", () => {
    // Before the fix, rich-text payloads were hashed as "[object Object]",
    // causing every rich-text shared string to collide on a single index.
    const ss = new SharedStrings();

    const a1 = ss.add({ richText: [{ text: "hello " }, { text: "world" }] });
    const a2 = ss.add({ richText: [{ text: "hello " }, { text: "world" }] });
    const b = ss.add({ richText: [{ text: "different" }] });

    expect(a1).toBe(a2);
    expect(b).not.toBe(a1);
    expect(ss.count).toBe(2);
    expect(ss.totalRefs).toBe(3);
  });

  it("Does not conflate a plain string with a rich-text entry of the same flattened text", () => {
    // The key prefixes `s:` / `r:` must keep these two value types in distinct
    // buckets even when their display text is identical.
    const ss = new SharedStrings();

    const iPlain = ss.add("hello");
    const iRich = ss.add({ richText: [{ text: "hello" }] });

    expect(iPlain).not.toBe(iRich);
    expect(ss.count).toBe(2);
  });

  it("De-duplicates rich-text runs regardless of property insertion order", () => {
    // Two run objects with the same fields but different key insertion order
    // are semantically identical and must hash to the same bucket.
    const ss = new SharedStrings();

    const runA: { text: string; font: { bold: boolean; italic: boolean } } = {
      text: "x",
      font: { bold: true, italic: false }
    };
    const runB: { font: { italic: boolean; bold: boolean }; text: string } = {
      font: { italic: false, bold: true },
      text: "x"
    };

    const iA = ss.add({ richText: [runA] });
    const iB = ss.add({ richText: [runB] });

    expect(iA).toBe(iB);
    expect(ss.count).toBe(1);
    expect(ss.totalRefs).toBe(2);
  });
});
