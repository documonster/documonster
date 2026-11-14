/**
 * A Flat OPC part's `binaryData` decodes to its bytes — read in one call or a step at a time — however its base64 is
 * laid out: wrapped or not, and whatever its length against the steps it is decoded in; and a malformed package is
 * refused for its first fault either way.
 */
import { createYieldBudget } from "@utils/yield";
import { describe, expect, it } from "vitest";

import { parseFlatOpc, parseFlatOpcPaced } from "../convert/flat-opc";

/** The bytes a step of the decoder — 256 Ki characters of unbroken base64 — decodes to. */
const STEP = ((256 * 1024) / 4) * 3;

function base64(bytes: Uint8Array): string {
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  }
  return btoa(binary);
}

const EMPTY_PACKAGE = `<?xml version="1.0"?><pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"></pkg:package>`;

/** A Flat OPC package of `parts`, each a name and its `pkg:part` content. */
function flatOpc(parts: readonly (readonly [string, string])[], after = ""): Uint8Array {
  const body = parts
    .map(
      ([name, content]) =>
        `<pkg:part pkg:name="/${name}" pkg:contentType="application/octet-stream">${content}</pkg:part>`
    )
    .join("");
  return new TextEncoder().encode(
    EMPTY_PACKAGE.replace("</pkg:package>", `${body}${after}</pkg:package>`)
  );
}

/** A package of one binary part, `word/media/x.bin`, holding `binaryData`. */
function binaryPart(binaryData: string): Uint8Array {
  return flatOpc([["word/media/x.bin", `<pkg:binaryData>${binaryData}</pkg:binaryData>`]]);
}

describe("a Flat OPC part's binaryData", () => {
  it.each([0, 1, 2, 3, 5, STEP - 1, STEP + 1, 3 * STEP + 2])(
    "decodes %i bytes, read whole or in steps, wrapped or not",
    async length => {
      const data = new Uint8Array(length).map((_, i) => (i * 37) & 255);
      for (const wrap of [0, 76]) {
        const text = wrap === 0 ? base64(data) : base64(data).replace(/(.{76})/g, "$1\n  ");
        const flat = binaryPart(text);
        for (const parts of [
          parseFlatOpc(flat),
          await parseFlatOpcPaced(flat, createYieldBudget())
        ]) {
          const read = parts.get("word/media/x.bin")!;
          expect(read.length).toBe(length);
          expect(read.every((byte, i) => byte === data[i])).toBe(true);
          // Allocated at its size, not sized by the text and cut down: a cut view would keep the larger buffer.
          expect(read.buffer.byteLength).toBe(length);
        }
      }
    }
  );

  it("skips XML's white space in it, and only that", () => {
    const parts = parseFlatOpc(binaryPart("QU\tJD\r\n RUY="));
    expect([...parts.get("word/media/x.bin")!]).toEqual([0x41, 0x42, 0x43, 0x45, 0x46]);
    // A no-break space is not white space in XML, and not base64 either.
    expect(() => parseFlatOpc(binaryPart("QU\u00a0JD"))).toThrow(/is not base64/);
  });

  it.each(["QQ=", "QQ==QQ==", "QQ==\nQQ", "Q!==", "-_-_"])("refuses %j", async text => {
    const flat = binaryPart(text);
    expect(() => parseFlatOpc(flat)).toThrow(/is not base64/);
    await expect(parseFlatOpcPaced(flat, createYieldBudget())).rejects.toThrow(/is not base64/);
  });
});

describe("a malformed Flat OPC package", () => {
  // A part's content is finished after the event that ends it, so a read in steps ran ahead of it: it reported a
  // malformation later in the document, where a whole read reports the first.
  it.each<[string, Uint8Array]>([
    [
      "a part that is not base64, then malformed XML",
      flatOpc([["a", "<pkg:binaryData>Q!==</pkg:binaryData>"]], "<oops")
    ],
    [
      "a part that is not base64, then a part holding an element",
      flatOpc([
        ["a", "<pkg:binaryData>Q!==</pkg:binaryData>"],
        ["b", "<pkg:binaryData>QQ<x/></pkg:binaryData>"]
      ])
    ],
    [
      "a part whose base64 is broken, then a second data element",
      flatOpc([["a", "<pkg:binaryData>Q!==</pkg:binaryData><pkg:binaryData>QUJD</pkg:binaryData>"]])
    ]
  ])("is refused for its first fault, read whole or in steps: %s", async (_name, flat) => {
    expect(() => parseFlatOpc(flat)).toThrow('the binaryData of part "a" is not base64');
    for (const yieldPoint of [() => undefined, () => Promise.resolve()]) {
      await expect(parseFlatOpcPaced(flat, yieldPoint)).rejects.toThrow(
        'the binaryData of part "a" is not base64'
      );
    }
  });
});

describe("an empty Flat OPC document", () => {
  // A read in steps gave the parser no slice of an empty body, and its error lost the byte it names.
  it.each(["", "   ", "\ufeff"])("is refused as a whole read refuses it: %j", async text => {
    const bytes = new TextEncoder().encode(text);
    let expected = "";
    try {
      parseFlatOpc(bytes);
    } catch (error) {
      expected = (error as Error).message;
    }
    expect(expected).toMatch(/at byte \d+/);
    await expect(parseFlatOpcPaced(bytes, () => undefined)).rejects.toThrow(expected);
  });
});
