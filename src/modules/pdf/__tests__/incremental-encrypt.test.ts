import { PdfDict } from "@pdf/core/pdf-object";
import { buildIncremental } from "@pdf/core/pdf-writer";
import type { IncrementalObject } from "@pdf/core/pdf-writer";
import { describe, expect, it } from "vitest";

const ORIGINAL = new TextEncoder().encode(
  "%PDF-1.7\n1 0 obj\n<< /Type /Catalog >>\nendobj\nxref\n0 2\n0000000000 65535 f \n0000000009 00000 n \ntrailer << /Size 2 /Root 1 0 R >>\nstartxref\n44\n%%EOF\n"
);
const latin1 = (bytes: Uint8Array) => new TextDecoder("latin1").decode(bytes);

function build(encryptMetadata: boolean | undefined): string {
  const objects = new Map<number, IncrementalObject>([
    [2, { gen: 0, body: "<< /Title (abc) >>" }],
    [
      3,
      {
        gen: 0,
        body: {
          dict: new PdfDict().set("Type", "/Metadata").set("Subtype", "/XML"),
          data: new TextEncoder().encode("XMPDATA")
        }
      }
    ],
    [4, { gen: 0, body: { dict: new PdfDict(), data: new TextEncoder().encode("CONTENT") } }]
  ]);
  const out = buildIncremental(ORIGINAL, objects, new Map([["Size", "2"]]), {
    xrefStream: false,
    prevXrefOffset: 44,
    // Tags the data with its kind instead of encrypting it.
    encrypt: (data, _objNum, _gen, kind) => new TextEncoder().encode(`${kind}:${latin1(data)}`),
    encryptMetadata
  });
  return latin1(out.subarray(ORIGINAL.length));
}

describe("buildIncremental encryption", () => {
  it("passes strings and stream data to the handler with their kind", () => {
    const out = build(undefined);
    expect(out).toContain("<737472696e673a616263>"); // "string:abc" as hex
    expect(out).toContain("stream:XMPDATA");
    expect(out).toContain("stream:CONTENT");
  });

  it("keeps a metadata stream in the clear when /EncryptMetadata is false", () => {
    const out = build(false);
    expect(out).toContain("stream\nXMPDATA");
    expect(out).toContain("stream:CONTENT");
  });
});
