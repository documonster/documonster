import type { ZipStringCodec } from "@archive/shared/text";
import { ZipParser } from "@archive/unzip/zip-parser";
import { parseZipExtraFields } from "@archive/zip-spec/zip-extra-fields";
import { createZip } from "@archive/zip/zip-bytes";
import { describe, it, expect } from "vitest";

/**
 * Creates a simple custom codec for testing purposes.
 * Maps '★' to 0x80, other chars use ASCII if < 0x80, else '?'.
 */
function createStarCodec(name: string): ZipStringCodec {
  return {
    name,
    encode(value: string): Uint8Array {
      const out: number[] = [];
      for (const ch of value) {
        if (ch === "★") {
          out.push(0x80);
        } else {
          const code = ch.codePointAt(0) ?? 0x3f;
          out.push(code < 0x80 ? code : 0x3f);
        }
      }
      return Uint8Array.from(out);
    },
    decode(bytes: Uint8Array): string {
      let result = "";
      for (const byte of bytes) {
        result += byte === 0x80 ? "★" : String.fromCharCode(byte);
      }
      return result;
    },
    useUtf8Flag: false,
    useUnicodeExtraFields: false
  };
}

describe("zip encoding", () => {
  it("should include Unicode extra fields for non-UTF8 encodings", async () => {
    const content = new TextEncoder().encode("hello");
    const name = "Grüße.txt";

    const zip = await createZip([{ name, data: content }], { encoding: "cp437" });
    const parser = new ZipParser(zip);
    const entries = parser.getEntries();

    expect(entries.length).toBe(1);
    expect(entries[0]!.path).toBe(name);

    const extra = parseZipExtraFields(entries[0]!.extraField ?? new Uint8Array(0), {
      compressedSize: entries[0]!.compressedSize,
      uncompressedSize: entries[0]!.uncompressedSize,
      offsetToLocalFileHeader: entries[0]!.localHeaderOffset
    });

    expect(extra.unicodePath?.unicodeValue).toBe(name);
  });

  it("should decode names using custom codec when Unicode extra fields are disabled", async () => {
    const customCodec = createStarCodec("x-test");

    const content = new TextEncoder().encode("payload");
    const name = "star★.txt";
    const comment = "note★";

    const zip = await createZip([{ name, data: content, comment, encoding: customCodec }]);

    const parser = new ZipParser(zip, { encoding: customCodec });
    const entries = parser.getEntries();

    expect(entries.length).toBe(1);
    expect(entries[0]!.path).toBe(name);
    expect(entries[0]!.comment).toBe(comment);
  });

  it("should decode archive comment using custom codec", async () => {
    const customCodec = createStarCodec("x-comment");

    const zip = await createZip([], { comment: "root★", encoding: customCodec });
    const parser = new ZipParser(zip, { encoding: customCodec });

    expect(parser.getZipComment()).toBe("root★");
  });
});
