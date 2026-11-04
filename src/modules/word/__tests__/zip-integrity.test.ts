import { Crc32MismatchError } from "@archive/core/errors";
import { Document, DocxParseError, Io } from "@word/index";
import { describe, expect, it } from "vitest";

describe("DOCX package integrity", () => {
  it("Io.read rejects a damaged part, naming the CRC mismatch as the cause", async () => {
    const doc = Document.create();
    Document.addParagraph(doc, "integrity marker text");
    const docx = await Io.toBuffer(doc, { compressionLevel: 0 });
    const bad = docx.slice();
    const at = new TextDecoder("latin1").decode(bad).indexOf("integrity marker");
    expect(at).toBeGreaterThan(0);
    bad[at] = "I".charCodeAt(0);

    const error = await Io.read(bad).then(
      () => undefined,
      (e: unknown) => e
    );
    expect(error).toBeInstanceOf(DocxParseError);
    expect((error as Error).cause).toBeInstanceOf(Crc32MismatchError);

    const lenient = await Io.read(bad, { checkCrc32: false });
    expect(JSON.stringify(lenient.body)).toContain("Integrity marker text");
  });
});
