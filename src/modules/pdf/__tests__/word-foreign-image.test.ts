import { foreignImageDocx } from "@word/__tests__/foreign-image.helpers";
import { Document } from "@word/index";
import { describe, expect, it } from "vitest";

import { Pdf } from "../index";

describe("Word → PDF with pictures typed by the source package", () => {
  it("draws the JPEG stored as .jfif and skips the .emz without failing", async () => {
    const doc = await Document.read(await foreignImageDocx());
    const bytes = await Pdf.fromDocx(doc);
    const text = new TextDecoder("latin1").decode(bytes);
    expect(text).toContain("/DCTDecode");
  });
});
