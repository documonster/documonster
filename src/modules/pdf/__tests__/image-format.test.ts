import { describe, expect, it } from "vitest";

import { Pdf, PdfError } from "../index";

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

describe("image format at the API boundary", () => {
  it("page.drawImage rejects a MIME type and names the accepted values", () => {
    const page = new Pdf.Builder().addPage();
    const call = (): unknown =>
      page.drawImage({
        data: PNG,
        format: "image/png" as "png",
        x: 0,
        y: 0,
        width: 10,
        height: 10
      });
    expect(call).toThrow(PdfError);
    expect(call).toThrow(/jpeg, png/);
  });
});
