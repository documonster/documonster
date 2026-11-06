/**
 * What changed when the PDF options checks moved onto `@utils/schema`: every refusal reads
 * `<call>: <field> must be <expected>, got <value>`, and a few values were reported at the
 * wrong place or with the wrong description.
 */

import { Pdf, PdfArgumentError } from "@pdf/index";
import { describe, expect, it } from "vitest";

function refusal(run: () => unknown): { field: string; message: string } {
  try {
    run();
  } catch (err) {
    expect(err).toBeInstanceOf(PdfArgumentError);
    return { field: (err as PdfArgumentError).field!, message: (err as Error).message };
  }
  throw new Error("expected a PdfArgumentError");
}

async function refusalAsync(
  run: () => Promise<unknown>
): Promise<{ field: string; message: string }> {
  try {
    await run();
  } catch (err) {
    expect(err).toBeInstanceOf(PdfArgumentError);
    return { field: (err as PdfArgumentError).field!, message: (err as Error).message };
  }
  throw new Error("expected a PdfArgumentError");
}

describe("PDF option refusals name the value they refused", () => {
  it("shows an invalid Date as one, not as a quoted phrase", () => {
    expect(refusal(() => new Pdf.Builder().setMetadata({ modDate: new Date(Number.NaN) }))).toEqual(
      {
        field: "metadata.modDate",
        message:
          "Pdf.Builder#setMetadata: metadata.modDate must be a valid Date, got an invalid Date"
      }
    );
  });

  it("says empty font bytes are empty", () => {
    expect(
      refusal(() => new Pdf.Builder().embedFonts({ default: { regular: new Uint8Array() } }))
    ).toEqual({
      field: "config.default.regular",
      message:
        "Pdf.Builder#embedFonts: config.default.regular must be a non-empty .ttf or .ttc file, got a Uint8Array"
    });
  });

  it("reports a missing sheet, column or image as undefined, not null", async () => {
    expect(
      await refusalAsync(() => Pdf.fromTable({ data: [], columns: [undefined] } as never))
    ).toEqual({
      field: "input.columns[0]",
      message: "Pdf.fromTable: input.columns[0] must be an options object, got undefined"
    });
    expect(
      refusal(() => Pdf.createDrawSurface(new Pdf.Builder().addPage(), undefined as never))
    ).toEqual({
      field: "rect",
      message: "Pdf.createDrawSurface: rect must be an options object, got undefined"
    });
  });

  it("refuses an array as a watermark position where it was given", async () => {
    const { field } = await refusalAsync(() =>
      Pdf.fromTable([["a"]], { watermark: { type: "text", text: "W", position: [1, 2] as never } })
    );
    expect(field).toBe("options.watermark.position");
  });

  it("writes the font / fonts conflict as the other refusals are written", async () => {
    const fonts = { default: { regular: new Uint8Array([1]) } };
    expect(
      await refusalAsync(() => Pdf.fromTable([["a"]], { font: new Uint8Array([1]), fonts }))
    ).toEqual({
      field: "options.fonts",
      message:
        "Pdf.fromTable: options.fonts must be omitted when options.font is given — options.fonts names every face, got an object"
    });
  });

  it("shows the refused value of Pdf.extract options, not its type", async () => {
    expect(
      (await refusalAsync(() => Pdf.extract(new Uint8Array([1]), { pages: 3 as never }))).message
    ).toBe("Pdf.extract: options.pages must be an array of 1-based page numbers, got 3");
    expect(
      (await refusalAsync(() => Pdf.extract(new Uint8Array([1]), { extractText: "yes" as never })))
        .message
    ).toBe('Pdf.extract: options.extractText must be a boolean, got "yes"');
    expect(
      (
        await refusalAsync(() =>
          Pdf.extract(new Uint8Array([1]), { maxDecodedBytes: "1" as never })
        )
      ).message
    ).toBe('Pdf.extract: options.maxDecodedBytes must be a number >= 0 or Infinity, got "1"');
  });

  it("shows the refused value of the signature arguments", async () => {
    expect(
      (await refusalAsync(() => Pdf.Signature.verify(new Uint8Array([1]), "zz", [0, 0, 0, 0])))
        .message
    ).toBe(
      'Pdf.Signature.verify: signatureHex must be the hex string of /Contents, without its angle brackets, got "zz"'
    );
    expect(
      refusal(() =>
        new Pdf.Builder().sign({ certificate: [] as never, privateKey: new Uint8Array() })
      ).message
    ).toBe("Pdf.Builder#sign: options.certificate must be a Uint8Array of DER bytes, got an array");
  });

  it("shows the refused recalculate", async () => {
    expect(
      (await refusalAsync(() => Pdf.fromExcel(null as never, { recalculate: 1 as never }))).message
    ).toBe(
      "Pdf.fromExcel: options.recalculate must be a function, such as ExcelFormula.calculate, got 1"
    );
  });

  it("does not check onWarning on new Pdf.Builder, which has no such option", () => {
    expect(() => new Pdf.Builder({ onWarning: 1 } as never)).not.toThrow();
  });
});
