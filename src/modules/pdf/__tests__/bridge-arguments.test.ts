/**
 * `Pdf.fromExcel` `sheets` (X4) and the handle checks of the PDF entries (X10).
 */
import { Cell, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

import { Pdf, PdfArgumentError, createPdfDrawSurface } from "../index";

function book(): Workbook.Handle {
  const wb = Workbook.create();
  for (const name of ["First", "Second", "Third"]) {
    Cell.setValue(Workbook.addWorksheet(wb, name), "A1", `cell of ${name}`);
  }
  return wb;
}

async function rejection(run: () => Promise<unknown>): Promise<PdfArgumentError> {
  try {
    await run();
  } catch (err) {
    expect(err).toBeInstanceOf(PdfArgumentError);
    return err as PdfArgumentError;
  }
  throw new Error("expected a PdfArgumentError");
}

async function texts(bytes: Uint8Array): Promise<string[]> {
  return (await Pdf.extract(bytes)).pages.map(page => page.text.trim());
}

describe("Pdf.fromExcel sheets (X4)", () => {
  it("refuses 0: positions are 1-based", async () => {
    const err = await rejection(() => Pdf.fromExcel(book(), { sheets: [0] }));
    expect(err.field).toBe("options.sheets[0]");
    expect(err.message).toMatch(/^Pdf\.fromExcel: options\.sheets\[0\] .*1-based/);
  });

  it("refuses a name or position that matches no sheet, instead of skipping it", async () => {
    const unknown = await rejection(() => Pdf.fromExcel(book(), { sheets: ["Nope"] }));
    expect(unknown.field).toBe("options.sheets[0]");
    expect(unknown.message).toMatch(
      /^Pdf\.fromExcel: options\.sheets\[0\] — the workbook has no sheet named "Nope"/
    );
    expect(
      (await rejection(() => Pdf.fromExcel(book(), { sheets: ["Second", "Nope"] }))).field
    ).toBe("options.sheets[1]");
    expect((await rejection(() => Pdf.fromExcel(book(), { sheets: [4] }))).field).toBe(
      "options.sheets[0]"
    );
  });

  it("still selects by name (any case) and by 1-based position", async () => {
    expect(await texts(await Pdf.fromExcel(book(), { sheets: ["second", 3] }))).toEqual([
      "cell of Second",
      "cell of Third"
    ]);
    expect(await texts(await Pdf.fromExcel(book(), { sheets: [1] }))).toEqual(["cell of First"]);
  });
});

describe("Pdf.fromExcel cell text", () => {
  it("shows a cell without a number format as Excel and Cell.getDisplayText do", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", new Date(Date.UTC(2024, 2, 10)));
    Cell.setValue(ws, "A2", 0.1 + 0.2);
    const expected = [Cell.getDisplayText(ws, "A1"), Cell.getDisplayText(ws, "A2")];
    const text = (await texts(await Pdf.fromExcel(wb)))[0];
    for (const shown of expected) {
      expect(text).toContain(shown);
    }
    expect(text).not.toContain("0.30000000000000004");
  });
});

describe("handles are checked (X10)", () => {
  it("fromExcel and fromChart refuse what is not a workbook or a chart", async () => {
    for (const bad of [42, null, {}, []]) {
      const err = await rejection(() => Pdf.fromExcel(bad as never));
      expect(err.field).toBe("workbook");
      expect(err.message).toMatch(/must be a workbook/);
      expect((await rejection(() => Pdf.fromChart(bad as never))).field).toBe("chart");
    }
  });

  it("parseSvgPath refuses a non-string", () => {
    for (const bad of [42, null, {}]) {
      expect(() => Pdf.parseSvgPath(bad as never)).toThrow(PdfArgumentError);
    }
  });

  it("the deprecated createPdfDrawSurface checks its page", () => {
    for (const bad of [42, null, {}]) {
      expect(() =>
        createPdfDrawSurface(bad as never, { x: 0, y: 0, width: 10, height: 10 })
      ).toThrow(PdfArgumentError);
    }
  });
});
