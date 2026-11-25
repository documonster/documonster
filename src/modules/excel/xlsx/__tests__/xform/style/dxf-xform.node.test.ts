import { extractAll } from "@archive/unzip/extract";
import { Cell, ConditionalFormatting, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

/**
 * A differential format's number format, read and written.
 *
 * Regression: a conditional format's number format used to survive the first save and then be destroyed by the next
 * one — the reader stored the `{ id, formatCode }` pair, the writer interpolated it as a string, and wrote
 * `formatCode="[object Object]"`. The reader now gives the code, as `Style.numFmt` promises for a read; the writer
 * still accepts the pair, which a caller may set.
 */
async function styleSheet(bytes: Uint8Array): Promise<string> {
  const entries = await extractAll(bytes);
  const entry = entries.get("xl/styles.xml");
  return entry ? new TextDecoder().decode(entry.data) : "";
}

const dxfs = (xml: string): string => /<dxfs[\s\S]*?<\/dxfs>/.exec(xml)?.[0] ?? "";

describe("dxf numFmt round-trip", () => {
  it("preserves the conditional-format number format across read → write", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 5);
    ConditionalFormatting.add(ws, {
      ref: "A1:A10",
      rules: [
        {
          type: "cellIs",
          operator: "greaterThan",
          formulae: [3],
          priority: 1,
          style: { numFmt: "#,##0.00", font: { bold: true } }
        }
      ]
    });

    const first = await Workbook.toBuffer(wb);
    expect(dxfs(await styleSheet(first))).toContain('formatCode="#,##0.00"');

    // Read it back — the parsed dxf carries the `{ id, formatCode }` pair …
    const reloaded = Workbook.create();
    await Workbook.read(reloaded, first);
    const reloadedSheet = Workbook.getWorksheet(reloaded, "S")!;
    // Read back as its code, as `Style.numFmt` promises — not the `{ id, formatCode }` pair `styles.xml` holds.
    expect(reloadedSheet.conditionalFormattings[0].rules[0].style?.numFmt).toBe("#,##0.00");

    // … and writing again must still emit the format code, not "[object Object]".
    const second = await Workbook.toBuffer(reloaded);
    const secondDxfs = dxfs(await styleSheet(second));
    expect(secondDxfs).toContain('formatCode="#,##0.00"');
    expect(secondDxfs).not.toContain("[object Object]");
  });

  it("reads a numFmt that states only its id as the built-in format the id names", async () => {
    const base = Workbook.create();
    const ws = Workbook.addWorksheet(base, "S");
    Cell.setValue(ws, "A1", 5);
    ConditionalFormatting.add(ws, {
      ref: "A1",
      rules: [{ type: "expression", formulae: ["TRUE"], style: { numFmt: "0.00" } }]
    });
    const { Archive } = await import("@archive/index");
    const editor = await Archive.ZipEditor.open(await Workbook.toBuffer(base));
    const styles = await styleSheet(await Workbook.toBuffer(base));
    // `formatCode` is required, but a producer that leaves it out still names a built-in format by its id.
    editor.set(
      "xl/styles.xml",
      new TextEncoder().encode(
        styles.replace(/<dxf><numFmt [^>]*\/>/, '<dxf><numFmt numFmtId="2"/>')
      )
    );
    const read = await Workbook.read(await editor.toBuffer());
    const rule = Workbook.getWorksheet(read, "S")!.conditionalFormattings[0]!.rules[0]!;
    expect(rule.style?.numFmt).toBe("0.00");
  });
});
