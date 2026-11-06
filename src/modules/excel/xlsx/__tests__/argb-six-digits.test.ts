/**
 * A six-digit `argb` is an opaque colour — `"FF0000"` is `"FFFF0000"` — written the same way by both writers: XLSB
 * always padded it, while XLSX wrote `rgb="FF0000"`, which `ST_UnsignedIntHex` (four bytes) does not allow.
 */
import { extractAll } from "@archive/unzip/extract";
import { Cell, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

function styled() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Cell.setValue(ws, "A1", "x");
  Cell.setFont(ws, "A1", { color: { argb: "FF0000" } });
  Cell.setFill(ws, "A1", { type: "pattern", pattern: "solid", fgColor: { argb: "00ff00" } });
  return wb;
}

describe("six-digit argb", () => {
  it("is written as eight hex digits in XLSX", async () => {
    const entries = await extractAll(await Workbook.toBuffer(styled()));
    const styles = new TextDecoder().decode(entries.get("xl/styles.xml")!.data);
    const rgbs = [...styles.matchAll(/rgb="([^"]*)"/g)].map(m => m[1]);
    expect(rgbs).toContain("FFFF0000");
    expect(rgbs).toContain("FF00ff00");
    for (const rgb of rgbs) {
      expect(rgb).toMatch(/^[0-9A-Fa-f]{8}$/);
    }
  });

  it.each(["xlsx", "xlsb"] as const)("reads back from %s as the padded colour", async format => {
    const back = await Workbook.read(await Workbook.toBuffer(styled(), { format }));
    const ws = Workbook.getWorksheet(back, "S")!;
    expect(Cell.getFont(ws, "A1")?.color?.argb?.toUpperCase()).toBe("FFFF0000");
    expect(
      (Cell.getFill(ws, "A1") as { fgColor?: { argb?: string } }).fgColor?.argb?.toUpperCase()
    ).toBe("FF00FF00");
  });
});
