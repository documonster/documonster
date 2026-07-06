import { Writable } from "stream";

import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { expectValidXlsx } from "@excel/__tests__/helpers/expect-valid-xlsx";
import { requireEntryText } from "@excel/__tests__/helpers/zip-text";
import { cellSetStyle, cellSetValue } from "@excel/core/cell";
import { Cell, Column, Row, Workbook } from "@excel/index";
import { WorkbookWriter } from "@excel/stream/workbook-writer";
import { getUniqueTestFilePath } from "@test/utils";
import { describe, expect, it } from "vitest";

describe("named cell styles", () => {
  it("defines, applies and round-trips a custom named style", async () => {
    const filename = getUniqueTestFilePath(import.meta.url);

    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");

    Workbook.defineCellStyle(wb, "Heading 1", {
      font: { name: "Arial", size: 20, bold: true, color: { argb: "FF1F4E79" } }
    });

    Cell.setValue(ws, "A1", "My heading");
    Cell.applyCellStyle(ws, "A1", "Heading 1");

    // registry accessors
    expect(Workbook.getCellStyle(wb, "Heading 1")?.font).toMatchObject({
      name: "Arial",
      size: 20,
      bold: true
    });
    expect(Workbook.listCellStyles(wb).map(s => s.name)).toEqual(["Heading 1"]);

    await expectValidXlsx(new Uint8Array(await Workbook.toBuffer(wb)));
    await Workbook.writeFile(wb, filename);

    const wb2 = Workbook.create();
    await Workbook.readFile(wb2, filename);
    const ws2 = Workbook.getWorksheet(wb2, "Sheet1")!;

    // cell keeps its named-style reference
    expect(Cell.getStyle(ws2, "A1").styleName).toBe("Heading 1");

    // the named style itself is preserved on the workbook
    const restored = Workbook.getCellStyle(wb2, "Heading 1");
    expect(restored).toBeTruthy();
    expect(restored?.font).toMatchObject({ name: "Arial", size: 20, bold: true });
  });

  it("applies a built-in preset with the correct builtinId", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");

    const name = Workbook.useBuiltinCellStyle(wb, "Heading1");
    expect(name).toBe("Heading 1");
    expect(Workbook.getCellStyle(wb, "Heading 1")?.builtinId).toBe(16);

    Cell.setValue(ws, "A1", "Title text");
    Cell.applyCellStyle(ws, "A1", name);

    const buffer = new Uint8Array(await Workbook.toBuffer(wb));
    await expectValidXlsx(buffer);

    const wb2 = Workbook.create();
    await Workbook.read(wb2, buffer);
    expect(Workbook.getCellStyle(wb2, "Heading 1")?.builtinId).toBe(16);
  });

  it("Cell.setStyle accepts styleName directly", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Workbook.defineCellStyle(wb, "Accent", {
      fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } }
    });

    Cell.setValue(ws, "B2", 42);
    Cell.setStyle(ws, "B2", { styleName: "Accent" });

    expect(Cell.getStyle(ws, "B2").styleName).toBe("Accent");
    await expectValidXlsx(new Uint8Array(await Workbook.toBuffer(wb)));
  });

  it("rejects redefining the reserved Normal style", () => {
    const wb = Workbook.create();
    expect(() => Workbook.defineCellStyle(wb, "Normal", {})).toThrow();
    expect(() => Workbook.defineCellStyle(wb, "", {})).toThrow();
  });

  it("removeCellStyle deletes a defined style", () => {
    const wb = Workbook.create();
    Workbook.defineCellStyle(wb, "Temp", { font: { italic: true } });
    expect(Workbook.removeCellStyle(wb, "Temp")).toBe(true);
    expect(Workbook.getCellStyle(wb, "Temp")).toBeUndefined();
    expect(Workbook.removeCellStyle(wb, "Missing")).toBe(false);
  });

  it("streaming writer supports named styles", async () => {
    const filename = getUniqueTestFilePath(import.meta.url);

    const wb = new WorkbookWriter({ filename, useStyles: true });
    wb.defineCellStyle("Heading 1", {
      font: { name: "Arial", size: 20, bold: true }
    });
    const ws = wb.addWorksheet("Sheet1");
    const a1 = ws.getCell("A1");
    cellSetValue(a1, "Streamed heading");
    cellSetStyle(a1, { styleName: "Heading 1" });
    ws.commit();
    await wb.commit();

    const wb2 = Workbook.create();
    await Workbook.readFile(wb2, filename);
    const ws2 = Workbook.getWorksheet(wb2, "Sheet1")!;
    expect(Cell.getStyle(ws2, "A1").styleName).toBe("Heading 1");
    expect(Workbook.getCellStyle(wb2, "Heading 1")?.font).toMatchObject({
      name: "Arial",
      size: 20,
      bold: true
    });
  });

  it("row-level named style propagates to cells and round-trips", async () => {
    const filename = getUniqueTestFilePath(import.meta.url);

    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Workbook.defineCellStyle(wb, "Heading 2", { font: { size: 13, bold: true } });

    // Apply the named style to an entire row, then add cells to it.
    Row.setStyle(ws, 1, { styleName: "Heading 2" });
    Cell.setValue(ws, "A1", "H2 a");
    Cell.setValue(ws, "B1", "H2 b");

    // Cells created after the row style inherit the row's styleName.
    expect(Cell.getStyle(ws, "A1").styleName).toBe("Heading 2");
    expect(Cell.getStyle(ws, "B1").styleName).toBe("Heading 2");

    await expectValidXlsx(new Uint8Array(await Workbook.toBuffer(wb)));
    await Workbook.writeFile(wb, filename);

    const wb2 = Workbook.create();
    await Workbook.readFile(wb2, filename);
    const ws2 = Workbook.getWorksheet(wb2, "Sheet1")!;
    expect(Cell.getStyle(ws2, "A1").styleName).toBe("Heading 2");
    expect(Cell.getStyle(ws2, "B1").styleName).toBe("Heading 2");
  });

  it("column-level named style is not dropped as a default column", async () => {
    const filename = getUniqueTestFilePath(import.meta.url);

    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Workbook.defineCellStyle(wb, "Accent", {
      fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } }
    });

    // A column whose ONLY styling is a named style must survive serialization.
    Column.setStyle(ws, "B", { styleName: "Accent" });
    Cell.setValue(ws, "B2", 42);

    await expectValidXlsx(new Uint8Array(await Workbook.toBuffer(wb)));
    await Workbook.writeFile(wb, filename);

    const wb2 = Workbook.create();
    await Workbook.readFile(wb2, filename);
    const ws2 = Workbook.getWorksheet(wb2, "Sheet1")!;
    expect(Cell.getStyle(ws2, "B2").styleName).toBe("Accent");
  });

  it("named style plus a local facet override coexist (inheritance)", async () => {
    const filename = getUniqueTestFilePath(import.meta.url);

    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Workbook.defineCellStyle(wb, "Heading 1", {
      font: { name: "Arial", size: 20, bold: true }
    });

    // Reference the named style AND set a local facet the style does not touch.
    Cell.setValue(ws, "A1", "Mixed");
    Cell.setStyle(ws, "A1", {
      styleName: "Heading 1",
      alignment: { horizontal: "center" }
    });

    expect(Cell.getStyle(ws, "A1").styleName).toBe("Heading 1");
    expect(Cell.getStyle(ws, "A1").alignment).toMatchObject({ horizontal: "center" });

    await expectValidXlsx(new Uint8Array(await Workbook.toBuffer(wb)));
    await Workbook.writeFile(wb, filename);

    const wb2 = Workbook.create();
    await Workbook.readFile(wb2, filename);
    const ws2 = Workbook.getWorksheet(wb2, "Sheet1")!;
    const style = Cell.getStyle(ws2, "A1");
    // Both the named-style reference and the local override survive.
    expect(style.styleName).toBe("Heading 1");
    expect(style.alignment).toMatchObject({ horizontal: "center" });
  });

  it("applyCellStyle preserves a previously set direct facet", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Workbook.defineCellStyle(wb, "Heading 1", { font: { bold: true } });

    Cell.setValue(ws, "A1", "x");
    Cell.setStyle(ws, "A1", { alignment: { horizontal: "right" } });
    Cell.applyCellStyle(ws, "A1", "Heading 1");

    const style = Cell.getStyle(ws, "A1");
    expect(style.styleName).toBe("Heading 1");
    // applyCellStyle only sets styleName; it must not wipe the earlier facet.
    expect(style.alignment).toMatchObject({ horizontal: "right" });
    await expectValidXlsx(new Uint8Array(await Workbook.toBuffer(wb)));
  });

  it("emits true inheritance: cell xf references the named style's xf, not a copy", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    // The named style owns a font; the cell only adds a local alignment.
    Workbook.defineCellStyle(wb, "Heading 1", {
      font: { name: "Arial", size: 20, bold: true }
    });
    Cell.setValue(ws, "A1", "Mixed");
    Cell.setStyle(ws, "A1", {
      styleName: "Heading 1",
      alignment: { horizontal: "center" }
    });

    const buffer = new Uint8Array(await Workbook.toBuffer(wb));
    const entries = await extractAll(buffer);
    const stylesXml = requireEntryText(entries, "xl/styles.xml");

    // The named style must live in cellStyleXfs + cellStyles (builtin id absent
    // for custom styles), and the cell's cellXfs entry must point at it via a
    // non-zero xfId while carrying no font of its own (fontId=0 = inherited).
    expect(stylesXml).toMatch(/<cellStyle name="Heading 1" xfId="1"\/>/);

    // Find the cell's <xf> in cellXfs: it references xfId="1" and applies
    // alignment (its own facet) but NOT font — the font is inherited.
    const cellXfsBlock = stylesXml.slice(
      stylesXml.indexOf("<cellXfs"),
      stylesXml.indexOf("</cellXfs>")
    );
    const inheritingXf = cellXfsBlock
      .split("<xf ")
      .find(xf => xf.includes('xfId="1"') && xf.includes("applyAlignment"));
    expect(inheritingXf).toBeDefined();
    // Inherited font ⇒ fontId is 0 and applyFont is absent.
    expect(inheritingXf).toMatch(/fontId="0"/);
    expect(inheritingXf).not.toMatch(/applyFont/);
  });

  it("applyCellStyle throws for an undefined named style", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", "x");
    expect(() => Cell.applyCellStyle(ws, "A1", "Nope")).toThrow(/not defined/);
    // Low-level setStyle stays lenient (no workbook lookup / no throw).
    expect(() => Cell.setStyle(ws, "A1", { styleName: "Nope" })).not.toThrow();
  });

  it("streaming defineCellStyle validates name and dedups on redefine", async () => {
    const chunks: Uint8Array[] = [];
    const stream = new Writable({
      write(chunk: Uint8Array, _enc: string, cb: () => void) {
        chunks.push(chunk);
        cb();
      }
    });
    const wb = new WorkbookWriter({ stream, useStyles: true });
    expect(() => wb.defineCellStyle("", { font: { bold: true } })).toThrow();
    expect(() => wb.defineCellStyle("Normal", { font: { bold: true } })).toThrow();

    // Redefining the same name must not emit a duplicate <cellStyle>.
    wb.defineCellStyle("Accent", { font: { bold: true } });
    wb.defineCellStyle("Accent", { font: { italic: true } });
    const ws = wb.addWorksheet("Sheet1");
    const a1 = ws.getCell("A1");
    cellSetValue(a1, "x");
    cellSetStyle(a1, { styleName: "Accent" });
    ws.commit();
    await wb.commit();

    const total = chunks.reduce((n, c) => n + c.length, 0);
    const buffer = new Uint8Array(total);
    let offset = 0;
    for (const c of chunks) {
      buffer.set(c, offset);
      offset += c.length;
    }

    const entries = await extractAll(buffer);
    const stylesXml = requireEntryText(entries, "xl/styles.xml");
    const accentCount = (stylesXml.match(/<cellStyle name="Accent"/g) ?? []).length;
    expect(accentCount).toBe(1);
  });

  it("preserves cellStyle hidden/customBuiltin/iLevel across a full read→write round-trip", async () => {
    // 1. Author a workbook with a custom named style and emit it.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Workbook.defineCellStyle(wb, "Custom", { font: { bold: true } });
    Cell.setValue(ws, "A1", "x");
    Cell.applyCellStyle(ws, "A1", "Custom");
    const first = new Uint8Array(await Workbook.toBuffer(wb));

    // 2. Inject the extra OOXML attributes onto the <cellStyle> and repack.
    const entries = await extractAll(first);
    const stylesXml = new TextDecoder().decode(entries.get("xl/styles.xml")!.data);
    const patched = stylesXml.replace(
      /<cellStyle name="Custom" xfId="1"\s*\/>/,
      '<cellStyle name="Custom" xfId="1" iLevel="2" hidden="1" customBuiltin="1"/>'
    );
    expect(patched).not.toBe(stylesXml); // replacement actually happened
    const repackEntries = [...entries].map(([name, file]) => ({
      name,
      data: name === "xl/styles.xml" ? new TextEncoder().encode(patched) : file.data
    }));
    const repacked = await createZip(repackEntries);

    // 3. Read the patched package, then write it out again.
    const wb2 = Workbook.create();
    await Workbook.read(wb2, repacked);
    const outBytes = new Uint8Array(await Workbook.toBuffer(wb2));

    // 4. The extra attributes must survive the round-trip.
    const outEntries = await extractAll(outBytes);
    const outStyles = new TextDecoder().decode(outEntries.get("xl/styles.xml")!.data);
    const customCellStyle = /<cellStyle name="Custom"[^/]*\/>/.exec(outStyles)?.[0];
    expect(customCellStyle).toBeDefined();
    expect(customCellStyle).toMatch(/iLevel="2"/);
    expect(customCellStyle).toMatch(/hidden="1"/);
    expect(customCellStyle).toMatch(/customBuiltin="1"/);
  });

  it("keeps distinct names for two identical-content named styles on round-trip", async () => {
    // Regression: named-style xfs must NOT be content-deduplicated, otherwise
    // two differently-named styles with identical formatting collapse onto one
    // cellStyleXf and read back as the same name.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Workbook.defineCellStyle(wb, "Heading A", { font: { bold: true } });
    Workbook.defineCellStyle(wb, "Heading B", { font: { bold: true } }); // identical content
    Cell.setValue(ws, "A1", "a");
    Cell.applyCellStyle(ws, "A1", "Heading A");
    Cell.setValue(ws, "B1", "b");
    Cell.applyCellStyle(ws, "B1", "Heading B");

    const buffer = new Uint8Array(await Workbook.toBuffer(wb));
    await expectValidXlsx(buffer);
    const wb2 = Workbook.create();
    await Workbook.read(wb2, buffer);
    const ws2 = Workbook.getWorksheet(wb2, "Sheet1")!;
    expect(Cell.getStyle(ws2, "A1").styleName).toBe("Heading A");
    expect(Cell.getStyle(ws2, "B1").styleName).toBe("Heading B");
  });

  it("supports multiple named styles round-trip with correct name mapping", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Workbook.defineCellStyle(wb, "S1", { font: { bold: true } });
    Workbook.defineCellStyle(wb, "S2", { font: { italic: true } });
    Workbook.defineCellStyle(wb, "S3", { numFmt: "0.00%" });
    Cell.setValue(ws, "A1", 1);
    Cell.applyCellStyle(ws, "A1", "S3");
    Cell.setValue(ws, "A2", 2);
    Cell.applyCellStyle(ws, "A2", "S1");
    Cell.setValue(ws, "A3", 3);
    Cell.applyCellStyle(ws, "A3", "S2");

    const buffer = new Uint8Array(await Workbook.toBuffer(wb));
    await expectValidXlsx(buffer);
    const wb2 = Workbook.create();
    await Workbook.read(wb2, buffer);
    const ws2 = Workbook.getWorksheet(wb2, "Sheet1")!;
    expect(Cell.getStyle(ws2, "A1").styleName).toBe("S3");
    expect(Cell.getStyle(ws2, "A2").styleName).toBe("S1");
    expect(Cell.getStyle(ws2, "A3").styleName).toBe("S2");
    // S3 carries a numFmt in its base xf — must survive.
    expect(Workbook.getCellStyle(wb2, "S3")?.numFmt).toBe("0.00%");
  });

  it("applyCellStyle rejects the reserved built-in Normal style", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", "x");
    // "Normal" cannot be defined, so applyCellStyle should reject it clearly.
    expect(() => Cell.applyCellStyle(ws, "A1", "Normal")).toThrow(/not defined/);
  });

  it("applyCellStyle works with the (ws, row, col, name) overload", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Workbook.defineCellStyle(wb, "H", { font: { bold: true } });
    Cell.setValue(ws, 2, 3, "x"); // C2
    Cell.applyCellStyle(ws, 2, 3, "H");
    expect(Cell.getStyle(ws, "C2").styleName).toBe("H");
    // Undefined name via the same overload must still throw.
    expect(() => Cell.applyCellStyle(ws, 2, 3, "Missing")).toThrow(/not defined/);
    await expectValidXlsx(new Uint8Array(await Workbook.toBuffer(wb)));
  });

  it("setStyle keeps styleName when a later setStyle changes another facet", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Workbook.defineCellStyle(wb, "H", { font: { bold: true } });
    Cell.setValue(ws, "A1", "x");
    Cell.setStyle(ws, "A1", { styleName: "H" });
    Cell.setStyle(ws, "A1", { alignment: { horizontal: "center" } });
    expect(Cell.getStyle(ws, "A1").styleName).toBe("H");
    expect(Cell.getStyle(ws, "A1").alignment).toMatchObject({ horizontal: "center" });
  });

  it("gracefully degrades a cell whose named style was removed", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Workbook.defineCellStyle(wb, "Gone", { font: { bold: true } });
    Cell.setValue(ws, "A1", "x");
    Cell.applyCellStyle(ws, "A1", "Gone");
    // Remove the definition but leave the cell's dangling reference.
    Workbook.removeCellStyle(wb, "Gone");
    // Writing must not throw; the cell falls back to a plain style.
    const buffer = new Uint8Array(await Workbook.toBuffer(wb));
    await expectValidXlsx(buffer);
    const wb2 = Workbook.create();
    await Workbook.read(wb2, buffer);
    const ws2 = Workbook.getWorksheet(wb2, "Sheet1")!;
    expect(Cell.getStyle(ws2, "A1").styleName).toBeUndefined();
    expect(Workbook.getCellStyle(wb2, "Gone")).toBeUndefined();
  });

  it("is stable across repeated toBuffer() calls", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Workbook.defineCellStyle(wb, "H", { font: { name: "Arial", bold: true } });
    Cell.setValue(ws, "A1", "x");
    Cell.applyCellStyle(ws, "A1", "H");

    const b1 = new Uint8Array(await Workbook.toBuffer(wb));
    const b2 = new Uint8Array(await Workbook.toBuffer(wb));
    // Both writes must round-trip to the same named-style reference.
    for (const b of [b1, b2]) {
      const r = Workbook.create();
      await Workbook.read(r, b);
      expect(Cell.getStyle(Workbook.getWorksheet(r, "Sheet1")!, "A1").styleName).toBe("H");
      expect(Workbook.getCellStyle(r, "H")?.font).toMatchObject({ name: "Arial", bold: true });
    }
  });
});
