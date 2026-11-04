import { unzip } from "@archive/read-archive";
import { Cell, Column, ExcelArgumentError, Image, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

const BAD_LITERALS =
  /="(?:undefined|null|NaN|-?Infinity|\[object Object\])"|>(?:undefined|null|NaN|-?Infinity|\[object Object\])</;
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
const OBJ = {} as never;

type Ws = ReturnType<typeof Workbook.addWorksheet>;
type Wb = ReturnType<typeof Workbook.create>;

async function partsWithBadLiterals(bytes: Uint8Array): Promise<string[]> {
  const hits: string[] = [];
  for await (const entry of unzip(bytes).entries()) {
    const text = new TextDecoder().decode(await entry.bytes());
    if (BAD_LITERALS.test(text)) {
      hits.push(entry.path);
    }
  }
  return hits;
}

describe("wrong-typed model values never reach a workbook part", () => {
  it("a valid workbook carries none of the literals", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 1.5);
    Column.setWidth(ws, "A", 12.5);
    expect(await partsWithBadLiterals(await Workbook.toBuffer(wb))).toEqual([]);
  });

  it.each<[string, (wb: Wb, ws: Ws) => void]>([
    [
      "a picture size given as an object",
      (wb, ws) => {
        const id = Image.add(wb, { buffer: PNG, extension: "png" });
        Image.place(ws, id, { tl: { col: 0, row: 0 }, ext: { width: OBJ, height: 10 } });
      }
    ],
    [
      "a shape outline width given as an object",
      (_, ws) => Image.addShape(ws, { type: "rect", range: "B2:D5", lineWidth: OBJ })
    ]
  ])("%s is refused at the call, before anything is placed", async (_, apply) => {
    // Both used to be stored and fail only at write time, with an `XmlError` from inside the writer.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    expect(() => apply(wb, ws)).toThrow(ExcelArgumentError);
    expect(Image.list(ws)).toEqual([]);
    expect(Image.listShapes(ws)).toEqual([]);
    expect(await partsWithBadLiterals(await Workbook.toBuffer(wb))).toEqual([]);
  });

  it("a column width or a font size given as an object is refused at the call, before anything is stored", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    expect(() => Column.setWidth(ws, "A", OBJ)).toThrow(ExcelArgumentError);
    expect(Column.getWidth(ws, "A")).toBeUndefined();
    Cell.setValue(ws, "A1", 1);
    expect(() => Cell.setFont(ws, "A1", { size: OBJ })).toThrow(ExcelArgumentError);
    expect(Cell.getFont(ws, "A1")).toBeUndefined();
  });
});
