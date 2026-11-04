/**
 * The invariant between the engine and the writers, for a formula set through the API: **the engine spills it ⇔ the
 * writers store it as a dynamic array** — XLSX, streaming XLSX, XLSB and streaming XLSB — and after a read-back the
 * two still agree. A formula a file stored as an ordinary one stays ordinary: Excel computes it with implicit
 * intersection, and so does the engine.
 */
import { Archive } from "@archive";
import { calculateFormulas } from "@excel/core/formula-adapter";
import { Cell, Row, Stream, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

/** Formulas that produce an array: each spills, and each is a dynamic array in every container. */
const SPILLING = [
  "SEQUENCE(3)",
  "_xlfn.SEQUENCE(2,2)",
  "FILTER(A1:A4,B1:B4>1)",
  "_xlfn._xlws.SORT(A1:A4)",
  "SORTBY(A1:A4,B1:B4,-1)",
  "UNIQUE(B1:B4)",
  "RANDARRAY(2,2)",
  "TOCOL(A1:B2)",
  "TOROW(A1:B2)",
  "VSTACK(A1:A2,B1:B2)",
  "HSTACK(A1:A2,B1:B2)",
  "TAKE(A1:A4,2)",
  "DROP(A1:A4,2)",
  "CHOOSECOLS(A1:B4,2)",
  "CHOOSEROWS(A1:B4,1,2)",
  "EXPAND(A1:A2,3,2,0)",
  "WRAPROWS(A1:A4,2)",
  "WRAPCOLS(A1:A4,2)",
  'TEXTSPLIT("a,b,c",",")',
  "MAKEARRAY(2,2,LAMBDA(r,c,r*c))",
  "MAP(A1:A3,LAMBDA(x,x*2))",
  "SCAN(0,A1:A3,LAMBDA(a,x,a+x))",
  "BYROW(A1:B3,LAMBDA(r,SUM(r)))",
  "BYCOL(A1:B3,LAMBDA(c,SUM(c)))",
  "MUNIT(2)",
  "TRANSPOSE(A1:A3)",
  "MMULT(A1:B2,A1:B2)",
  "{1,2,3}",
  "SEQUENCE(3)*2",
  "ABS(-SEQUENCE(3))",
  "LET(x,SEQUENCE(3),x*2)",
  "LAMBDA(x,SEQUENCE(x))(3)",
  "XLOOKUP(2,A1:A4,A1:B4)",
  'REGEXEXTRACT("a1b22c333","\\d+",1)',
  "A1:A3+{1;2;3}"
];

/** Formulas that stay ordinary: the engine intersects them, as Excel does with the file. */
const ORDINARY = [
  "A1:A3",
  "A1:A3*2",
  "ABS(A1:A3)",
  "INDEX(A1:B4,0,1)",
  "OFFSET(A1,0,0,3)",
  "SUM(SEQUENCE(3))",
  "ROWS(SEQUENCE(3))",
  "1+2",
  'TEXTAFTER("a-b","-")',
  'REGEXEXTRACT("a1b22","\\d+")',
  "@SEQUENCE(3)"
];

/** A dynamic array whose result is one value: stored as one, and computed alike either way. */
const DYNAMIC_SCALAR = [
  "XLOOKUP(2,A1:A4,B1:B4)",
  "LET(x,SEQUENCE(3),INDEX(x,2))",
  "INDEX({1,2},1,2)"
];

function data(ws: Parameters<typeof Row.addMany>[0]) {
  Row.addMany(ws, [
    [1, 4],
    [2, 1],
    [3, 3],
    [4, 2]
  ]);
}

/** One sheet per formula, at E6, so nothing it spills can reach the next. */
function book(formulas: readonly string[]) {
  const wb = Workbook.create();
  formulas.forEach((formula, i) => {
    const ws = Workbook.addWorksheet(wb, `S${i + 1}`);
    data(ws);
    Cell.setValue(ws, "E6", { formula });
  });
  return wb;
}

/** Whether the engine put anything beside E6 — a spill. */
function spilled(wb: ReturnType<typeof Workbook.create>, i: number): boolean {
  const ws = Workbook.getWorksheet(wb, `S${i + 1}`)!;
  for (let row = 6; row <= 12; row++) {
    for (let col = 5; col <= 10; col++) {
      if ((row !== 6 || col !== 5) && Cell.getValue(ws, row, col) !== null) {
        return true;
      }
    }
  }
  return false;
}

async function xlsxDynamic(bytes: Uint8Array, count: number): Promise<boolean[]> {
  const files = await Archive.unpackZip(bytes);
  return Array.from({ length: count }, (_, i) => {
    const xml = new TextDecoder().decode(files.get(`xl/worksheets/sheet${i + 1}.xml`)!.data);
    return /<c r="E6"[^>]* cm="1"[^>]*><f t="array" ref="E6[:"]/.test(xml);
  });
}

function readDynamic(wb: ReturnType<typeof Workbook.create>, count: number): boolean[] {
  return Array.from({ length: count }, (_, i) => {
    const value = Cell.getValue(Workbook.getWorksheet(wb, `S${i + 1}`)!, "E6") as {
      isDynamicArray?: boolean;
    };
    return value.isDynamicArray === true;
  });
}

async function streamed(formulas: readonly string[], format: "xlsx" | "xlsb"): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  const stream = new WritableStream<Uint8Array>({
    write(chunk) {
      chunks.push(chunk);
    }
  });
  const writer = new Stream.WorkbookWriter({ stream, format });
  formulas.forEach((formula, i) => {
    const ws = writer.addWorksheet(`S${i + 1}`);
    for (const values of [[1, 4], [2, 1], [3, 3], [4, 2], []]) {
      Stream.commitRow(ws.addRow(values));
    }
    Stream.setCellValue(ws.getCell("E6"), { formula });
    ws.commit();
  });
  await writer.commit();
  const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
  let at = 0;
  for (const chunk of chunks) {
    out.set(chunk, at);
    at += chunk.length;
  }
  return out;
}

describe("engine spills ⇔ writers store a dynamic array", () => {
  const all = [...SPILLING, ...ORDINARY, ...DYNAMIC_SCALAR];
  const expectedDynamic = all.map(f => !ORDINARY.includes(f));
  const expectedSpill = all.map(f => SPILLING.includes(f));

  it("the engine spills exactly the formulas that produce an array", () => {
    const wb = book(all);
    calculateFormulas(wb);
    expect(all.map((f, i) => [f, spilled(wb, i)])).toEqual(
      all.map((f, i) => [f, expectedSpill[i]])
    );
  });

  it("the XLSX writer stores exactly the dynamic ones with dynamic-array metadata", async () => {
    const wb = book(all);
    calculateFormulas(wb);
    const dynamic = await xlsxDynamic(await Workbook.toBuffer(wb), all.length);
    expect(all.map((f, i) => [f, dynamic[i]])).toEqual(all.map((f, i) => [f, expectedDynamic[i]]));
  });

  it.each(["xlsx", "xlsb"] as const)("after a %s read-back both still agree", async format => {
    const wb = book(all);
    calculateFormulas(wb);
    const read = await Workbook.read(await Workbook.toBuffer(wb, { format }));
    expect(all.map((f, i) => [f, readDynamic(read, all.length)[i]])).toEqual(
      all.map((f, i) => [f, expectedDynamic[i]])
    );
    calculateFormulas(read);
    expect(all.map((f, i) => [f, spilled(read, i)])).toEqual(
      all.map((f, i) => [f, expectedSpill[i]])
    );
  });

  it.each(["xlsx", "xlsb"] as const)("the streaming %s writer agrees", async format => {
    const read = await Workbook.read(await streamed(all, format));
    expect(all.map((f, i) => [f, readDynamic(read, all.length)[i]])).toEqual(
      all.map((f, i) => [f, expectedDynamic[i]])
    );
  });
});

describe("a formula a file stored as an ordinary one", () => {
  it("is computed with Excel's intersection, not spilled", async () => {
    // Written by hand as Excel 2016 would: no `cm`, no `t="array"`.
    const wb = book(["SEQUENCE(3)"]);
    const bytes = await Workbook.toBuffer(wb);
    const files = await Archive.unpackZip(bytes);
    const sheet = new TextDecoder()
      .decode(files.get("xl/worksheets/sheet1.xml")!.data)
      .replace(/<c r="E6"[^>]*>.*?<\/c>/, '<c r="E6"><f>_xlfn.SEQUENCE(3)</f><v>1</v></c>');
    const entries = new Map([...files].map(([path, entry]) => [path, entry.data]));
    entries.set("xl/worksheets/sheet1.xml", new TextEncoder().encode(sheet));
    const read = await Workbook.read(await Archive.packZip(entries));
    calculateFormulas(read);
    expect(spilled(read, 0)).toBe(false);
    expect(Cell.getResult(Workbook.getWorksheet(read, "S1")!, "E6")).toBe(1);
    // And it is written back as it came.
    const [dynamic] = await xlsxDynamic(await Workbook.toBuffer(read), 1);
    expect(dynamic).toBe(false);
  });
});

describe("isDynamicArray: true without a range", () => {
  it.each(["xlsx", "xlsb"] as const)("is a dynamic array over its own cell in %s", async format => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S1");
    Cell.setValue(ws, "E6", { formula: "1+1", isDynamicArray: true });
    const bytes = await Workbook.toBuffer(wb, { format });
    if (format === "xlsx") {
      const [dynamic] = await xlsxDynamic(bytes, 1);
      expect(dynamic).toBe(true);
    }
    expect(readDynamic(await Workbook.read(bytes), 1)).toEqual([true]);
    expect(
      Cell.getValue(Workbook.getWorksheet(await Workbook.read(bytes), "S1")!, "E6")
    ).toMatchObject({
      shareType: "array",
      ref: "E6"
    });
  });
});
