import { zip } from "@archive/create-archive";
import { extractAll } from "@archive/unzip/extract";
import { cellGetValue } from "@excel/core/cell";
import type { CellData } from "@excel/core/cell";
import { Enums } from "@excel/core/enums";
import { XlsxParseError } from "@excel/errors";
import { Cell, Workbook } from "@excel/index";
import { WorkbookReader } from "@excel/stream/workbook-reader";
import { Readable } from "@stream";
import { describe, expect, it } from "vitest";

/**
 * The full reader and the streaming reader must read the same `<sheetData>` the same way — or both refuse it.
 * Each case is a sheet our writer never produces but other producers, or damaged files, do.
 */

async function withSheetData(sheetData: string): Promise<Uint8Array> {
  const wb = Workbook.create();
  Cell.setValue(Workbook.addWorksheet(wb, "Sheet1"), "A1", 0);
  const archive = zip();
  for (const [name, file] of await extractAll(await Workbook.toBuffer(wb))) {
    if (file.type === "directory") {
      continue;
    }
    archive.add(
      name,
      name === "xl/worksheets/sheet1.xml"
        ? new TextDecoder()
            .decode(file.data)
            .replace(/<sheetData>.*<\/sheetData>|<sheetData\/>/s, sheetData)
        : file.data
    );
  }
  return archive.toBuffer();
}

async function loadWorkbook(buffer: Uint8Array) {
  const wb = Workbook.create();
  await Workbook.read(wb, buffer);
  return wb;
}

async function readFull(buffer: Uint8Array) {
  return Workbook.getWorksheet(await loadWorkbook(buffer), "Sheet1")!;
}

/** Write the workbook holding `buffer`, read that back, and return its sheet. */
async function roundTrip(buffer: Uint8Array) {
  return readFull(await Workbook.toBuffer(await loadWorkbook(buffer)));
}

/** Every cell the stream yields, as `address → value`. */
async function readStream(buffer: Uint8Array): Promise<Record<string, unknown>> {
  const reader = new WorkbookReader(buffer, { worksheets: "emit", sharedStrings: "cache" });
  const cells: Record<string, unknown> = {};
  for await (const ws of reader) {
    for await (const row of ws) {
      for (const cell of (row as { cells: (CellData | undefined)[] }).cells) {
        if (cell) {
          cells[cell.address] = cellGetValue(cell);
        }
      }
    }
  }
  return cells;
}

const sheet = (...rows: string[]) => `<sheetData>${rows.join("")}</sheetData>`;

describe("row and cell references", () => {
  it.each(["x", "12abc", "0", "-1", "1048577", "+", " "])(
    "both readers reject <row r=%j>",
    async r => {
      const buffer = await withSheetData(sheet(`<row r="${r}"><c><v>1</v></c></row>`));
      await expect(readFull(buffer)).rejects.toThrow(`Invalid row number "${r}"`);
      await expect(readStream(buffer)).rejects.toThrow(`Invalid row number "${r}"`);
    }
  );

  it.each(["007", " 7 ", "+7"])(
    "both readers accept <row r=%j>, as xsd:unsignedInt does",
    async r => {
      const buffer = await withSheetData(sheet(`<row r="${r}"><c r="B7"><v>1</v></c></row>`));
      expect(Cell.getValue(await readFull(buffer), "B7")).toBe(1);
      expect(await readStream(buffer)).toEqual({ B7: 1 });
    }
  );

  it("both readers reject a row inferred past the last one", async () => {
    const buffer = await withSheetData(
      sheet('<row r="1048576"><c><v>1</v></c></row><row><c><v>2</v></c></row>')
    );
    await expect(readFull(buffer)).rejects.toThrow("expected 1 to 1048576");
    await expect(readStream(buffer)).rejects.toThrow("expected 1 to 1048576");
  });

  it.each(["A1x", "1A", "A 1"])("both readers reject the cell reference %j", async r => {
    const buffer = await withSheetData(sheet(`<row r="1"><c r="${r}"><v>1</v></c></row>`));
    await expect(readFull(buffer)).rejects.toThrow(r);
    await expect(readStream(buffer)).rejects.toThrow(r);
  });

  it("both readers reject a cell that names another row", async () => {
    const buffer = await withSheetData(sheet('<row r="5"><c r="A7"><v>1</v></c></row>'));
    await expect(readFull(buffer)).rejects.toThrow("cell A7 is written inside row 5");
    await expect(readStream(buffer)).rejects.toThrow("cell A7 is written inside row 5");
  });

  it("both readers place an absolute reference at the cell it names", async () => {
    const buffer = await withSheetData(sheet('<row r="5"><c r="$C$5"><v>7</v></c></row>'));
    expect(Cell.getValue(await readFull(buffer), "C5")).toBe(7);
    expect(await readStream(buffer)).toEqual({ C5: 7 });
  });
});

describe("a cell outside any row", () => {
  // A `<c>` directly inside `<sheetData>`, from a damaged file. Before the first row, `Workbook.read` and
  // `Workbook.readStream` threw an untyped `TypeError` out of `RowXform.parseClose`: the row parser took the stray
  // cell as a child of a row that was not open, and dereferenced the row it did not have. After a row, the same
  // mistake attached the cell to the row that had already closed — while the streaming reader skipped it.
  const cases = {
    "before the first row": sheet(
      '<c r="B1"><v>2</v></c>',
      '<row r="1"><c r="A1"><v>1</v></c></row>',
      '<row r="2"><c r="A2"><v>3</v></c></row>'
    ),
    "between two rows": sheet(
      '<row r="1"><c r="A1"><v>1</v></c></row>',
      '<c r="B1"><v>2</v></c>',
      '<row r="2"><c r="A2"><v>3</v></c></row>'
    )
  };
  const values = (ws: ReturnType<typeof Workbook.addWorksheet>) =>
    ["A1", "B1", "A2"].map(address => Cell.getValue(ws, address));

  it.each(Object.entries(cases))(
    "is skipped by every reader when it is %s",
    async (_where, xml) => {
      const buffer = await withSheetData(xml);
      expect(values(await readFull(buffer))).toEqual([1, null, 3]);
      const streamed = Workbook.create();
      await Workbook.readStream(streamed, Readable.from([buffer]));
      expect(values(Workbook.getWorksheet(streamed, "Sheet1")!)).toEqual([1, null, 3]);
      expect(await readStream(buffer)).toEqual({ A1: 1, A2: 3 });
    }
  );
});

describe("a cell inside a cell", () => {
  // `<c>` is not a child of `<c>`, but a damaged file can nest one. The inner cell ends the outer one, as the start
  // of a sibling would; the outer's surplus `</c>` then closes nothing. Before, the full reader read that surplus close
  // as the row's own and the next one as `<sheetData>`'s, losing A1, B1 and every following row — and the streaming
  // reader wrote the inner cell's value over the outer's, at the outer's address.
  const cases = {
    "after the outer value": sheet(
      '<row r="1"><c r="A1"><v>1</v><c r="X1"><v>4</v></c></c><c r="B1"><v>3</v></c></row>',
      '<row r="2"><c r="A2"><v>2</v></c></row>'
    ),
    "inside rich text": sheet(
      '<row r="1"><c r="A1" t="inlineStr"><is><r><t>a</t></r><r><t>b</t><c r="X1"><v>4</v></c></r></is></c>',
      '<c r="B1"><v>3</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>'
    ),
    "with content after the inner cell": sheet(
      '<row r="1"><c r="A1"><v>1</v><c r="X1"><v>4</v></c><v>9</v></c><c r="B1"><v>3</v></c></row>',
      '<row r="2"><c r="A2"><v>2</v></c></row>'
    )
  };
  const addresses = ["A1", "X1", "B1", "A2"];

  it.each(Object.entries(cases))("is read the same way by every reader, %s", async (where, xml) => {
    const buffer = await withSheetData(xml);
    const a1 = where === "inside rich text" ? "ab" : 1;
    const expected = [a1, 4, 3, 2];
    const plain = (v: unknown) =>
      v && typeof v === "object" && "richText" in v
        ? (v.richText as { text: string }[]).map(r => r.text).join("")
        : v;
    const full = await readFull(buffer);
    expect(addresses.map(a => plain(Cell.getValue(full, a)))).toEqual(expected);
    const streamed = Workbook.create();
    await Workbook.readStream(streamed, Readable.from([buffer]));
    const ws = Workbook.getWorksheet(streamed, "Sheet1")!;
    expect(addresses.map(a => plain(Cell.getValue(ws, a)))).toEqual(expected);
    const cells = await readStream(buffer);
    expect(addresses.map(a => plain(cells[a]))).toEqual(expected);
  });
});

describe("formula cached results", () => {
  it("report a boolean result's effective type as Boolean", async () => {
    const buffer = await withSheetData(
      sheet('<row r="1"><c r="A1" t="b"><f>1=1</f><v>1</v></c></row>')
    );
    expect(Cell.getEffectiveType(await readFull(buffer), "A1")).toBe(Enums.ValueType.Boolean);
  });

  it('reads a t="d" result as a date in both readers, and writes it back as one', async () => {
    const buffer = await withSheetData(
      sheet('<row r="1"><c r="A1" t="d"><f>TODAY()</f><v>2020-01-15T00:00:00</v></c></row>')
    );
    const expected = { formula: "TODAY()", result: new Date(Date.UTC(2020, 0, 15)) };
    const ws = await readFull(buffer);
    expect(Cell.getValue(ws, "A1")).toEqual(expected);
    expect((await readStream(buffer)).A1).toEqual(expected);
    expect(Cell.getValue(await roundTrip(buffer), "A1")).toEqual(expected);
  });
});

describe("cell values that are not what their type says", () => {
  // Read with `parseFloat`, which takes the number at the front of the text, a malformed value became a plausible
  // different one: a price of `49,99` read as `49`.
  it.each(["49,99", "0x10", "49_x002E_99", "12abc", "INF", "NaN", "1e400", " "])(
    "both readers refuse the number value <v>%s</v>, naming the part and the cell",
    async v => {
      const buffer = await withSheetData(sheet(`<row r="2"><c r="A2"><v>${v}</v></c></row>`));
      const message = `xl/worksheets/sheet1.xml: cell A2: <v>${v}</v> is not a number`;
      await expect(readFull(buffer)).rejects.toThrow(XlsxParseError);
      await expect(readFull(buffer)).rejects.toThrow(message);
      await expect(readStream(buffer)).rejects.toThrow(XlsxParseError);
      await expect(readStream(buffer)).rejects.toThrow(message);
    }
  );

  it("both readers refuse a formula's cached number that is not one", async () => {
    const buffer = await withSheetData(
      sheet('<row r="2"><c r="A2"><f>1+1</f><v>2,5</v></c></row>')
    );
    await expect(readFull(buffer)).rejects.toThrow("cell A2: <v>2,5</v> is not a number");
    await expect(readStream(buffer)).rejects.toThrow("cell A2: <v>2,5</v> is not a number");
  });

  it.each([
    [" 12 ", 12],
    ["+.5", 0.5],
    ["5.", 5],
    ["-1.5E3", -1500],
    ["1e-400", 0]
  ])("both readers read the xsd:double <v>%s</v> as %d", async (v, expected) => {
    const buffer = await withSheetData(sheet(`<row r="2"><c r="A2"><v>${v}</v></c></row>`));
    expect(Cell.getValue(await readFull(buffer), "A2")).toBe(expected);
    expect(await readStream(buffer)).toEqual({ A2: expected });
  });

  it("both readers refuse a shared string index that is not an integer", async () => {
    const buffer = await withSheetData(sheet('<row r="2"><c r="A2" t="s"><v>0abc</v></c></row>'));
    await expect(readFull(buffer)).rejects.toThrow(
      "cell A2: <v>0abc</v> is not a shared string index"
    );
    await expect(readStream(buffer)).rejects.toThrow(
      "cell A2: <v>0abc</v> is not a shared string index"
    );
  });

  it("both readers refuse a negative shared string index, and one past the shared strings", async () => {
    const negative = await withSheetData(sheet('<row r="2"><c r="A2" t="s"><v>-1</v></c></row>'));
    await expect(readFull(negative)).rejects.toThrow(
      "cell A2: <v>-1</v> is not a shared string index"
    );
    await expect(readStream(negative)).rejects.toThrow(
      "cell A2: <v>-1</v> is not a shared string index"
    );

    const wb = Workbook.create();
    Cell.setValue(Workbook.addWorksheet(wb, "Sheet1"), "A1", "only");
    const archive = zip();
    for (const [name, file] of await extractAll(await Workbook.toBuffer(wb))) {
      if (file.type !== "directory") {
        archive.add(
          name,
          name === "xl/worksheets/sheet1.xml"
            ? new TextDecoder()
                .decode(file.data)
                .replace(
                  /<sheetData>.*<\/sheetData>/s,
                  sheet('<row r="2"><c r="A2" t="s"><v>5</v></c></row>')
                )
            : file.data
        );
      }
    }
    const past = await archive.toBuffer();
    const message = "xl/worksheets/sheet1.xml: cell A2: <v>5</v> names a shared string";
    await expect(readFull(past)).rejects.toThrow(XlsxParseError);
    await expect(readFull(past)).rejects.toThrow(message);
    await expect(readStream(past)).rejects.toThrow(message);
  });

  // A cell may omit `r`, and follows the previous one; a refusal of its content named `cell undefined`.
  it("both readers name a cell with no r by where it is", async () => {
    const buffer = await withSheetData(sheet('<row r="2"><c><v>1</v></c><c><v>x</v></c></row>'));
    await expect(readFull(buffer)).rejects.toThrow("cell B2: <v>x</v> is not a number");
    await expect(readStream(buffer)).rejects.toThrow("cell B2: <v>x</v> is not a number");
  });

  // A cell's text may arrive in pieces, which are joined; a second element was joined the same way.
  it.each([
    ["<v>", '<c r="A3"><v>12.5</v><v>99</v></c>'],
    ["<f>", '<c r="A3"><f>1</f><f>+1</f><v>2</v></c>'],
    ["<is>", '<c r="A3" t="inlineStr"><is><t>a</t></is><is><t>b</t></is></c>']
  ])("both readers refuse a cell with more than one %s", async (element, cell) => {
    const buffer = await withSheetData(sheet(`<row r="3">${cell}</row>`));
    await expect(readFull(buffer)).rejects.toThrow(`cell A3: has more than one ${element}`);
    await expect(readStream(buffer)).rejects.toThrow(`cell A3: has more than one ${element}`);
  });

  it.each([
    ["1", true],
    ["0", false],
    ["true", true],
    ["false", false]
  ])("both readers read the boolean <v>%s</v> as %s", async (v, expected) => {
    const buffer = await withSheetData(sheet(`<row r="2"><c r="A2" t="b"><v>${v}</v></c></row>`));
    expect(Cell.getValue(await readFull(buffer), "A2")).toBe(expected);
    expect(await readStream(buffer)).toEqual({ A2: expected });
  });

  // Read as "not zero", `false` was `true`, and so was any text.
  it.each([
    ["a boolean", '<c r="A2" t="b"><v>2x</v></c>', "<v>2x</v> is not a boolean"],
    ["a cached boolean", '<c r="A2" t="b"><f>1=1</f><v>yes</v></c>', "<v>yes</v> is not a boolean"],
    ["a date", '<c r="A2" t="d"><v>soon</v></c>', "<v>soon</v> is not a date"]
  ])("both readers refuse %s that is not one", async (_what, cell, detail) => {
    const buffer = await withSheetData(sheet(`<row r="2">${cell}</row>`));
    await expect(readFull(buffer)).rejects.toThrow(`cell A2: ${detail}`);
    await expect(readStream(buffer)).rejects.toThrow(`cell A2: ${detail}`);
  });
});

describe("inline strings", () => {
  // `<is>` is read as the shared strings' `<si>` is. Its text used to be collected from every `<t>` below the cell.
  it.each([
    [
      "a phonetic run, which is not part of the text",
      '<c r="A1" t="inlineStr"><is><t>漢字</t><rPh sb="0" eb="2"><t>かんじ</t></rPh><phoneticPr fontId="1"/></is></c>',
      "漢字"
    ],
    [
      "a <t> before runs, as the first run",
      '<c r="A1" t="inlineStr"><is><t>a</t><r><t>b</t></r></is></c>',
      { richText: [{ text: "a" }, { text: "b" }] }
    ],
    [
      "<is> rather than <v> in an inline string",
      '<c r="A1" t="inlineStr"><v>x</v><is><t>y</t></is></c>',
      "y"
    ],
    ["an escape once", '<c r="A1" t="inlineStr"><is><t>_x005F_x0041_</t></is></c>', "_x0041_"],
    ["an empty <is> as the empty string", '<c r="A1" t="inlineStr"><is/></c>', ""]
  ])("both readers read %s", async (_what, cell, expected) => {
    const buffer = await withSheetData(sheet(`<row r="1">${cell}<c r="B1"><v>2</v></c></row>`));
    const ws = await readFull(buffer);
    expect(Cell.getValue(ws, "A1")).toEqual(expected);
    expect(Cell.getValue(ws, "B1")).toBe(2);
    expect(await readStream(buffer)).toEqual({ A1: expected, B1: 2 });
  });

  it.each([
    ["more than one <t>", "<t>a</t><t>b</t>"],
    ["a <t> after an <r>", "<r><t>a</t></r><t>b</t>"],
    ["a run with more than one <t>", "<r><t>a</t><t>b</t></r>"]
  ])("both readers refuse an <is> with %s", async (detail, content) => {
    const buffer = await withSheetData(
      sheet(`<row r="1"><c r="A1" t="inlineStr"><is>${content}</is></c></row>`)
    );
    await expect(readFull(buffer)).rejects.toThrow(`cell A1: its <is> has ${detail}`);
    await expect(readStream(buffer)).rejects.toThrow(`cell A1: its <is> has ${detail}`);
  });

  // An `<extLst>` (or a foreign element) in a cell ended the cell at its closing tag in the full reader.
  it("both readers skip an element of a cell they do not model", async () => {
    const buffer = await withSheetData(
      sheet(
        '<row r="1"><c r="A1"><v>5</v><extLst><ext uri="x"><v>9</v></ext></extLst></c><c r="B1"><v>2</v></c></row>'
      )
    );
    const ws = await readFull(buffer);
    expect([Cell.getValue(ws, "A1"), Cell.getValue(ws, "B1")]).toEqual([5, 2]);
    expect(await readStream(buffer)).toEqual({ A1: 5, B1: 2 });
  });
});

describe("shared formulas without a usable master", () => {
  it("keep the cached result of a dependent whose master never appears, and write it back", async () => {
    const buffer = await withSheetData(
      sheet('<row r="1"><c r="A1"><f t="shared" si="3"/><v>9</v></c></row>')
    );
    const ws = await readFull(buffer);
    expect(Cell.getValue(ws, "A1")).toBe(9);
    expect(await readStream(buffer)).toEqual({ A1: 9 });
    expect(Cell.getValue(await roundTrip(buffer), "A1")).toBe(9);
  });

  it("give a dependent before its master the master's formula when loading, and refuse when streaming", async () => {
    const buffer = await withSheetData(
      sheet(
        '<row r="1"><c r="A1"><f t="shared" si="0"/><v>18</v></c></row>',
        '<row r="2"><c r="A2"><f t="shared" ref="A1:A2" si="0">B2*2</f><v>2</v></c></row>'
      )
    );
    const ws = await readFull(buffer);
    expect(Cell.getValue(ws, "A1")).toEqual({ formula: "B1*2", result: 18 });
    expect(Cell.getValue(await roundTrip(buffer), "A1")).toEqual({ formula: "B1*2", result: 18 });

    await expect(readStream(buffer)).rejects.toThrow("comes before its master A2");
  });

  it("leave strings and quoted sheet names alone when sliding", async () => {
    const buffer = await withSheetData(
      sheet(
        `<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">IF(B1="B1",'Q1 C1'!C1,"x""D1")</f><v>1</v></c></row>`,
        '<row r="2"><c r="A2"><f t="shared" si="0"/><v>2</v></c></row>'
      )
    );
    const expected = `IF(B2="B1",'Q1 C1'!C2,"x""D1")`;
    expect(Cell.getFormula(await readFull(buffer), "A2")).toBe(expected);
    expect((await readStream(buffer)).A2).toMatchObject({ formula: expected });
  });

  it("slide the formula to a dependent outside the declared range, as Excel computes it", async () => {
    const buffer = await withSheetData(
      sheet(
        '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1+1</f><v>1</v></c></row>',
        '<row r="9"><c r="A9"><f t="shared" si="0"/><v>3</v></c></row>'
      )
    );
    expect(Cell.getFormula(await readFull(buffer), "A9")).toBe("B9+1");
    expect((await readStream(buffer)).A9).toMatchObject({ formula: "B9+1", result: 3 });
  });
});
