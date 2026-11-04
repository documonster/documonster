/**
 * Every capability the deprecated handle API offered has a position-based,
 * non-deprecated form — and the wrappers reproduce the old behaviour exactly.
 *
 * The `@excel/core/*` imports are deliberate: they are the handle-based functions the deprecated API
 * was built on, used here as the baseline each position form is compared against. Everything else
 * goes through the public namespaces, and workbooks carrying parts the library does not model are
 * read from a real package (`readWithParts`) rather than assembled by writing internal fields.
 */
import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { expectValidXlsx } from "@excel/__tests__/helpers/expect-valid-xlsx";
import { entryText } from "@excel/__tests__/helpers/zip-text";
import type { AddChartOptions } from "@excel/chart/model/types";
import { cellView } from "@excel/core/cell";
import {
  rowActualCellCount,
  rowCellCount,
  rowDimensions,
  rowGetModel,
  rowHasValues
} from "@excel/core/row";
import { insertRow, insertRows } from "@excel/core/worksheet";
import { addRow, eachRow, findRow, getRow } from "@excel/core/worksheet-core";
import type { WorksheetData } from "@excel/core/worksheet-core";
import { PivotTableError } from "@excel/errors";
import {
  Cell,
  Chart,
  Column,
  Form,
  Image,
  Pivot,
  Row,
  Table,
  Watermark,
  Workbook
} from "@excel/index";
import { describe, expect, it } from "vitest";

const PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

function sheet() {
  const wb = Workbook.create();
  return { wb, ws: Workbook.addWorksheet(wb, "S") };
}

function styled(ws: WorksheetData): void {
  Row.add(ws, ["a", , "c"]);
  Row.setFont(ws, 1, { bold: true });
  Row.setHeight(ws, 1, 30);
  Cell.setFill(ws, "B1", { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } });
  Row.add(ws, ["x", "y"]);
  Row.setFont(ws, 2, { italic: true });
}

function snapshot(ws: WorksheetData): unknown {
  const rows: unknown[] = [];
  for (let r = 1; r <= Row.getCount(ws); r++) {
    rows.push(Row.getModel(ws, r) ?? null);
  }
  return rows;
}

describe("Cell.getView — the position form of find + view", () => {
  it("reads the same fields as the handle view, and creates nothing", () => {
    const { ws } = sheet();
    Cell.setValue(ws, "B2", 1.5);
    Cell.setNumFmt(ws, "B2", "0.00");
    Cell.setFont(ws, "B2", { bold: true });
    const view = Cell.getView(ws, "B2")!;
    const handle = cellView(getRow(ws, 2).cells[1]);
    expect(view.value).toBe(handle.value);
    expect(view.text).toBe(handle.text);
    expect(view.numFmt).toBe("0.00");
    expect(view.effectiveType).toBe(handle.effectiveType);
    expect(view.font?.bold).toBe(true);
    expect(view.date1904).toBe(false);
    expect(Cell.getView(ws, 2, 2)?.value).toBe(1.5);

    expect(Cell.getView(ws, "Z9")).toBeUndefined();
    expect(Cell.has(ws, "Z9")).toBe(false);
    expect(Row.has(ws, 9)).toBe(false);
  });

  it("is live", () => {
    const { ws } = sheet();
    Cell.setValue(ws, "A1", 1);
    const view = Cell.getView(ws, "A1")!;
    Cell.setValue(ws, "A1", 2);
    expect(view.value).toBe(2);
  });
});

describe("Row — what a RowData handle exposed, by number", () => {
  it("hasValues / cell counts / dimensions / model / collapsed match the handle", () => {
    const { ws } = sheet();
    Cell.setValue(ws, "C3", 1);
    Cell.setValue(ws, "E3", 2);
    Cell.setStyle(ws, "G3", { font: { bold: true } });
    const r = findRow(ws, 3)!;
    expect(Row.hasValues(ws, 3)).toBe(rowHasValues(r));
    expect(Row.getCellCount(ws, 3)).toBe(rowCellCount(r));
    expect(Row.getActualCellCount(ws, 3)).toBe(rowActualCellCount(r));
    expect(Row.getDimensions(ws, 3)).toEqual(rowDimensions(r));
    expect(Row.getModel(ws, 3)).toEqual(rowGetModel(r));
    expect(Row.getCellCount(ws, 3)).toBe(7);
    expect(Row.getActualCellCount(ws, 3)).toBe(2);
    expect(Row.getDimensions(ws, 3)).toEqual({ min: 3, max: 5 });
    expect(Row.isCollapsed(ws, 3)).toBe(false);
  });

  it("answers for absent rows without creating them", () => {
    const { ws } = sheet();
    expect(Row.hasValues(ws, 5)).toBe(false);
    expect(Row.getCellCount(ws, 5)).toBe(0);
    expect(Row.getActualCellCount(ws, 5)).toBe(0);
    expect(Row.getDimensions(ws, 5)).toBeUndefined();
    expect(Row.getModel(ws, 5)).toBeUndefined();
    expect(Row.isCollapsed(ws, 5)).toBe(false);
    expect(Row.has(ws, 5)).toBe(false);
  });

  it("Row.each visits exactly the rows eachRow visited", () => {
    const { ws } = sheet();
    Cell.setValue(ws, "A2", 1);
    Cell.setStyle(ws, "A3", { font: { bold: true } });
    Cell.setValue(ws, "A5", 1);
    const before: number[] = [];
    eachRow(ws, (_r, n) => before.push(n));
    const after: number[] = [];
    Row.each(ws, n => after.push(n));
    expect(after).toEqual(before);
    expect(after).toEqual([2, 5]);
  });
});

describe("Row.add / insert reproduce every legacy style code", () => {
  const addCases: Array<[string, Row.AddOptions | undefined]> = [
    ["n", undefined],
    ["i", { copyStyle: "above" }],
    ["i+", { copyStyle: "above", includeEmptyCells: true }]
  ];
  for (const [code, options] of addCases) {
    it(`addRow "${code}"`, () => {
      const a = sheet().ws;
      const b = sheet().ws;
      styled(a);
      styled(b);
      addRow(a, ["new"], code);
      Row.add(b, ["new"], options);
      expect(snapshot(b)).toEqual(snapshot(a));
    });
  }

  it('addRow "o" / "o+" behave as "n" — and so does Row.add without options', () => {
    const a = sheet().ws;
    const b = sheet().ws;
    styled(a);
    styled(b);
    addRow(a, ["new"], "o+");
    Row.add(b, ["new"]);
    expect(snapshot(b)).toEqual(snapshot(a));
  });

  const insertCases: Array<[string, Row.InsertOptions | undefined]> = [
    ["n", undefined],
    ["i", { copyStyle: "above" }],
    ["i+", { copyStyle: "above", includeEmptyCells: true }],
    ["o", { copyStyle: "below" }],
    ["o+", { copyStyle: "below", includeEmptyCells: true }]
  ];
  for (const [code, options] of insertCases) {
    it(`insertRow "${code}"`, () => {
      const a = sheet().ws;
      const b = sheet().ws;
      styled(a);
      styled(b);
      insertRow(a, 2, ["new"], code);
      expect(Row.insert(b, 2, ["new"], options)).toBe(2);
      expect(snapshot(b)).toEqual(snapshot(a));
    });
    it(`insertRows "${code}"`, () => {
      const a = sheet().ws;
      const b = sheet().ws;
      styled(a);
      styled(b);
      insertRows(a, 2, [["p"], ["q"]], code);
      expect(Row.insertMany(b, 2, [["p"], ["q"]], options)).toEqual([2, 3]);
      expect(snapshot(b)).toEqual(snapshot(a));
    });
  }
});

describe("Column — every ColumnView field by position", () => {
  it("number, letter, header, key, width, hidden, outline, style, bestFit", () => {
    const { ws } = sheet();
    Column.setDefinitions(ws, [
      { header: "Id", key: "id", width: 8 },
      { header: "Name", key: "name", width: 20, hidden: true, outlineLevel: 1, bestFit: true }
    ]);
    Column.setFont(ws, "name", { bold: true });
    expect(Column.getNumber(ws, "name")).toBe(2);
    expect(Column.getLetter(ws, "name")).toBe("B");
    expect(Column.getHeader(ws, "name")).toBe("Name");
    expect(Column.getKey(ws, 2)).toBe("name");
    expect(Column.getWidth(ws, "B")).toBe(20);
    expect(Column.isHidden(ws, "B")).toBe(true);
    expect(Column.getOutlineLevel(ws, "B")).toBe(1);
    expect(Column.getStyle(ws, "B").font?.bold).toBe(true);
    expect(Column.getDefinition(ws, "B").bestFit).toBe(true);
    expect(Column.getDefinitions(ws)).toHaveLength(2);
  });
});

describe("Watermark.get", () => {
  it("answers null without a watermark and the settings with one", () => {
    const { wb, ws } = sheet();
    expect(Watermark.get(ws)).toBeNull();
    const imageId = Image.add(wb, { base64: PNG, extension: "png" });
    Watermark.add(ws, { imageId });
    expect(Watermark.get(ws)?.imageId).toBe(String(imageId));
  });
});

/**
 * `wb` written, with `parts` added to the package, and read back — how a test gets a workbook that
 * carries parts this library preserves without modelling (a slicer cache, a chart it cannot parse).
 * Built from the library's own output rather than checked in, so the fixture cannot drift from it.
 */
async function readWithParts(
  wb: Workbook.Handle,
  parts: Record<string, string>
): Promise<Workbook.Handle> {
  const entries = await extractAll(await Workbook.toBuffer(wb));
  const encoder = new TextEncoder();
  const bytes = await createZip([
    ...[...entries].map(([name, file]) => ({ name, data: file.data })),
    ...Object.entries(parts).map(([name, xml]) => ({ name, data: encoder.encode(xml) }))
  ]);
  const back = Workbook.create();
  await Workbook.read(back, bytes);
  return back;
}

async function roundTrip(wb: Workbook.Handle, label: string): Promise<Workbook.Handle> {
  const buffer = await Workbook.toBuffer(wb);
  await expectValidXlsx(buffer, { label });
  const back = Workbook.create();
  await Workbook.read(back, buffer);
  return back;
}

describe("Image.remove", () => {
  it("removes a placed image by handle or index, and round-trips", async () => {
    const { wb, ws } = sheet();
    const imageId = Image.add(wb, { base64: PNG, extension: "png" });
    Image.place(ws, imageId, "A1:B2");
    Image.place(ws, imageId, "D1:E2");
    Image.place(ws, imageId, "G1:H2");
    const [first] = Image.list(ws);
    expect(Image.remove(ws, first)).toBe(true);
    expect(Image.remove(ws, first)).toBe(false);
    expect(Image.remove(ws, 5)).toBe(false);
    expect(Image.remove(ws, 1)).toBe(true);
    expect(Image.list(ws)).toHaveLength(1);

    const back = await roundTrip(wb, "image removed");
    const images = Image.list(Workbook.getWorksheet(back, "S")!);
    expect(images).toHaveLength(1);
  });

  it("removes a sheet's last image cleanly", async () => {
    const { wb, ws } = sheet();
    const imageId = Image.add(wb, { base64: PNG, extension: "png" });
    Image.place(ws, imageId, "A1:B2");
    expect(Image.remove(ws, 0)).toBe(true);
    const back = await roundTrip(wb, "last image removed");
    expect(Image.list(Workbook.getWorksheet(back, "S")!)).toHaveLength(0);
  });

  it("does not touch a watermark", () => {
    const { wb, ws } = sheet();
    const imageId = Image.add(wb, { base64: PNG, extension: "png" });
    Watermark.add(ws, { imageId });
    expect(Image.list(ws)).toHaveLength(0);
    expect(Image.remove(ws, 0)).toBe(false);
    expect(Watermark.get(ws)).not.toBeNull();
  });
});

const PIVOT_CHART: AddChartOptions = {
  type: "bar",
  barDir: "col",
  series: [{ name: "Revenue", categories: "Sales[Region]", values: "Sales[Revenue]" }]
};

function pivotBook() {
  const wb = Workbook.create();
  const data = Workbook.addWorksheet(wb, "Data");
  Row.addMany(data, [
    ["Region", "Revenue"],
    ["East", 10],
    ["West", 20],
    ["East", 5]
  ]);
  const table = Table.add(data, {
    name: "Sales",
    ref: "A1",
    columns: [{ name: "Region" }, { name: "Revenue" }],
    rows: [
      ["East", 10],
      ["West", 20],
      ["East", 5]
    ]
  });
  const p1 = Workbook.addWorksheet(wb, "P1");
  const p2 = Workbook.addWorksheet(wb, "P2");
  const a = Pivot.add(p1, { sourceTable: table, rows: ["Region"], values: ["Revenue"] });
  const b = Pivot.add(p2, { sourceTable: table, rows: ["Region"], values: ["Revenue"] });
  return { wb, p1, p2, a, b };
}

describe("Pivot.remove", () => {
  it("removes by handle or index, and round-trips", async () => {
    const { wb, p1, p2, a } = pivotBook();
    expect(Pivot.remove(p2, a)).toBe(false);
    expect(Pivot.remove(p1, a)).toBe(true);
    expect(Pivot.remove(p1, a)).toBe(false);
    expect(Pivot.list(p1)).toHaveLength(0);
    expect(Pivot.list(p2)).toHaveLength(1);

    const back = await roundTrip(wb, "pivot removed");
    expect(Pivot.list(Workbook.getWorksheet(back, "P1")!)).toHaveLength(0);
    expect(Pivot.list(Workbook.getWorksheet(back, "P2")!)).toHaveLength(1);

    expect(Pivot.remove(p2, 0)).toBe(true);
    const empty = await roundTrip(wb, "all pivots removed");
    expect(Pivot.list(Workbook.getWorksheet(empty, "P2")!)).toHaveLength(0);
  });

  it("takes the removed table's cache with it, and keeps the others'", async () => {
    const { wb, p1 } = pivotBook();
    const cacheParts = async (book: Workbook.Handle) =>
      [...(await extractAll(await Workbook.toBuffer(book))).keys()].filter(name =>
        /pivotCacheDefinition\d+\.xml$/.test(name)
      );
    expect(await cacheParts(wb)).toHaveLength(2);
    expect(Pivot.remove(p1, 0)).toBe(true);
    expect(await cacheParts(wb)).toHaveLength(1);
    await roundTrip(wb, "one cache left");
  });

  it("matches a chart's pivot reference case-insensitively, with or without quoting", async () => {
    for (const name of [
      `<c:name>[Book1.xlsx]'p1'!pivottable1</c:name><c:fmtId val="0"/>`,
      `<c:name>'[Book1.xlsx]P1'!PivotTable1</c:name>`
    ]) {
      const back = await readWithParts(pivotBook().wb, {
        "xl/charts/chart9.xml": `<c:chartSpace xmlns:c="x"><c:pivotSource>${name}</c:pivotSource></c:chartSpace>`
      });
      expect(() => Pivot.remove(Workbook.getWorksheet(back, "P1")!, 0)).toThrow(PivotTableError);
    }
  });

  it("refuses while a slicer or timeline of a loaded workbook filters the table", async () => {
    const { wb, p1: source } = pivotBook();
    const [target] = Pivot.list(source);
    const name = target.name ?? `PivotTable${target.tableNumber}`;
    const tabId = source.id;
    const filters: Record<string, string>[] = [
      {
        "xl/slicerCaches/slicerCache1.xml": `<slicerCacheDefinition xmlns="x" name="Slicer_Region"><pivotTables><pivotTable tabId="${tabId}" name="${name}"/></pivotTables></slicerCacheDefinition>`
      },
      {
        "xl/timelineCaches/timelineCache1.xml": `<x:timelineCacheDefinition xmlns:x="x"><x:pivotTables><x:pivotTable tabId='${tabId}' name='${name}'/></x:pivotTables></x:timelineCacheDefinition>`
      }
    ];
    for (const parts of filters) {
      const back = await readWithParts(wb, parts);
      const p1 = Workbook.getWorksheet(back, "P1")!;
      expect(() => Pivot.remove(p1, 0)).toThrow(PivotTableError);
      // Same name on another sheet is not this table.
      expect(Pivot.remove(Workbook.getWorksheet(back, "P2")!, 0)).toBe(true);
    }
    const plain = await roundTrip(wb, "pivots");
    expect(Pivot.remove(Workbook.getWorksheet(plain, "P1")!, 0)).toBe(true);
  });

  it("refuses while a pivot chart is bound to the table, and allows it once the chart is gone", async () => {
    const { wb, p1, a } = pivotBook();
    Chart.addPivot(p1, a, PIVOT_CHART, "E1:L12");
    const [chart] = Chart.list(p1);
    expect(() => Pivot.remove(p1, a)).toThrow(PivotTableError);
    expect(Pivot.list(p1)).toHaveLength(1);
    expect(Chart.remove(p1, chart)).toBe(true);
    expect(Pivot.remove(p1, a)).toBe(true);
    await roundTrip(wb, "pivot removed after its chart");
  });

  it("refuses when a chart part read from a file names the pivot", async () => {
    // The reference carries the authoring workbook's name, as Excel writes it, and is matched in the
    // part's bytes as well as in the parsed model.
    const back = await readWithParts(pivotBook().wb, {
      "xl/charts/chart9.xml": `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:pivotSource><c:name>[Book1.xlsx]P2!PivotTable2</c:name><c:fmtId val="0"/></c:pivotSource></c:chartSpace>`
    });
    // Bound to the other pivot table: this one may go.
    expect(Pivot.remove(Workbook.getWorksheet(back, "P1")!, 0)).toBe(true);
    const other = Workbook.getWorksheet(back, "P2")!;
    expect(() => Pivot.remove(other, 0)).toThrow(PivotTableError);
    expect(Pivot.list(other)).toHaveLength(1);
  });

  it("refuses on a loaded workbook whose chart names the pivot", async () => {
    const { wb, p1, a } = pivotBook();
    Chart.addPivot(p1, a, PIVOT_CHART, "E1:L12");
    const back = await roundTrip(wb, "pivot chart");
    const loaded = Workbook.getWorksheet(back, "P1")!;
    expect(() => Pivot.remove(loaded, 0)).toThrow(PivotTableError);
  });
});

describe("Chart.place — add returns what it attached", () => {
  const HISTOGRAM = { type: "histogram" as const, series: [{ values: "Data!B2:B4" }] };

  it("returns the handle of each chart family, which remove takes", () => {
    const { wb, a } = pivotBook();
    const ws = Workbook.addWorksheet(wb, "C");
    const other = Workbook.addWorksheet(wb, "D");
    Chart.add(other, PIVOT_CHART, "A1:D8"); // takes classic number 1 on another sheet
    const classic = Chart.place(ws, "A1:D8", PIVOT_CHART);
    const ex = Chart.place(ws, "F1:J8", HISTOGRAM);
    const combo = Chart.place(ws, "A10:D18", {
      groups: [
        { type: "bar", barDir: "col", series: [{ values: "Data!B2:B4" }] },
        { type: "line", series: [{ values: "Data!B2:B4" }] }
      ]
    });
    const pivot = Chart.place(ws, "F10:J18", { ...PIVOT_CHART, pivotTable: a });
    expect(Chart.list(ws)).toEqual([classic, ex, combo, pivot]);
    expect([classic.chartNumber, ex.chartNumber, ex.chartExNumber]).toEqual([2, 0, 1]);
    expect(Chart.isEx(ex)).toBe(true);
    expect(Chart.getModel(pivot)?.pivotSource).toBeDefined();

    expect(Chart.remove(ws, ex)).toBe(true);
    expect(Chart.remove(ws, ex)).toBe(false);
    expect(Chart.list(ws)).toEqual([classic, combo, pivot]);
    expect(Chart.list(other)).toHaveLength(1);
  });

  it("removing by an add* part number is a position — the trap place avoids", () => {
    const { wb } = pivotBook();
    const ws = Workbook.addWorksheet(wb, "C");
    const other = Workbook.addWorksheet(wb, "D");
    Chart.add(other, PIVOT_CHART, "A1:D8");
    const first = Chart.place(ws, "A1:D8", PIVOT_CHART);
    const n = Chart.add(ws, PIVOT_CHART, "F1:J8"); // part number 3, position 1
    expect(n).toBe(3);
    const handle = Chart.list(ws).find(c => c.chartNumber === n)!;
    expect(Chart.remove(ws, n)).toBe(false); // no chart at position 3
    expect(Chart.remove(ws, handle)).toBe(true);
    expect(Chart.list(ws)).toEqual([first]);
  });
});

describe("Form.remove", () => {
  it("Form.list returns a copy; the deprecated listCheckboxes still hands out the live array", () => {
    const { ws } = sheet();
    Form.addCheckbox(ws, "B2:C2", { text: "A" });
    Form.list(ws).length = 0;
    expect(Form.list(ws)).toHaveLength(1);
    Form.listCheckboxes(ws).splice(0, 1);
    expect(Form.list(ws)).toEqual([]);
  });

  it("removes by handle or index, never reissues a live shape id, and round-trips", async () => {
    const { wb, ws } = sheet();
    const a = Form.addCheckbox(ws, "B2:C2", { text: "A" });
    const b = Form.addCheckbox(ws, "B3:C3", { text: "B" });
    expect(Form.remove(ws, a)).toBe(true);
    expect(Form.remove(ws, a)).toBe(false);
    expect(Form.remove(ws, 9)).toBe(false);
    const c = Form.addCheckbox(ws, "B4:C4", { text: "C" });
    expect(c.model.shapeId).not.toBe(b.model.shapeId);
    expect(c.model.shapeId).toBeGreaterThan(b.model.shapeId);

    expect(Form.list(ws).map(cb => Form.getText(cb))).toEqual(["B", "C"]);

    const buffer = await Workbook.toBuffer(wb);
    await expectValidXlsx(buffer, { label: "checkbox removed" });
    const entries = await extractAll(buffer);
    const sheetXml = entryText(entries, "xl/worksheets/sheet1.xml")!;
    const controlIds = [...sheetXml.matchAll(/<control shapeId="(\d+)"/g)].map(m => Number(m[1]));
    expect(controlIds).toEqual([b.model.shapeId, c.model.shapeId]);
    expect([...entries.keys()].filter(name => name.includes("ctrlProps/"))).toHaveLength(2);
  });

  it("notes and checkboxes on one sheet never share a VML shape id", async () => {
    const { wb, ws } = sheet();
    Form.addCheckbox(ws, "B2:C2", { text: "A" });
    Form.addCheckbox(ws, "B3:C3", { text: "B" });
    Cell.setNote(ws, "E5", "one");
    Cell.setNote(ws, "E6", "two");
    const buffer = await Workbook.toBuffer(wb);
    await expectValidXlsx(buffer, { label: "notes + checkboxes" });
    const entries = await extractAll(buffer);
    const path = [...entries.keys()].find(name => /vmlDrawing\d+\.vml$/.test(name))!;
    const ids = [...entryText(entries, path)!.matchAll(/<v:shape [^>]*?id="([^"]+)"/g)].map(
      m => m[1]
    );
    expect(ids).toHaveLength(4);
    expect(new Set(ids).size).toBe(4);
  });
});
