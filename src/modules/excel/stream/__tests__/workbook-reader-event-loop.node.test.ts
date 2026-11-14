/**
 * The streaming reader hands a worksheet's rows out as it parses them, and lets the event loop run in between, so a
 * server reading an upload keeps answering other requests meanwhile.
 *
 * That used to depend on the source. The reader parsed each chunk synchronously and handed rows out through promises
 * alone, so the event loop ran only when the source happened to wait — on a file read, or on zlib's thread pool. A
 * source that already held its data never waited: a small part inflated in one `inflateRawSync` call, a stored part,
 * a sheet spooled in memory. Each of those read the whole sheet in one task, and no timer could run until the last row
 * was out. Every case below is one of those shapes.
 */

import { packZipSync, unpackZipSync } from "@archive/zip-one-shot";
import { cellGetValue } from "@excel/core/cell";
import { rowGetCell } from "@excel/core/worksheet";
import { Workbook, Worksheet } from "@excel/index";
import { WorkbookReader } from "@excel/stream/workbook-reader";
import { WorkbookReader as BrowserWorkbookReader } from "@excel/stream/workbook-reader.browser";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

// The budget is measured with `performance.now`, so on a fast enough machine a whole read could fit inside it and
// never yield — a test that passes or fails by the speed of the machine. A clock that advances a millisecond per
// reading makes the budget run out by the amount of work done instead: what is checked is that the reader asks
// between its pieces.
beforeEach(() => {
  let clock = 0;
  vi.spyOn(performance, "now").mockImplementation(() => (clock += 1));
});
afterEach(() => {
  vi.restoreAllMocks();
});

const ROWS = 10_000;

const PREREQUISITES = [
  "[Content_Types].xml",
  "_rels/.rels",
  "xl/workbook.xml",
  "xl/_rels/workbook.xml.rels",
  "xl/sharedStrings.xml",
  "xl/styles.xml"
];
const SHEET_FIRST = ["xl/worksheets/sheet1.xml", ...PREREQUISITES];

/**
 * `wb` written as `format` and repacked at `level` with `first` leading the archive: the prerequisites first means the
 * sheet is parsed as it arrives, the sheet first means it is spooled and replayed.
 */
async function repack(
  wb: ReturnType<typeof Workbook.create>,
  {
    first = [],
    level,
    format = "xlsx"
  }: { first?: readonly string[]; level: number; format?: "xlsx" | "xlsb" }
): Promise<Uint8Array> {
  const parts = unpackZipSync(await Workbook.toBuffer(wb, { format }));
  const ordered = new Map<string, Uint8Array>();
  for (const name of [...first, ...parts.keys()]) {
    const part = parts.get(name);
    if (part && !ordered.has(name)) {
      ordered.set(name, part.data);
    }
  }
  return packZipSync(ordered, { level });
}

/** A workbook of a header and {@link ROWS} rows. */
function articles(): ReturnType<typeof Workbook.create> {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "Sheet1");
  Worksheet.addRow(ws, ["Code", "Name", "Material", "Weight", "Supplier", "Country"]);
  for (let i = 0; i < ROWS; i++) {
    Worksheet.addRow(ws, [
      `ART-${i}`,
      `Article ${i} for product line ${i % 97}`,
      ["PET", "Glass", "Paper"][i % 3],
      i / 7,
      `Supplier ${i % 431}`,
      i % 2 ? "DE" : "AT"
    ]);
  }
  return wb;
}

/** Read every row, recording how many turns the event loop had taken when the first and the last arrived. */
async function readCountingTurns(
  reader: AsyncIterable<AsyncIterable<unknown>>
): Promise<{ rows: number; turnsAtFirstRow: number; turnsAtLastRow: number }> {
  // A callback that reschedules itself runs once per turn of the event loop.
  let turns = 0;
  let running = true;
  const tick = (): void => {
    if (running) {
      turns++;
      setImmediate(tick);
    }
  };
  setImmediate(tick);

  let rows = 0;
  let turnsAtFirstRow: number | undefined;
  let turnsAtLastRow = 0;
  try {
    for await (const sheet of reader) {
      for await (const _row of sheet) {
        rows++;
        turnsAtFirstRow ??= turns;
        turnsAtLastRow = turns;
      }
    }
  } finally {
    running = false;
  }
  return { rows, turnsAtFirstRow: turnsAtFirstRow ?? 0, turnsAtLastRow };
}

describe("WorkbookReader and the event loop", () => {
  it.each([
    ["a deflated sheet small enough to inflate in one call", PREREQUISITES, 6],
    ["a stored sheet", PREREQUISITES, 0],
    ["a sheet spooled to a file and replayed", SHEET_FIRST, 6]
  ])("lets the event loop run while it reads %s", async (_name, first, level) => {
    const xlsx = await repack(articles(), { first, level });
    const result = await readCountingTurns(new WorkbookReader(xlsx, { styles: "cache" }));
    expect(result.rows).toBe(ROWS + 1);
    expect(result.turnsAtLastRow).toBeGreaterThan(result.turnsAtFirstRow);
  });

  // The binary package's writer already puts the shared strings, styles and workbook before the sheets.
  it.each([
    ["a deflated binary sheet small enough to inflate in one call", 6],
    ["a stored binary sheet", 0]
  ])("lets the event loop run while it reads %s", async (_name, level) => {
    const xlsb = await repack(articles(), { level, format: "xlsb" });
    const result = await readCountingTurns(new WorkbookReader(xlsb));
    expect(result.rows).toBe(ROWS + 1);
    expect(result.turnsAtLastRow).toBeGreaterThan(result.turnsAtFirstRow);
  });

  it.each([
    ["a stored sheet", PREREQUISITES, 0],
    ["a sheet spooled in memory and replayed", SHEET_FIRST, 0]
  ])("lets the event loop run while the browser reader reads %s", async (_name, first, level) => {
    const xlsx = await repack(articles(), { first, level });
    const result = await readCountingTurns(new BrowserWorkbookReader(xlsx, { styles: "cache" }));
    expect(result.rows).toBe(ROWS + 1);
    expect(result.turnsAtLastRow).toBeGreaterThan(result.turnsAtFirstRow);
  });

  // The parts a sheet needs are read before its first row: a large shared-string table was read in one call, and the
  // event loop waited for all of it before the first row.
  it.each([
    ["the shared strings of a binary package", "xlsb"],
    ["the shared strings of an XML package", "xlsx"]
  ] as const)("lets the event loop run while it reads %s", async (_name, format) => {
    const strings = 50_000;
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    for (let i = 0; i < strings; i++) {
      Worksheet.addRow(ws, [`distinct string number ${i}`]);
    }
    const bytes = await repack(wb, { first: PREREQUISITES, level: 6, format });
    const result = await readCountingTurns(new WorkbookReader(bytes, { styles: "cache" }));
    expect(result.rows).toBe(strings);
    expect(result.turnsAtFirstRow).toBeGreaterThan(0);
  });

  it("reads text the same when its slices cut through multi-byte characters", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    const values: string[] = [];
    // Odd lengths of three- and four-byte characters, so slice boundaries land inside them.
    for (let i = 0; i < 5_000; i++) {
      values.push(`${"数据".repeat(1 + (i % 7))}${"😀".repeat(i % 3)}-${i}`);
      Worksheet.addRow(ws, [values[i]!]);
    }
    const read: unknown[] = [];
    for await (const sheet of new WorkbookReader(
      await repack(wb, { first: PREREQUISITES, level: 0 })
    )) {
      for await (const row of sheet) {
        read.push(cellGetValue(rowGetCell(row, 1)));
      }
    }
    expect(read).toEqual(values);
  });
});
