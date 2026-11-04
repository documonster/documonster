/**
 * Sheet-level features a streamed XLSB writes in the sheet part's epilogue, against the buffered writer's output for the
 * same workbook.
 *
 * The streamed writer closes a sheet part before the package is assembled, and used to write none of conditional
 * formatting, hyperlinks, sheet protection, the auto filter or ignored errors — reporting nothing. Each now goes through
 * the record writers the buffered path uses, so the records are compared byte for byte, and read back.
 */
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { extractAll } from "@archive/unzip/extract";
import { Cell, ConditionalFormatting, Stream, Workbook, Worksheet } from "@excel/index";
import { iterateInterpretableRecords } from "@excel/xlsb/binary";
import { recordSpec } from "@excel/xlsb/spec/records";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

let dir: string;

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "documonster-sxsf-"));
});

afterEach(async () => {
  await rm(dir, { recursive: true, force: true });
});

const RULES = [
  { type: "expression", formulae: ["B1>2"], style: { font: { bold: true } } },
  {
    type: "cellIs",
    operator: "between",
    formulae: ["1", "4"],
    style: { fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FF00FF00" } } }
  },
  { type: "dataBar", cfvo: [{ type: "min" }, { type: "max" }], color: { argb: "FF638EC6" } },
  {
    type: "colorScale",
    cfvo: [{ type: "min" }, { type: "max" }],
    color: [{ argb: "FFFF0000" }, { argb: "FF00FF00" }]
  },
  {
    type: "iconSet",
    iconSet: "3Arrows",
    cfvo: [
      { type: "percent", value: 0 },
      { type: "percent", value: 33 },
      { type: "percent", value: 67 }
    ]
  },
  { type: "top10", rank: 2, style: { font: { italic: true } } },
  {
    type: "containsText",
    operator: "containsText",
    text: "x",
    style: { font: { bold: true } }
  }
];
const SECOND = [{ type: "expression", formulae: ["C1<0"], style: { font: { bold: true } } }];

const OPTIONS = { autoFilter: "A1:C5" } as const;
const PROTECTION = { selectLockedCells: true, formatCells: true } as const;
const IGNORED = { ref: "A1:A3", numberStoredAsText: true };

async function buffered(): Promise<Uint8Array> {
  const workbook = Workbook.create();
  const sheet = Workbook.addWorksheet(workbook, "S", OPTIONS as never);
  Cell.setValue(sheet, "A1", { text: "link", hyperlink: "https://example.com/" } as never);
  Cell.setValue(sheet, "A2", { text: "in", hyperlink: "#S!B2" } as never);
  Cell.setValue(sheet, "B1", 5);
  ConditionalFormatting.add(sheet, { ref: "B1:B5 C1:C5", rules: RULES } as never);
  ConditionalFormatting.add(sheet, { ref: "D1:D3", rules: SECOND } as never);
  await Worksheet.protect(sheet, PROTECTION as never);
  (sheet as unknown as { ignoredErrors: unknown[] }).ignoredErrors.push(IGNORED);
  return Workbook.toBuffer(workbook, { format: "xlsb" });
}

async function streamed(): Promise<{ bytes: Uint8Array; unsupported: readonly string[] }> {
  const file = join(dir, "streamed.xlsb");
  const writer = new Stream.WorkbookWriter({ filename: file, format: "xlsb" });
  const sheet = writer.addWorksheet("S", OPTIONS as never);
  Stream.setCellValue(sheet.getCell("A1"), {
    text: "link",
    hyperlink: "https://example.com/"
  } as never);
  Stream.setCellValue(sheet.getCell("A2"), { text: "in", hyperlink: "#S!B2" } as never);
  Stream.setCellValue(sheet.getCell("B1"), 5);
  sheet.addConditionalFormatting({ ref: "B1:B5 C1:C5", rules: RULES } as never);
  sheet.addConditionalFormatting({ ref: "D1:D3", rules: SECOND } as never);
  await sheet.protect(PROTECTION as never);
  sheet.ignoredErrors.push(IGNORED as never);
  sheet.commit();
  await writer.commit();
  return { bytes: new Uint8Array(await readFile(file)), unsupported: writer.unsupported };
}

/** The records of the sheet part from the first conditional-formatting record to the end, by name and payload. */
async function epilogue(bytes: Uint8Array): Promise<string[]> {
  const parts = await extractAll(bytes);
  const out: string[] = [];
  let started = false;
  for (const record of iterateInterpretableRecords(
    parts.get("xl/worksheets/sheet1.bin")!.data,
    "s"
  )) {
    const name = recordSpec(record.id)?.name ?? `0x${record.id.toString(16)}`;
    started ||= name === "BrtEndSheetData";
    if (started) {
      const hex = Buffer.from(record.payload ?? []).toString("hex");
      // A relationship id is a name, and a streamed sheet claims its links' ids before the package allocates any
      // (`streamedHyperlinkRelId`), so `rId4` stands where the buffered writer has `rId1`. Everything else must match.
      out.push(
        `${name} ${name === "BrtHLink" ? hex.replace(/720049006400(?:3[0-9]00)+/, "rId?") : hex}`
      );
    }
  }
  return out;
}

async function features(bytes: Uint8Array) {
  const workbook = Workbook.create();
  await Workbook.read(workbook, bytes);
  const model = Workbook.getModel(workbook).worksheets[0]! as unknown as Record<string, unknown> & {
    rows?: { cells: { address: string; hyperlink?: string }[] }[];
  };
  return {
    conditionalFormattings: model.conditionalFormattings,
    hyperlinks: (model.rows ?? [])
      .flatMap(row => row.cells)
      .filter(cell => cell.hyperlink !== undefined)
      .map(cell => [cell.address, cell.hyperlink]),
    sheetProtection: model.sheetProtection,
    autoFilter: model.autoFilter,
    ignoredErrors: model.ignoredErrors
  };
}

describe("streamed XLSB sheet features", () => {
  it("writes the same epilogue records as the buffered writer", async () => {
    const reference = await buffered();
    const { bytes, unsupported } = await streamed();
    expect(unsupported).toEqual([]);
    const expected = await epilogue(reference);
    expect(expected.some(record => record.startsWith("BrtBeginCFRule"))).toBe(true);
    expect(await epilogue(bytes)).toEqual(expected);
  });

  it("reads back what the buffered writer's output reads back as", async () => {
    const reference = await features(await buffered());
    expect(reference.conditionalFormattings).toHaveLength(2);
    expect(reference.hyperlinks).toHaveLength(2);
    expect(reference.autoFilter).toBeDefined();
    expect(await features((await streamed()).bytes)).toEqual(reference);
  });

  it("states the dxf ids the conditional formats use in the styles part", async () => {
    const parts = await extractAll((await streamed()).bytes);
    const names = [...iterateInterpretableRecords(parts.get("xl/styles.bin")!.data, "s")].map(
      record => recordSpec(record.id)?.name
    );
    const reference = await extractAll(await buffered());
    const expected = [
      ...iterateInterpretableRecords(reference.get("xl/styles.bin")!.data, "s")
    ].filter(record => recordSpec(record.id)?.name === "BrtDXF").length;
    expect(expected).toBeGreaterThan(0);
    expect(names.filter(name => name === "BrtDXF")).toHaveLength(expected);
  });

  it("numbers rules uniquely across streamed sheets", async () => {
    const file = join(dir, "two.xlsb");
    const writer = new Stream.WorkbookWriter({ filename: file, format: "xlsb" });
    for (const name of ["One", "Two"]) {
      const sheet = writer.addWorksheet(name);
      Stream.setCellValue(sheet.getCell("A1"), 1);
      sheet.addConditionalFormatting({ ref: "A1", rules: SECOND } as never);
      sheet.commit();
    }
    await writer.commit();
    const workbook = Workbook.create();
    await Workbook.read(workbook, new Uint8Array(await readFile(file)));
    const priorities = Workbook.getModel(workbook).worksheets.flatMap(sheet =>
      (sheet.conditionalFormattings ?? []).flatMap(block =>
        block.rules.map(rule => (rule as { priority?: number }).priority)
      )
    );
    expect(priorities).toHaveLength(2);
    expect(new Set(priorities).size).toBe(2);
  });
});
