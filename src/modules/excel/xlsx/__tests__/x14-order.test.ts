import { extractAll } from "@archive/unzip/extract";
import { expectValidXlsx } from "@excel/__tests__/helpers/expect-valid-xlsx";
import { Cell, ConditionalFormatting, Sparkline, Stream, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

/**
 * Microsoft's `x14:` extensions — the data bars, icon sets and sparklines of Excel 2010 — are not in ECMA-376, but
 * Excel holds their children to the order [MS-XLSX] §5.4 declares as strictly as it holds any other. The validator
 * ranks them by content models generated from that schema (`check-sequences.ts`); this writes every child the
 * writers can emit, from the document model and from the streaming writer, and holds the file to it. The data bar
 * wrote its negative fill after its negative border until this test existed.
 */

const argb = (value: string) => ({ argb: value });

/** Every option of the `x14:` conditional formats, on `ws`. */
function addConditionalFormats(ws: Parameters<typeof ConditionalFormatting.add>[0]): void {
  ConditionalFormatting.add(ws, {
    ref: "A1:A6",
    rules: [
      {
        type: "dataBar",
        gradient: false,
        minLength: 5,
        maxLength: 90,
        border: true,
        negativeBarColorSameAsPositive: false,
        negativeBarBorderColorSameAsPositive: false,
        axisPosition: "middle",
        direction: "leftToRight",
        cfvo: [
          { type: "num", value: -3 },
          { type: "num", value: 3 }
        ],
        color: argb("FF638EC6"),
        borderColor: argb("FF2F5597"),
        negativeFillColor: argb("FFFF0000"),
        negativeBorderColor: argb("FFC00000"),
        axisColor: argb("FF000000")
      }
    ]
  });
  ConditionalFormatting.add(ws, {
    ref: "B1:B6",
    rules: [
      {
        type: "iconSet",
        iconSet: "3Arrows",
        custom: true,
        reverse: true,
        showValue: false,
        cfvo: [
          { type: "percent", value: 0 },
          { type: "percent", value: 33 },
          { type: "percent", value: 67 }
        ]
      }
    ]
  });
}

/** Every option of an `x14:` sparkline group — the document model only: the streaming writer has no sparklines. */
function addSparklines(ws: Parameters<typeof Sparkline.add>[0]): void {
  Sparkline.add(ws, {
    type: "line",
    markers: true,
    high: true,
    low: true,
    first: true,
    last: true,
    negative: true,
    lineColor: "4472C4",
    highColor: "70AD47",
    lowColor: "C00000",
    firstColor: "5B9BD5",
    lastColor: "FFC000",
    markerColor: "595959",
    negativeColor: "FF0000",
    sparklines: [{ dataRef: "S!C1:F1", cellRef: "G1" }]
  });
}

const value = (row: number, col: number) => ((row * col) % 7) - 3;

async function expectExtensionsInOrder(bytes: Uint8Array, sparklines: boolean): Promise<void> {
  const sheet = new TextDecoder().decode(
    (await extractAll(bytes)).get("xl/worksheets/sheet1.xml")!.data
  );
  // The extensions are there, so the check has something to rank.
  expect(sheet).toContain("<x14:dataBar");
  expect(sheet).toContain("<x14:iconSet");
  if (sparklines) {
    expect(sheet).toContain("<x14:sparklineGroup");
  }
  await expectValidXlsx(bytes);
}

describe("x14 extension element order", () => {
  it("writes a data bar, an icon set and sparklines with every option in schema order", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    for (let row = 1; row <= 6; row++) {
      for (let col = 1; col <= 6; col++) {
        Cell.setValue(ws, { row, col }, value(row, col));
      }
    }
    addConditionalFormats(ws);
    addSparklines(ws);
    await expectExtensionsInOrder(await Workbook.toBuffer(wb), true);
  });

  it("writes the conditional formats in schema order from the streaming writer too", async () => {
    const writer = new Stream.WorkbookWriter({});
    const ws = writer.addWorksheet("S");
    addConditionalFormats(ws);
    for (let row = 1; row <= 6; row++) {
      Stream.commitRow(ws.addRow([1, 2, 3, 4, 5, 6].map(col => value(row, col))));
    }
    await expectExtensionsInOrder(await writer.toBuffer(), false);
  });
});
