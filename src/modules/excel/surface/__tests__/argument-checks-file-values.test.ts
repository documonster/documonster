/**
 * Arguments that end up in the file are checked where they are passed. Each refused value here used to be written as
 * something Excel repairs, dropped without a word, or refused only when the workbook was written — with an untyped
 * error from deep inside a writer.
 *
 * Plain assignable settings (`ws.pageSetup`, `ws.views`, …) cannot be checked at assignment, so the writers refuse
 * them; a value read from a file is the file's, so the readers keep it as read instead (see the last block).
 */
import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import {
  Cell,
  Chart,
  Chartsheet,
  ConditionalFormatting,
  DataValidation,
  ExcelArgumentError,
  Form,
  HeaderFooterImage,
  Image,
  ExcelImageError,
  ExcelInvalidAddressError,
  Note,
  Sparkline,
  Watermark,
  Workbook,
  Worksheet
} from "@excel/index";
import {
  findKeptSheetSettings,
  setAsideInvalidSheetSettings,
  withKeptSheetSettings
} from "@excel/utils/sheet-settings-checks";
import { describe, expect, it } from "vitest";

type Ws = ReturnType<typeof Workbook.addWorksheet>;

const PNG = new Uint8Array([
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0,
  0, 31, 21, 196, 137, 0, 0, 0, 13, 73, 68, 65, 84, 120, 156, 99, 248, 15, 0, 0, 1, 1, 0, 5, 24,
  216, 78, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130
]);
/** A value of the wrong type, passed where the type system would stop it. */
const bad = <T>(value: unknown) => value as T;

function setup() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Cell.setValue(ws, "A1", 1);
  Cell.setValue(ws, "A2", 2);
  return { wb, ws, image: Image.add(wb, { buffer: PNG, extension: "png" }) };
}

function chartOn(ws: Ws) {
  return Chart.place(ws, "C1:H10", { type: "bar", series: [{ values: "S!$A$1:$A$2" }] });
}

const sparkline = { dataRef: "S!A1:A2", cellRef: "B1" };
const cfExpression = { type: "expression" as const, formulae: ["A1>1"] };

type Case = [
  label: string,
  apply: (env: ReturnType<typeof setup>) => unknown,
  error: new (...args: never[]) => Error,
  message: RegExp
];

const REFUSED: Case[] = [
  // Data validation
  [
    "validation type",
    ({ ws }) => Cell.setValidation(ws, "A1", bad({ type: "bogus" })),
    ExcelArgumentError,
    /rule\.type/
  ],
  [
    "validation operator",
    ({ ws }) =>
      Cell.setValidation(ws, "A1", bad({ type: "whole", operator: "near", formulae: [1] })),
    ExcelArgumentError,
    /rule\.operator/
  ],
  [
    "validation error style",
    ({ ws }) =>
      DataValidation.add(ws.dataValidations, "A1", bad({ type: "any", errorStyle: "loud" })),
    ExcelArgumentError,
    /rule\.errorStyle/
  ],
  [
    "validation bound",
    ({ ws }) => Cell.setValidation(ws, "A1", { type: "whole", formulae: [NaN] }),
    ExcelArgumentError,
    /rule\.formulae\[0\]/
  ],
  [
    "validation with three bounds",
    ({ ws }) => Cell.setValidation(ws, "A1", { type: "whole", formulae: [1, 2, 3] }),
    ExcelArgumentError,
    /rule\.formulae/
  ],
  // Panes and the auto filter
  ["freeze -2", ({ ws }) => Worksheet.freeze(ws, { columns: -2 }), ExcelArgumentError, /columns/],
  ["freeze NaN", ({ ws }) => Worksheet.freeze(ws, { columns: NaN }), ExcelArgumentError, /columns/],
  ["freeze 1.5", ({ ws }) => Worksheet.freeze(ws, { rows: 1.5 }), ExcelArgumentError, /rows/],
  [
    "freeze past the grid",
    ({ ws }) => Worksheet.freeze(ws, { rows: 1048576 }),
    ExcelArgumentError,
    /rows/
  ],
  [
    "freeze positional NaN",
    ({ ws }) => Worksheet.freeze(ws, NaN, 1),
    ExcelArgumentError,
    /columns/
  ],
  ["split NaN", ({ ws }) => Worksheet.split(ws, { x: NaN }), ExcelArgumentError, /\bx\b/],
  ["split negative", ({ ws }) => Worksheet.split(ws, { y: -5 }), ExcelArgumentError, /\by\b/],
  [
    "auto filter",
    ({ ws }) => Worksheet.setAutoFilter(ws, "zz"),
    ExcelInvalidAddressError,
    /setAutoFilter/
  ],
  [
    "auto filter on another sheet",
    ({ ws }) => Worksheet.setAutoFilter(ws, "S!A1:B2"),
    ExcelInvalidAddressError,
    /setAutoFilter/
  ],
  [
    "auto filter corner",
    ({ ws }) => Worksheet.setAutoFilter(ws, { from: { row: 0, col: 1 }, to: "B2" }),
    ExcelArgumentError,
    /range\.from\.row/
  ],
  // Other worksheet operations
  ["unmerge", ({ ws }) => Worksheet.unmerge(ws, "zz"), ExcelInvalidAddressError, /unmerge/],
  [
    "fill formula range",
    ({ ws }) => Worksheet.fillFormula(ws, "zz", "1"),
    ExcelInvalidAddressError,
    /fillFormula/
  ],
  [
    "fill formula text",
    ({ ws }) => Worksheet.fillFormula(ws, "A1:A2", ""),
    ExcelArgumentError,
    /formula/
  ],
  [
    "fill formula share type",
    ({ ws }) => Worksheet.fillFormula(ws, "A1:A2", "1", { shareType: bad("both") }),
    ExcelArgumentError,
    /shareType/
  ],
  [
    "add rows of arrays: not an array",
    ({ ws }) => Worksheet.addAoa(ws, bad(5)),
    ExcelArgumentError,
    /data/
  ],
  [
    "add rows of arrays: a row",
    ({ ws }) => Worksheet.addAoa(ws, bad([5])),
    ExcelArgumentError,
    /data\[0\]/
  ],
  [
    "add rows of arrays past the grid",
    ({ ws }) => Worksheet.addAoa(ws, [[1, 2]], { origin: "XFD1" }),
    Error,
    /out of bounds/
  ],
  [
    "add records: a record",
    ({ ws }) => Worksheet.addJson(ws, bad([5])),
    ExcelArgumentError,
    /data\[0\]/
  ],
  // Conditional formatting
  [
    "rule type",
    ({ ws }) => ConditionalFormatting.add(ws, { ref: "A1", rules: [bad({ type: "magic" })] }),
    ExcelArgumentError,
    /rules\[0\]\.type/
  ],
  [
    "empty ref",
    ({ ws }) => ConditionalFormatting.add(ws, { ref: "", rules: [cfExpression] }),
    ExcelInvalidAddressError,
    /ref/
  ],
  [
    "no rules",
    ({ ws }) => ConditionalFormatting.add(ws, { ref: "A1", rules: [] }),
    ExcelArgumentError,
    /rules/
  ],
  [
    "expression with no formula",
    ({ ws }) => ConditionalFormatting.add(ws, { ref: "A1", rules: [{ type: "expression" }] }),
    ExcelArgumentError,
    /needs 1 formula/
  ],
  [
    "between with one operand",
    ({ ws }) =>
      ConditionalFormatting.add(ws, {
        ref: "A1",
        rules: [{ type: "cellIs", operator: "between", formulae: [1] }]
      }),
    ExcelArgumentError,
    /needs 2 formula/
  ],
  [
    "comparison operator",
    ({ ws }) =>
      ConditionalFormatting.add(ws, {
        ref: "A1",
        rules: [bad({ type: "cellIs", operator: "near", formulae: [1] })]
      }),
    ExcelArgumentError,
    /operator/
  ],
  [
    "colour scale without thresholds",
    ({ ws }) => ConditionalFormatting.add(ws, { ref: "A1", rules: [{ type: "colorScale" }] }),
    ExcelArgumentError,
    /cfvo/
  ],
  [
    "icon set name",
    ({ ws }) =>
      ConditionalFormatting.add(ws, {
        ref: "A1",
        rules: [
          bad({ type: "iconSet", iconSet: "9Bogus", cfvo: [{ type: "min" }, { type: "max" }] })
        ]
      }),
    ExcelArgumentError,
    /iconSet/
  ],
  [
    "time period",
    ({ ws }) =>
      ConditionalFormatting.add(ws, {
        ref: "A1",
        rules: [bad({ type: "timePeriod", timePeriod: "someday" })]
      }),
    ExcelArgumentError,
    /timePeriod/
  ],
  [
    "text rule without its search string",
    ({ ws }) => ConditionalFormatting.add(ws, { ref: "A1", rules: [{ type: "containsText" }] }),
    ExcelArgumentError,
    /text/
  ],
  [
    "top-10 rank",
    ({ ws }) =>
      ConditionalFormatting.add(ws, {
        ref: "A1",
        rules: [{ type: "top10", rank: -1, percent: false }]
      }),
    ExcelArgumentError,
    /rank/
  ],
  [
    "priority",
    ({ ws }) =>
      ConditionalFormatting.add(ws, { ref: "A1", rules: [{ ...cfExpression, priority: 0 }] }),
    ExcelArgumentError,
    /priority/
  ],
  // Sparklines
  [
    "sparkline type",
    ({ ws }) => Sparkline.add(ws, bad({ type: "pie", sparklines: [sparkline] })),
    ExcelArgumentError,
    /type/
  ],
  [
    "sparkline members",
    ({ ws }) => Sparkline.add(ws, bad({ type: "line" })),
    ExcelArgumentError,
    /sparklines/
  ],
  [
    "sparkline cell",
    ({ ws }) => Sparkline.add(ws, { type: "line", sparklines: [{ ...sparkline, cellRef: "zz" }] }),
    ExcelInvalidAddressError,
    /cellRef/
  ],
  [
    "sparkline colour",
    ({ ws }) => Sparkline.add(ws, { type: "line", lineColor: "red", sparklines: [sparkline] }),
    ExcelArgumentError,
    /lineColor/
  ],
  [
    "sparkline weight",
    ({ ws }) => Sparkline.add(ws, { type: "line", lineWeight: NaN, sparklines: [sparkline] }),
    ExcelArgumentError,
    /lineWeight/
  ],
  [
    "sparkline axis",
    ({ ws }) =>
      Sparkline.add(ws, bad({ type: "line", minAxisType: "auto", sparklines: [sparkline] })),
    ExcelArgumentError,
    /minAxisType/
  ],
  // Notes
  ["note of a number", ({ ws }) => Cell.setNote(ws, "A1", bad(123)), ExcelArgumentError, /note/],
  [
    "note texts",
    ({ ws }) => Cell.setNote(ws, "A1", bad({ texts: "x" })),
    ExcelArgumentError,
    /note\.texts/
  ],
  [
    "note text",
    ({ ws }) => Cell.setNote(ws, "A1", bad({ texts: [{ text: 5 }] })),
    ExcelArgumentError,
    /texts\[0\]\.text/
  ],
  [
    "note positioning",
    ({ ws }) => Cell.setNote(ws, "A1", bad({ texts: [{ text: "a" }], editAs: "floating" })),
    ExcelArgumentError,
    /editAs/
  ],
  [
    "note size",
    ({ ws }) => Cell.setNote(ws, "A1", { texts: [{ text: "a" }], width: NaN }),
    ExcelArgumentError,
    /width/
  ],
  [
    "comment of a number",
    ({ ws }) => Cell.setComment(ws, "A1", bad(5)),
    ExcelArgumentError,
    /comment/
  ],
  ["note created from a number", () => Note.create(bad(5)), ExcelArgumentError, /note/],
  ["note author", () => Note.create("x", { author: bad(5) }), ExcelArgumentError, /author/],
  // Form controls
  [
    "checkbox range",
    ({ ws }) => Form.addCheckbox(ws, "nowhere"),
    ExcelInvalidAddressError,
    /range/
  ],
  [
    "checkbox position",
    ({ ws }) => Form.addCheckbox(ws, { startCol: NaN, startRow: 0, endCol: 1, endRow: 1 }),
    ExcelArgumentError,
    /range\.startCol/
  ],
  [
    "checkbox label",
    ({ ws }) => Form.addCheckbox(ws, "B2", bad({ text: 5 })),
    ExcelArgumentError,
    /text/
  ],
  [
    "checkbox relabelled",
    ({ ws }) => Form.setText(Form.addCheckbox(ws, "B2"), bad(5)),
    ExcelArgumentError,
    /Form\.setText: value/
  ],
  // Protection
  [
    "sheet protection flag",
    ({ ws }) => Worksheet.protect(ws, bad({ formatCells: "yes" })),
    ExcelArgumentError,
    /formatCells/
  ],
  [
    "sheet password",
    ({ ws }) => Worksheet.protect(ws, bad({ password: 123 })),
    ExcelArgumentError,
    /password/
  ],
  [
    "workbook protection flag",
    ({ wb }) => Workbook.protect(wb, bad({ lockStructure: "yes" })),
    ExcelArgumentError,
    /lockStructure/
  ],
  // Workbook
  ["person", ({ wb }) => Workbook.registerPerson(wb, bad(5)), ExcelArgumentError, /displayName/],
  [
    "external link",
    ({ wb }) => Workbook.addExternalLink(wb, bad({ target: 5 })),
    ExcelArgumentError,
    /target/
  ],
  [
    "empty external link",
    ({ wb }) => Workbook.addExternalLink(wb, { target: "" }),
    ExcelArgumentError,
    /target/
  ],
  [
    "worksheet options",
    ({ wb }) => Workbook.addWorksheet(wb, "T", bad({ pageSetup: { orientation: "sideways" } })),
    ExcelArgumentError,
    /pageSetup\.orientation/
  ],
  [
    "worksheet state",
    ({ wb }) => Workbook.addWorksheet(wb, "T", bad({ state: "invisible" })),
    ExcelArgumentError,
    /state/
  ],
  // Pictures in the header and on the sheet
  ["watermark image", ({ ws }) => Watermark.set(ws, { imageId: 99 }), ExcelImageError, /no image/],
  [
    "watermark mode",
    ({ ws, image }) => Watermark.set(ws, bad({ imageId: image, mode: "ghost" })),
    ExcelArgumentError,
    /mode/
  ],
  [
    "watermark opacity",
    ({ ws, image }) => Watermark.set(ws, { imageId: image, opacity: 5 }),
    ExcelArgumentError,
    /opacity/
  ],
  [
    "watermark section",
    ({ ws, image }) => Watermark.set(ws, bad({ imageId: image, mode: "header", position: "XX" })),
    ExcelArgumentError,
    /position/
  ],
  [
    "header image size",
    ({ ws, image }) => HeaderFooterImage.set(ws, { imageId: image, position: "CH", width: NaN }),
    ExcelArgumentError,
    /width/
  ],
  [
    "header image",
    ({ ws }) => HeaderFooterImage.set(ws, { imageId: 77, position: "CH" }),
    ExcelImageError,
    /no image/
  ],
  [
    "picture range",
    ({ ws, image }) => Image.place(ws, image, "zz"),
    ExcelInvalidAddressError,
    /Image\.place/
  ],
  // Charts
  [
    "chart range",
    ({ ws }) => Chart.add(ws, { type: "bar", series: [{ values: "S!$A$1:$A$2" }] }, "zz"),
    ExcelInvalidAddressError,
    /Chart\.add/
  ],
  [
    "chart anchor",
    ({ ws }) =>
      Chart.place(
        ws,
        {
          tl: { col: NaN, row: 0 },
          br: { col: 2, row: 2 }
        },
        { type: "bar", series: [{ values: "S!$A$1:$A$2" }] }
      ),
    ExcelArgumentError,
    /range\.tl\.col/
  ],
  [
    "chart series name",
    ({ ws }) => Chart.setSeriesName(chartOn(ws), 0, bad(5)),
    ExcelArgumentError,
    /name/
  ],
  [
    "chart series name in the options",
    ({ ws }) =>
      Chart.place(ws, "C1:H10", {
        type: "bar",
        series: [{ values: "S!$A$1:$A$2", name: bad({ bogus: 1 }) }]
      }),
    ExcelArgumentError,
    /series\[0\]\.name/
  ],
  [
    "chart series categories",
    ({ ws }) => Chart.setSeriesCategories(chartOn(ws), 0, bad(5)),
    ExcelArgumentError,
    /categories/
  ],
  [
    "chart legend position",
    ({ ws }) => Chart.setLegend(chartOn(ws), bad({ legendPos: "middle" })),
    ExcelArgumentError,
    /legendPos/
  ],
  // Chartsheets
  [
    "chartsheet page setup",
    ({ wb }) => {
      const cs = Workbook.addChartsheet(wb, "CS", {
        chart: { type: "bar", series: [{ values: "S!$A$1:$A$2" }] }
      });
      Chartsheet.setPageSetup(cs, bad({ orientation: "sideways" }));
    },
    ExcelArgumentError,
    /orientation/
  ]
];

describe("a value that cannot reach the file is refused at the call", () => {
  it.each(REFUSED)("%s", async (_, apply, error, message) => {
    const env = setup();
    let thrown: unknown;
    try {
      // Async refusals (`protect`) reject rather than throw.
      await apply(env);
    } catch (e) {
      thrown = e;
    }
    expect(thrown).toBeInstanceOf(error);
    expect(String((thrown as Error).message)).toMatch(message);
    // What was refused was not stored, so the workbook still writes.
    await expect(Workbook.toBuffer(env.wb)).resolves.toBeInstanceOf(Uint8Array);
  });

  it("a refused rule, sparkline or checkbox leaves nothing behind", () => {
    const { ws } = setup();
    expect(() => Cell.setValidation(ws, "D9", bad({ type: "bogus" }))).toThrow();
    expect(Cell.getValidation(ws, "D9")).toBeUndefined();
    expect(() =>
      ConditionalFormatting.add(ws, { ref: "A1", rules: [bad({ type: "magic" })] })
    ).toThrow();
    expect(ConditionalFormatting.list(ws)).toEqual([]);
    expect(() => Sparkline.add(ws, bad({ type: "line" }))).toThrow();
    expect(Sparkline.list(ws)).toEqual([]);
    expect(() => Form.addCheckbox(ws, "nowhere")).toThrow();
    expect(Form.list(ws)).toEqual([]);
    expect(() => Worksheet.addAoa(ws, [[5], [NaN]], { origin: "F1" })).toThrow();
    expect(Cell.getValue(ws, "F1")).toBeNull();
  });
});

describe("plain settings the caller assigned are refused when the workbook is written", () => {
  const ASSIGNED: [string, (ws: Ws) => void, RegExp][] = [
    ["orientation", ws => (ws.pageSetup.orientation = "sideways"), /pageSetup\.orientation/],
    ["page order", ws => (ws.pageSetup.pageOrder = "sideways"), /pageSetup\.pageOrder/],
    ["printed errors", ws => (ws.pageSetup.errors = "hidden"), /pageSetup\.errors/],
    ["scale", ws => (ws.pageSetup.scale = 12.5), /pageSetup\.scale/],
    ["margin", ws => (ws.pageSetup.margins.left = NaN), /pageSetup\.margins\.left/],
    [
      "header flag",
      ws => (ws.headerFooter.differentFirst = bad("yes")),
      /headerFooter\.differentFirst/
    ],
    ["tab colour", ws => (ws.properties.tabColor = { argb: "red" }), /properties\.tabColor\.argb/],
    ["state", ws => (ws.state = bad("invisible")), /\bstate\b/],
    ["view", ws => (ws.views = [{ state: "frozen", xSplit: NaN }]), /views\[0\]\.xSplit/],
    [
      "view style",
      ws => (ws.views = [bad({ state: "normal", style: "wide" })]),
      /views\[0\]\.style/
    ]
  ];

  for (const format of ["xlsx", "xlsb"] as const) {
    it.each(ASSIGNED)(`%s (${format})`, async (_, assign, field) => {
      const { wb, ws } = setup();
      assign(ws);
      const write = Workbook.toBuffer(wb, { format });
      await expect(write).rejects.toBeInstanceOf(ExcelArgumentError);
      await expect(write).rejects.toThrow(field);
      await expect(write).rejects.toThrow(/worksheet "S"/);
    });
  }
});

describe("values Excel accepts still work, at the boundaries", () => {
  it("in every namespace checked above", async () => {
    const { wb, ws, image } = setup();
    Cell.setValidation(ws, "A1", { type: "list", formulae: ['"a,b"'] });
    Cell.setValidation(ws, "A2", { type: "whole", operator: "between", formulae: [] });
    Cell.setValidation(ws, "A3", { type: "date", formulae: [new Date(Date.UTC(2020, 0, 1))] });
    Cell.setValidation(ws, "A4", undefined as never);
    DataValidation.add(ws.dataValidations, "B:B", { type: "any", errorStyle: "warning" });
    Worksheet.freeze(ws, { columns: 16383, rows: 0 });
    Worksheet.freeze(ws, null as never, 2);
    Worksheet.freeze(ws, "2" as never, true as never);
    Worksheet.split(ws, { x: 1500.5 });
    Worksheet.unfreeze(ws);
    Worksheet.setAutoFilter(ws, "A1:B2");
    Worksheet.setAutoFilter(ws, { from: "A1", to: { row: 1048576, col: 16384 } });
    Worksheet.setAutoFilter(ws, "$A$1:$B$9");
    Worksheet.setAutoFilter(ws, null);
    Worksheet.fillFormula(ws, "D1:D3", "A1*2");
    Worksheet.addAoa(ws, [[1, 2], null as never, [3]], { origin: "F1" });
    Worksheet.addJson(ws, [{ a: 1, b: null }], { origin: -1, nullError: true });
    ConditionalFormatting.add(ws, {
      ref: "A1:A3 C1:C3",
      rules: [
        cfExpression,
        { type: "cellIs", operator: "between", formulae: [1, 5] },
        { type: "cellIs", operator: "notEqual" as never, formulae: [1] },
        { type: "top10", rank: 100, percent: true },
        {
          type: "colorScale",
          cfvo: [{ type: "min" }, { type: "max" }],
          color: [{ argb: "FFFF0000" }, { argb: "FF00FF00" }]
        },
        { type: "dataBar" },
        { type: "timePeriod", timePeriod: "today" },
        { type: "containsText", operator: "containsBlanks" },
        { type: "containsText", text: "x" },
        { type: "containsBlanks" } as never,
        { type: "uniqueValues" } as never
      ]
    });
    ConditionalFormatting.add(ws, { ref: "A:A", rules: [{ ...cfExpression, priority: 1 }] });
    Sparkline.add(ws, {
      type: "column",
      lineColor: "#FF0000",
      negativeColor: "FFFF0000",
      highColor: "",
      lineWeight: 0,
      sparklines: [{ dataRef: "Rates", cellRef: "XFD1048576" }]
    });
    Cell.setNote(ws, "B1", "plain");
    Cell.setNote(ws, "B2", {
      texts: [{ text: "rich", font: { bold: true } }],
      margins: { insetmode: "custom", inset: [0, 0.1, 0.2, 0.3] },
      protection: { locked: "False", lockText: "f" },
      editAs: "oneCells",
      width: 0
    });
    Cell.setComment(ws, "B3", Note.create("n", { author: "Ann" }));
    Cell.setComment(ws, "B3", undefined);
    Note.create("t", 5 as never);
    const box = Form.addCheckbox(
      ws,
      { tl: "E5", br: { col: 6, row: 6, colOff: 0 } },
      { link: "Other!A1", text: "" }
    );
    Form.setLink(box, undefined);
    Form.addCheckbox(ws, { startCol: 0, startRow: 0, endCol: 16383, endRow: 1048575 });
    await Worksheet.protect(ws, { password: "pw", formatCells: true, spinCount: 1000 });
    await Worksheet.protect(ws, undefined, { formatCells: false });
    await Workbook.protect(wb, { lockStructure: true });
    Workbook.registerPerson(wb, "Ann", { userId: "ann@example.com" });
    Workbook.registerPerson(wb, "Bob", null as never, "AD");
    Workbook.addExternalLink(wb, { target: "data/ref.xlsx", sheetNames: ["Data"] });
    Workbook.addWorksheet(wb, "T", {
      state: "veryHidden",
      pageSetup: { orientation: "landscape", scale: 400, fitToWidth: 0 },
      properties: { tabColor: { argb: "FF00FF00" } },
      views: [{ state: "frozen", xSplit: 1, ySplit: 1, topLeftCell: "B2" }],
      autoFilter: "A1:C3"
    });
    Watermark.set(ws, { imageId: image, opacity: 0 });
    Watermark.set(ws, {
      imageId: image,
      mode: "header",
      position: "RF",
      headerWidth: 0,
      applyTo: "all"
    });
    HeaderFooterImage.set(ws, { imageId: image, position: "LH", width: 100, applyTo: "first" });
    Image.place(ws, image, "J1:K2");
    const chart = chartOn(ws);
    Chart.setSeriesName(chart, 0, { formula: "S!$A$1" });
    Chart.setSeriesCategories(chart, 0, "S!$A$1:$A$2");
    Chart.setLegend(chart, { legendPos: "tr", overlay: false });
    Chart.setLegend(chart, undefined);
    Chart.setSpPr(chart, undefined);
    Chart.add(
      ws,
      { type: "bar", series: [{ values: "S!$A$1:$A$2" }] },
      { tl: { col: 0.5, row: 0 }, ext: { cx: 100, cy: 100 } }
    );
    ws.pageSetup.orientation = "default";
    ws.pageSetup.cellComments = "asDisplayed";
    ws.views = [{ state: "split", xSplit: 1200.5, activePane: "topRight" }];
    for (const format of ["xlsx", "xlsb"] as const) {
      expect(
        (await Workbook.toBuffer(wb, { format, unsupported: "ignore" })).length
      ).toBeGreaterThan(0);
    }
  });

  it("writes cellComments as the schema spells it, for both spellings", async () => {
    for (const spelling of ["asDisplayed", "asDisplyed"]) {
      const { wb, ws } = setup();
      ws.pageSetup.cellComments = spelling;
      const entries = await extractAll(await Workbook.toBuffer(wb));
      const sheet = new TextDecoder().decode(entries.get("xl/worksheets/sheet1.xml")!.data);
      expect(sheet).toMatch(/cellComments="asDisplayed"/);
    }
  });
});

describe("a value the source file carried is kept as read, not dropped or refused", () => {
  /** An XLSX whose first sheet's markup has `patch` applied. */
  async function patched(patch: (xml: string) => string): Promise<Uint8Array> {
    const { wb, ws } = setup();
    Worksheet.freeze(ws, { rows: 1 });
    const entries = await extractAll(await Workbook.toBuffer(wb));
    return createZip(
      [...entries]
        .filter(([, file]) => file.type !== "directory")
        .map(([name, file]) => ({
          name,
          data:
            name === "xl/worksheets/sheet1.xml"
              ? new TextEncoder().encode(patch(new TextDecoder().decode(file.data)))
              : file.data
        }))
    );
  }

  async function sheetXml(bytes: Uint8Array): Promise<string> {
    return new TextDecoder().decode(
      (await extractAll(bytes)).get("xl/worksheets/sheet1.xml")!.data
    );
  }

  const sideways = (xml: string) => xml.replace('orientation="portrait"', 'orientation="sideways"');

  it('orientation="sideways" reads without the field and is written back as it was', async () => {
    const bytes = await patched(sideways);
    expect(await sheetXml(bytes)).toContain('orientation="sideways"');
    const wb = await Workbook.read(bytes);
    const ws = Workbook.getWorksheet(wb, "S")!;
    expect(ws.pageSetup.orientation).toBeUndefined();
    expect(await sheetXml(await Workbook.toBuffer(wb))).toContain('orientation="sideways"');
    // XLSB has no record for the text, so converting refuses unless told to drop it.
    await expect(Workbook.toBuffer(wb, { format: "xlsb" })).rejects.toThrow(
      /pageSetup\.orientation "sideways"/
    );
    await expect(
      Workbook.toBuffer(wb, { format: "xlsb", unsupported: "ignore" })
    ).resolves.toBeInstanceOf(Uint8Array);
  });

  it("a value set afterwards replaces it, and is checked as any caller's", async () => {
    const wb = await Workbook.read(await patched(sideways));
    const ws = Workbook.getWorksheet(wb, "S")!;
    ws.pageSetup.orientation = "landscape";
    const xml = await sheetXml(await Workbook.toBuffer(wb));
    expect(xml).toContain('orientation="landscape"');
    expect(xml).not.toContain("sideways");
    ws.pageSetup.orientation = "sideways";
    await expect(Workbook.toBuffer(wb)).rejects.toThrow(/pageSetup\.orientation/);
  });

  it("an unknown view type, pane or number is kept too, and the rest of the view read", async () => {
    const bytes = await patched(xml =>
      xml
        .replace("<sheetView ", '<sheetView view="wide" zoomScale="abc" ')
        .replace(/activePane="[a-zA-Z]+"/, 'activePane="middle"')
    );
    const wb = await Workbook.read(bytes);
    const view = Workbook.getWorksheet(wb, "S")!.views[0]!;
    expect(view).toMatchObject({ state: "frozen", ySplit: 1 });
    expect(view.style).toBeUndefined();
    expect(view.zoomScale).toBeUndefined();
    const xml = await sheetXml(await Workbook.toBuffer(wb));
    expect(xml).toContain('view="wide"');
    expect(xml).toContain('zoomScale="abc"');
  });

  it("readWithDiagnostics reports each kept value, from either format", async () => {
    const bytes = await patched(xml =>
      sideways(xml).replace("<sheetView ", '<sheetView zoomScale="abc" ')
    );
    const report = await Workbook.readWithDiagnostics(bytes);
    expect(report.lost).toEqual([]);
    // One entry each, in the shape the discard filter receives.
    expect(report.unrecognizedValues).toEqual([
      { sheet: "S", path: "pageSetup.orientation", value: "sideways" },
      { sheet: "S", path: "views[0].zoomScale", value: "abc" }
    ]);
    const seen: unknown[] = [];
    Workbook.discardUnrecognizedValues(report.workbook, {
      filter: setting => {
        seen.push(setting);
        return false;
      }
    });
    expect(seen).toEqual(report.unrecognizedValues);
    const clean = await Workbook.readWithDiagnostics(await Workbook.toBuffer(Workbook.create()));
    expect(clean.unrecognizedValues).toEqual([]);
  });

  it("Workbook.discardUnrecognizedValues removes them, which clearing the field does not", async () => {
    const bytes = await patched(xml =>
      sideways(xml).replace("<sheetView ", '<sheetView zoomScale="abc" ')
    );
    const wb = await Workbook.read(bytes);
    const ws = Workbook.getWorksheet(wb, "S")!;
    ws.pageSetup.orientation = "landscape";
    delete (ws.pageSetup as { orientation?: string }).orientation;
    expect(await sheetXml(await Workbook.toBuffer(wb))).toContain('orientation="sideways"');

    expect(Workbook.discardUnrecognizedValues(wb)).toBe(2);
    const xml = await sheetXml(await Workbook.toBuffer(wb));
    expect(xml).not.toContain("sideways");
    expect(xml).not.toContain('zoomScale="abc"');
    expect(Workbook.discardUnrecognizedValues(wb)).toBe(0);
    const report = await Workbook.readWithDiagnostics(await Workbook.toBuffer(wb));
    expect(report.unrecognizedValues).toEqual([]);
  });

  it("Workbook.discardUnrecognizedValues with a filter drops only the settings it picks", async () => {
    const bytes = await patched(xml =>
      sideways(xml).replace("<sheetView ", '<sheetView zoomScale="abc" ')
    );
    const wb = await Workbook.read(bytes);
    const seen: unknown[] = [];
    expect(
      Workbook.discardUnrecognizedValues(wb, {
        filter: setting => {
          seen.push(setting);
          return setting.path === "views[0].zoomScale";
        }
      })
    ).toBe(1);
    expect(seen).toEqual([
      { sheet: "S", path: "pageSetup.orientation", value: "sideways" },
      { sheet: "S", path: "views[0].zoomScale", value: "abc" }
    ]);
    const xml = await sheetXml(await Workbook.toBuffer(wb));
    expect(xml).toContain('orientation="sideways"');
    expect(xml).not.toContain('zoomScale="abc"');
    expect(() => Workbook.discardUnrecognizedValues(wb, { filter: 1 as never })).toThrow(
      /options\.filter/
    );
  });

  it("setting aside keeps the text on the object that held it, and drops only what has none", () => {
    const sheet = {
      pageSetup: { orientation: "sideways", scale: 100 },
      views: [{ state: "frozen", xSplit: NaN, ySplit: 2.5 }]
    };
    expect(setAsideInvalidSheetSettings("S", sheet)).toEqual([
      "S: views[0].xSplit NaN is not an integer from 0 to 16383, and was read as absent"
    ]);
    expect(sheet).toEqual({
      pageSetup: { scale: 100, _unrecognized: { orientation: { value: "sideways" } } },
      views: [{ state: "frozen", _unrecognized: { ySplit: { value: "2.5" } } }]
    });
    expect(findKeptSheetSettings(sheet).map(k => k.path)).toEqual([
      "pageSetup.orientation",
      "views[0].ySplit"
    ]);
    const written = withKeptSheetSettings(sheet);
    expect(written.pageSetup).toEqual({ scale: 100, orientation: "sideways" });
    expect(written.views[0]).toEqual({ state: "frozen", ySplit: "2.5" });
    // The sheet itself is not changed.
    expect(sheet.pageSetup).not.toHaveProperty("orientation");
  });
});
