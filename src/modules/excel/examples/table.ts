import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";
import { Table, Workbook, Worksheet } from "@excel/index";

const [, , filename] = process.argv;

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "Foo");

const now = new Date();
const today = Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDay());

Worksheet.setColumns(ws, [{ key: "date", width: 32 }, { key: "number" }, { key: "word" }]);

const words = [
  "Twas",
  "brillig",
  "and",
  "the",
  "slithy",
  "toves",
  "did",
  "gyre",
  "and",
  "gimble",
  "in",
  "the",
  "wabe"
];

Table.add(ws, {
  name: "TestTable",
  ref: "A1",
  headerRow: true,
  totalsRow: true,
  style: {
    theme: "TableStyleDark3",
    showRowStripes: true
  },
  columns: [
    { name: "Date", totalsRowLabel: "Totally", filterButton: true },
    {
      name: "Id",
      totalsRowFunction: "max",
      filterButton: true,
      totalsRowResult: 8,
      style: { numFmt: "0.00%" }
    },
    {
      name: "Word",
      filterButton: false,
      style: { font: { bold: true, name: "Comic Sans MS" } }
    }
  ],
  rows: words.map((word, i) => {
    const additionalDays = 86400 * i;
    return [new Date(today + additionalDays), i, word];
  })
});

const stopwatch = new HrStopwatch();
stopwatch.start();
try {
  await Workbook.writeFile(wb, filename);
  const micros = stopwatch.microseconds;
  console.log("Done.");
  console.log("Time taken:", micros);
} catch (error) {
  console.log(error.message);
}
