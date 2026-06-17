import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";
import { Table, Workbook, Worksheet } from "@excel/index";

const [, , filename] = process.argv;

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "Foo");

const now = Date.now();
const today = now - (now % 86400000);

const getRows = () => {
  const rows: (number | Date)[][] = [];
  for (let i = 0; i < 20; i++) {
    rows.push([new Date(today + 86400000 * i), Math.random() * 10]);
  }
  return rows;
};

Worksheet.setColumns(ws, [{ key: "date", width: 16 }, { key: "number" }]);

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
    { name: "Date", totalsRowLabel: "Max:", filterButton: true },
    {
      name: "Value",
      totalsRowFunction: "max",
      filterButton: true,
      totalsRowResult: 8
    }
  ],
  rows: getRows()
});

const stopwatch = new HrStopwatch();
stopwatch.start();

try {
  await Workbook.writeXlsx(wb, filename);
  const micros = stopwatch.microseconds;
  console.log("Done.");
  console.log("Time taken:", micros);
} catch (error) {
  console.log((error as Error).message);
}
