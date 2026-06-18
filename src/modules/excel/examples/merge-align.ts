import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";
import { Cell, Workbook, Worksheet } from "@excel/index";

const filename = process.argv[2];

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "blort");

Worksheet.addRow(ws, [1, 2, 3, 4]);
Worksheet.addRow(ws, ["one", "two", "three", "four"]);
Worksheet.addRow(ws, ["une", "deux", "trois", "quatre"]);
Worksheet.addRow(ws, ["uno", "due", "tre", "quatro"]);

Worksheet.merge(ws, "B2:C3");
Cell.setStyle(ws, "B2", { alignment: { horizontal: "center", vertical: "middle" } });

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
