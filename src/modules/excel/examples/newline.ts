import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";
import { Cell, Workbook } from "@excel/index";

const [, , filename] = process.argv;

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "Foo");

Cell.setValue(ws, "A1", " H, \n W! ");
Cell.setNote(ws, "A1", " Hello, \n World! ");
Cell.setStyle(ws, "A1", { alignment: { wrapText: true } });

Cell.setValue(ws, "C1", "H,\nW!");
Cell.setNote(ws, "C1", "H,\nW!");
Cell.setStyle(ws, "C1", { alignment: { wrapText: true } });

const stopwatch = new HrStopwatch();
stopwatch.start();
try {
  await Workbook.writeXlsx(wb, filename);
  const micros = stopwatch.microseconds;
  console.log("Done.");
  console.log("Time taken:", micros);
} catch (error) {
  console.log(error.message);
}
