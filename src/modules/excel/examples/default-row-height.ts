import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";
import { Cell, Workbook } from "@excel/index";

const [, , filename] = process.argv;

if (!filename) {
  console.error("Must specify a filename");
  process.exit(1);
}

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "blort");

Cell.setValue(ws, "B2", "Hello");
ws.properties.defaultRowHeight = 50;

const stopwatch = new HrStopwatch();
stopwatch.start();

try {
  await Workbook.writeXlsx(wb, filename);
  const micros = stopwatch.microseconds;
  console.log("Done.");
  console.log("Time taken:", micros);
} catch (error) {
  console.error((error as Error).stack);
}
