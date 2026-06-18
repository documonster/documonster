import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";
import { Cell, Workbook } from "@excel/index";

const filename = process.argv[2];

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "blort");

Cell.setValue(ws, "A1", 7);
Cell.setValue(ws, "B1", "Hello, World!");

const stopwatch = new HrStopwatch();
stopwatch.start();
Workbook.writeFile(wb, filename).then(() => {
  const micros = stopwatch.microseconds;
  console.log("Done.");
  console.log("Time taken:", micros);
});
// .catch(function(error) {
//    console.log(error.message);
// })
