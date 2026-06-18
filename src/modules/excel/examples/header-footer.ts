import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";
import { Cell, Workbook } from "@excel/index";

const [, , filename] = process.argv;

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "Foo");
Cell.setValue(ws, "B2", "Hello, World!");

ws.headerFooter.oddHeader = "&CHello, Header!";
ws.headerFooter.oddFooter = "&CPage &P of &N";

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
