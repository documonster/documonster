import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";
import { Cell, Workbook } from "@excel/index";

const [, , filename] = process.argv;

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "Foo");

Cell.setValue(ws, "A1", {
  hyperlink: "https://www.npmjs.com/package/documonster",
  text: "Documonster",
  tooltip: "https://www.npmjs.com/package/documonster"
});

const stopwatch = new HrStopwatch();
stopwatch.start();

try {
  await Workbook.writeFile(wb, filename);
  const micros = stopwatch.microseconds;
  console.log("Done.");
  console.log("Time taken:", micros);
} catch (error) {
  console.log((error as Error).message);
}
