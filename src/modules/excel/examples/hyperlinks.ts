import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";

import { Workbook } from "../../../index";

const [, , filename] = process.argv;

const wb = new Workbook();
const ws = wb.addWorksheet("Foo");

ws.getCell("A1").value = {
  hyperlink: "https://www.npmjs.com/package/excelts",
  text: "ExcelTS",
  tooltip: "https://www.npmjs.com/package/excelts"
};

const stopwatch = new HrStopwatch();
stopwatch.start();

try {
  await wb.xlsx.writeFile(filename);
  const micros = stopwatch.microseconds;
  console.log("Done.");
  console.log("Time taken:", micros);
} catch (error) {
  console.log((error as Error).message);
}
