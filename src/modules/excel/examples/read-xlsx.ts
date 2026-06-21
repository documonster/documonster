import path from "node:path";
import { fileURLToPath } from "node:url";

import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";
import { Workbook, Worksheet } from "@excel/index";

const exampleDir = path.dirname(fileURLToPath(import.meta.url));

const filename = process.argv[2] ?? path.join(exampleDir, "data/table.xlsx");
const wb = Workbook.create();

const stopwatch = new HrStopwatch();
stopwatch.start();

Workbook.getXlsxIo(wb)
  .readFile(filename)
  .then(() => {
    const micros = stopwatch.microseconds;

    console.log("Loaded", filename);
    console.log("Time taken:", micros / 1000000);

    Workbook.eachSheet(wb, (sheet, id) => {
      console.log(id, Worksheet.getName(sheet));
    });
  })
  .catch(error => {
    console.error("something went wrong", error.stack);
  });
