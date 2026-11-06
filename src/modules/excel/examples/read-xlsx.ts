import path from "node:path";
import { fileURLToPath } from "node:url";

import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";
import { Workbook, Worksheet } from "@excel/index";

const exampleDir = path.dirname(fileURLToPath(import.meta.url));

const filename = process.argv[2] ?? path.join(exampleDir, "data/table.xlsx");

const stopwatch = new HrStopwatch();
stopwatch.start();

Workbook.readFile(filename)
  .then(wb => {
    const micros = stopwatch.microseconds;

    console.log("Loaded", filename);
    console.log("Time taken:", micros / 1000000);

    Workbook.eachWorksheet(wb, (sheet, id) => {
      console.log(id, Worksheet.getName(sheet));
    });
  })
  .catch(error => {
    console.error("something went wrong", error.stack);
  });
