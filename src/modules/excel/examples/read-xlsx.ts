import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";

import { Workbook } from "../../../index";

const filename = process.argv[2];
const wb = new Workbook();

const stopwatch = new HrStopwatch();
stopwatch.start();

wb.xlsx
  .readFile(filename)
  .then(() => {
    const micros = stopwatch.microseconds;

    console.log("Loaded", filename);
    console.log("Time taken:", micros / 1000000);

    wb.eachSheet((sheet, id) => {
      console.log(id, sheet.name);
    });
  })
  .catch(error => {
    console.error("something went wrong", error.stack);
  });
