import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";
import { Workbook } from "@excel/index";
import { getXlsxIo } from "@excel/workbook";
import { getSheetName } from "@excel/worksheet";

const filename = process.argv[2];
const wb = Workbook.create();

const stopwatch = new HrStopwatch();
stopwatch.start();

getXlsxIo(wb)
  .readFile(filename)
  .then(() => {
    const micros = stopwatch.microseconds;

    console.log("Loaded", filename);
    console.log("Time taken:", micros / 1000000);

    Workbook.eachSheet(wb, (sheet, id) => {
      console.log(id, getSheetName(sheet));
    });
  })
  .catch(error => {
    console.error("something went wrong", error.stack);
  });
