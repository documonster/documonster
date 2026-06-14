import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";
import { Workbook } from "@excel/index";
import { getXlsxIo } from "@excel/workbook";

const filenameIn = process.argv[2];
const filenameOut = process.argv[3];

// all this script does is read a file and write to another
// useful for testing for lost properties

const stopwatch = new HrStopwatch();
const wb = Workbook.create();
stopwatch.start();
getXlsxIo(wb)
  .readFile(filenameIn)
  .then(() => Workbook.writeXlsx(wb, filenameOut))
  .then(() => {
    const micros = stopwatch.microseconds;
    console.log("Done.");
    console.log("Time taken:", micros);
  })
  .catch(error => {
    console.error("Error", error.stack);
  });
