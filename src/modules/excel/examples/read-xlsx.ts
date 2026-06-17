import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";
import { Workbook, Worksheet } from "@excel/index";

const filename = process.argv[2];
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
