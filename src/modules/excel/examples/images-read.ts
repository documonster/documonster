import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";
import { Workbook } from "@excel/index";
import { getXlsxIo } from "@excel/workbook";

const filename = process.argv[2];

const wb = Workbook.create();
const stopwatch = new HrStopwatch();
stopwatch.start();
getXlsxIo(wb)
  .readFile(filename)
  .then(() => {
    const micros = stopwatch.microseconds;
    console.log("Done.");
    console.log("Time taken:", micros);

    const _ws = Workbook.getWorksheet(wb, "blort")!;

    // const { image } = ws.background; // background property not supported
    // console.log('Media', image.name, image.type, image.buffer.length);
    console.log("Worksheet loaded successfully");
  })
  .catch(error => {
    console.log(error.message);
  });
