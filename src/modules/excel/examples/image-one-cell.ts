import path from "node:path";
import { fileURLToPath } from "node:url";

import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";
import { Workbook } from "@excel/index";
import { addWorkbookImage } from "@excel/workbook-core";
import { addImage } from "@excel/worksheet";

const exampleDir = path.dirname(fileURLToPath(import.meta.url));

const filename = process.argv[2];

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "blort");

const imageId = addWorkbookImage(wb, {
  filename: path.join(exampleDir, "data/image2.png"),
  extension: "png"
});
addImage(ws, imageId, {
  tl: { col: 0.1125, row: 0.4 },
  br: { col: 2.101046875, row: 3.4 },
  editAs: "oneCell"
});

const stopwatch = new HrStopwatch();
stopwatch.start();
try {
  await Workbook.writeXlsx(wb, filename);
  const micros = stopwatch.microseconds;
  console.log("Done.");
  console.log("Time taken:", micros);
} catch (error) {
  console.error(error.stack);
}
