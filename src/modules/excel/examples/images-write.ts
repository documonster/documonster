import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";
import { Cell, Workbook } from "@excel/index";
import { addWorkbookImage } from "@excel/workbook-core";
import { addBackgroundImage, addImage } from "@excel/worksheet";

const exampleDir = path.dirname(fileURLToPath(import.meta.url));

const filename = process.argv[2];

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "blort");

Cell.setValue(ws, "B2", "Hello, World!");

const imageId = addWorkbookImage(wb, {
  filename: path.join(exampleDir, "data/image2.png"),
  extension: "png"
});
const backgroundId = addWorkbookImage(wb, {
  buffer: fs.readFileSync(path.join(exampleDir, "data/bubbles.jpg")),
  extension: "jpeg"
});
addImage(ws, imageId, {
  tl: { col: 1, row: 1 },
  br: { col: 3.5, row: 5.5 }
});
addImage(ws, imageId, "B7:E12");

addBackgroundImage(ws, backgroundId);

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
