import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";

import { Workbook } from "../../../index";

const exampleDir = path.dirname(fileURLToPath(import.meta.url));

const filename = process.argv[2];

const wb = new Workbook();
const ws = wb.addWorksheet("blort");

ws.getCell("B2").value = "Hello, World!";

const imageId = wb.addImage({
  filename: path.join(exampleDir, "data/image2.png"),
  extension: "png"
});
const backgroundId = wb.addImage({
  buffer: fs.readFileSync(path.join(exampleDir, "data/bubbles.jpg")),
  extension: "jpeg"
});
ws.addImage(imageId, {
  // tl: { col: 1, row: 1 },
  tl: "B2",
  ext: { width: 100, height: 100 }
});

ws.addBackgroundImage(backgroundId);

const stopwatch = new HrStopwatch();
stopwatch.start();
try {
  await wb.xlsx.writeFile(filename);
  const micros = stopwatch.microseconds;
  console.log("Done.");
  console.log("Time taken:", micros);
} catch (error) {
  console.error(error.stack);
}
