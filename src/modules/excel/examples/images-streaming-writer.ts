import path from "node:path";
import { fileURLToPath } from "node:url";

import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";
import { Stream } from "@excel/index";

const exampleDir = path.dirname(fileURLToPath(import.meta.url));

const [, , filename] = process.argv;

const wb = new Stream.WorkbookWriter({ filename });

const imageId = wb.addImage({
  filename: path.join(exampleDir, "data/image2.png"),
  extension: "png"
});

const ws = wb.addWorksheet("Foo");
ws.addBackgroundImage(imageId);

const stopwatch = new HrStopwatch();
stopwatch.start();

wb.commit()
  .then(() => {
    const micros = stopwatch.microseconds;
    console.log("Done.");
    console.log("Time taken:", micros);
  })
  .catch(error => {
    console.log(error.message);
  });
