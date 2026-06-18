import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";
import { Cell, Workbook, Worksheet } from "@excel/index";

const filename = process.argv[2];

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "blort");

const fills = {
  redDarkVertical: {
    type: "pattern",
    pattern: "darkVertical",
    fgColor: { argb: "FFFF0000" }
  },
  redGreenDarkTrellis: {
    type: "pattern",
    pattern: "darkTrellis",
    fgColor: { argb: "FFFF0000" },
    bgColor: { argb: "FF00FF00" }
  },
  blueWhiteHGrad: {
    type: "gradient",
    gradient: "angle",
    degree: 0,
    stops: [
      { position: 0, color: { argb: "FF0000FF" } },
      { position: 1, color: { argb: "FFFFFFFF" } }
    ]
  },
  rgbPathGrad: {
    type: "gradient",
    gradient: "path",
    center: { left: 0.5, top: 0.5 },
    stops: [
      { position: 0, color: { argb: "FFFF0000" } },
      { position: 0.5, color: { argb: "FF00FF00" } },
      { position: 1, color: { argb: "FF0000FF" } }
    ]
  }
} as const;

Worksheet.addRow(ws, [1, 2, 3, 4]);
Worksheet.addRow(ws, ["one", "two", "three", "four"]);
Worksheet.addRow(ws, ["une", "deux", "trois", "quatre"]);
Worksheet.addRow(ws, ["uno", "due", "tre", "quatro"]);

Cell.setStyle(ws, "B2", { fill: fills.redDarkVertical });

const stopwatch = new HrStopwatch();
stopwatch.start();
try {
  await Workbook.writeFile(wb, filename);
  const micros = stopwatch.microseconds;
  console.log("Done.");
  console.log("Time taken:", micros);
} catch (error) {
  console.log(error.message);
}
