import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";
import { Stream } from "@excel/index";

const [, , filename] = process.argv;

const wb = new Stream.WorkbookWriter({ filename });
const ws = wb.addWorksheet("Foo");
Stream.setCellValue(ws.getCell("B2"), 5);
Stream.setCellNote(ws.getCell("B2"), {
  texts: [
    {
      font: {
        size: 12,
        color: { theme: 0 },
        name: "Calibri",
        family: 2,
        scheme: "minor"
      },
      text: "This is "
    },
    {
      font: {
        italic: true,
        size: 12,
        color: { theme: 0 },
        name: "Calibri",
        scheme: "minor"
      },
      text: "a"
    },
    {
      font: {
        size: 12,
        color: { theme: 1 },
        name: "Calibri",
        family: 2,
        scheme: "minor"
      },
      text: " "
    },
    {
      font: {
        size: 12,
        color: { argb: "FFFF6600" },
        name: "Calibri",
        scheme: "minor"
      },
      text: "colorful"
    },
    {
      font: {
        size: 12,
        color: { theme: 1 },
        name: "Calibri",
        family: 2,
        scheme: "minor"
      },
      text: " text "
    },
    {
      font: {
        size: 12,
        color: { argb: "FFCCFFCC" },
        name: "Calibri",
        scheme: "minor"
      },
      text: "with"
    },
    {
      font: {
        size: 12,
        color: { theme: 1 },
        name: "Calibri",
        family: 2,
        scheme: "minor"
      },
      text: " in-cell "
    },
    {
      font: {
        bold: true,
        size: 12,
        color: { theme: 1 },
        name: "Calibri",
        family: 2,
        scheme: "minor"
      },
      text: "format"
    }
  ]
});

Stream.setCellValue(ws.getCell("D2"), "Zoo");
Stream.setCellNote(ws.getCell("D2"), "Plain Text Comment");

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
