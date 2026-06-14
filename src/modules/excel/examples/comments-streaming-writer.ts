import { cellSetNote, cellSetValue } from "@excel/cell";
import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";

import { WorkbookWriter } from "../../../index";

const [, , filename] = process.argv;

const wb = new WorkbookWriter({ filename });
const ws = wb.addWorksheet("Foo");
cellSetValue(ws.getCell("B2"), 5);
cellSetNote(ws.getCell("B2"), {
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

cellSetValue(ws.getCell("D2"), "Zoo");
cellSetNote(ws.getCell("D2"), "Plain Text Comment");

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
