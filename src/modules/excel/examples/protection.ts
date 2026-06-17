import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";
import { Cell, Workbook, Worksheet } from "@excel/index";

const [, , filename, password] = process.argv;

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "Foo");
Cell.setValue(ws, "A1", 1);
Cell.setValue(ws, "B1", 2);
Cell.setValue(ws, "A2", { formula: "A1+2", result: 3 });
Cell.setValue(ws, "B2", { formula: "B1+2", result: 4 });

Cell.setStyle(ws, "B1", { protection: { locked: false } });
Cell.setStyle(ws, "A2", { protection: { locked: false } });
Cell.setStyle(ws, "B2", { protection: { hidden: true } });

async function save() {
  const stopwatch = new HrStopwatch();
  stopwatch.start();

  await Worksheet.protect(ws, password);
  console.log("Protection Time:", stopwatch.microseconds);

  stopwatch.start();
  await Workbook.writeXlsx(wb, filename);
  console.log("Done.");
  console.log("Time taken:", stopwatch.microseconds);
}

save().catch(error => {
  console.log(error.message);
});
