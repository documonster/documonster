import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";
import { Workbook, Worksheet } from "@excel/index";

const [, , filename] = process.argv;

const stopwatch = new HrStopwatch();
stopwatch.start();

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "blort");

for (let row = 1; row <= 100; row++) {
  const values: string[] = [];
  if (row === 1) {
    values.push("");
    for (let col = 2; col <= 100; col++) {
      values.push(`Col ${col}`);
    }
  } else {
    for (let col = 1; col <= 100; col++) {
      if (col === 1) {
        values.push(`Row ${row}`);
      } else {
        values.push(`${row}-${col}`);
      }
    }
  }
  Worksheet.addRow(ws, values);
}

ws.pageSetup.printTitlesColumn = "A:A";
ws.pageSetup.printTitlesRow = "1:1";

try {
  await Workbook.writeXlsx(wb, filename);
  const micros = stopwatch.microseconds;
  console.log("Done.");
  console.log("Time taken:", micros);
} catch (error) {
  console.log(error.message);
}
