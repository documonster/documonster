import { HrStopwatch } from "@excel/examples/utils/hr-stopwatch";
import { Cell, Workbook, Worksheet } from "@excel/index";

const [, , filename, password] = process.argv;

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "Foo");
Cell.setValue(ws, "A1", "Bar");

async function save() {
  const stopwatch = new HrStopwatch();

  stopwatch.start();
  await Worksheet.protect(ws, password); // default 100000
  console.log("Protection Time [spinCount default]:", stopwatch.microseconds);

  await Workbook.writeXlsx(wb, `${0}-${filename}`);

  // options defined but spinCount not
  stopwatch.start();
  await Worksheet.protect(ws, password, { insertRows: true }); // default 100000
  console.log("Protection Time [spinCount default]:", stopwatch.microseconds);

  await Workbook.writeXlsx(wb, `${1}-${filename}`);

  const values = [100000, 10000, 1, 0, -1, undefined, null, NaN, Infinity, -Infinity, 31415.9265];

  for (let index = 0; index < values.length; index += 1) {
    const value = values[index];

    stopwatch.start();
    await Worksheet.protect(ws, password, { spinCount: value ?? undefined });
    console.log(`Protection Time [spinCount ${value}]:`, stopwatch.microseconds);

    await Workbook.writeXlsx(wb, `${index + 2}-${filename}`);
  }
}

try {
  await save();
} catch (error) {
  console.log((error as Error).message);
}
