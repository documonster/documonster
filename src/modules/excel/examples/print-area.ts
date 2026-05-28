import { Workbook } from "../../../index";

const [, , filename] = process.argv;

const wb = new Workbook();
const ws = wb.addWorksheet("test sheet");

for (let row = 1; row <= 10; row++) {
  const values: string[] = [];
  if (row === 1) {
    values.push("");
    for (let col = 2; col <= 10; col++) {
      values.push(`Col ${col}`);
    }
  } else {
    for (let col = 1; col <= 10; col++) {
      if (col === 1) {
        values.push(`Row ${row}`);
      } else {
        values.push(`${row}-${col}`);
      }
    }
  }
  ws.addRow(values);
}

ws.pageSetup.printTitlesColumn = "A:A";
ws.pageSetup.printTitlesRow = "1:1";
ws.pageSetup.printArea = "A1:B10";

try {
  await wb.xlsx.writeFile(filename);
  console.log("Done.");
} catch (error) {
  console.log(error.message);
}
