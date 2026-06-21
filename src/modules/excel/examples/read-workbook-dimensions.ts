import path from "node:path";
import { fileURLToPath } from "node:url";

import { Workbook, Worksheet } from "@excel/index";

const exampleDir = path.dirname(fileURLToPath(import.meta.url));

const filename = process.argv[2] ?? path.join(exampleDir, "data/table.xlsx");

const workbook = Workbook.create();
Workbook.getXlsxIo(workbook)
  .readFile(filename)
  .then(() => {
    Workbook.eachSheet(workbook, worksheet => {
      console.log(
        `Sheet ${worksheet.id} - ${Worksheet.getName(worksheet)}, Dims=${JSON.stringify(Worksheet.dimensions(worksheet))}`
      );
    });
  })
  .catch(error => {
    console.log(error.message);
  });
