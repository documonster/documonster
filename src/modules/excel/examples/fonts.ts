import path from "node:path";
import { fileURLToPath } from "node:url";

import { Cell, Row, Workbook, Worksheet } from "@excel/index";

const exampleDir = path.dirname(fileURLToPath(import.meta.url));

const filename = process.argv[2] ?? path.join(exampleDir, "data/table.xlsx");

Workbook.readFile(filename)
  .then(workbook => {
    Workbook.eachWorksheet(workbook, worksheet => {
      console.log(
        `Sheet ${worksheet.id} - ${Worksheet.getName(worksheet)}, Dims=${JSON.stringify(Worksheet.getDimensions(worksheet))}`
      );
      Row.each(worksheet, rowNumber => {
        Row.eachAddress(worksheet, rowNumber, addr => {
          if (Cell.getFont(worksheet, addr)?.strike) {
            console.log(`Strikethrough: ${Cell.getValue(worksheet, addr)}`);
          }
        });
      });
    });
  })
  .catch(error => {
    console.log(error.message);
  });
