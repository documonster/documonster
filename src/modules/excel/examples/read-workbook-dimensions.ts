import { Workbook, Worksheet } from "@excel/index";
import { getXlsxIo } from "@excel/workbook";
import { getSheetName } from "@excel/worksheet";

const filename = process.argv[2];

const workbook = Workbook.create();
getXlsxIo(workbook)
  .readFile(filename)
  .then(() => {
    Workbook.eachSheet(workbook, worksheet => {
      console.log(
        `Sheet ${worksheet.id} - ${getSheetName(worksheet)}, Dims=${JSON.stringify(Worksheet.dimensions(worksheet))}`
      );
    });
  })
  .catch(error => {
    console.log(error.message);
  });
