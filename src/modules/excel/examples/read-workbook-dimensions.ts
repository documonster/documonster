import { Workbook, Worksheet } from "@excel/index";

const filename = process.argv[2];

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
