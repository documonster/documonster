import { Address, Cell, Row, Workbook, Worksheet } from "@excel/index";

const filename = process.argv[2];

const workbook = Workbook.create();
Workbook.getXlsxIo(workbook)
  .readFile(filename)
  .then(() => {
    Workbook.eachSheet(workbook, worksheet => {
      console.log(
        `Sheet ${worksheet.id} - ${Worksheet.getName(worksheet)}, Dims=${JSON.stringify(Worksheet.dimensions(worksheet))}`
      );
      Worksheet.eachRow(worksheet, row => {
        Row.eachCell(worksheet, row.number, (_cell, colNumber) => {
          const addr = `${Address.encodeCol(colNumber - 1)}${row.number}`;
          if (Cell.getFont(worksheet, addr)!.strike) {
            console.log(`Strikethrough: ${Cell.getValue(worksheet, addr)}`);
          }
        });
      });
    });
  })
  .catch(error => {
    console.log(error.message);
  });
