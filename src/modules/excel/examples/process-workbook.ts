import { Cell, Workbook, Worksheet } from "@excel/index";

const [, , inputFile, outputFile] = process.argv;

const wb = Workbook.create();

let passed = true;
function assert(value, failMessage, passMessage) {
  if (!value) {
    if (failMessage) {
      console.error(failMessage);
    }
    passed = false;
  } else if (passMessage) {
    console.log(passMessage);
  }
}

// assuming file created by testBookOut
Workbook.getXlsxIo(wb)
  .readFile(inputFile)
  .then(() => {
    console.log("Loaded", inputFile);

    Workbook.eachSheet(wb, sheet => {
      console.log(Worksheet.getName(sheet));
    });

    const ws = Workbook.getWorksheet(wb, "Sheet1")!;

    assert(ws, "Expected to find a worksheet called sheet1", "");

    Cell.setValue(ws!, "B1", new Date());
    Cell.setStyle(ws!, "B1", { numFmt: "hh:mm:ss" });

    Worksheet.addRow(ws!, [1, "hello"]);
    return Workbook.writeFile(wb, outputFile);
  })
  .then(() => {
    assert(passed, "Something went wrong", "All tests passed!");
  })
  .catch(error => {
    console.error(error.message);
    console.error(error.stack);
  });
