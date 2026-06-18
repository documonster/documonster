import { Workbook } from "@excel/index";
const workbook = Workbook.create();
Workbook.getXlsxIo(workbook)
  .readFile("./out/template.xlsx")
  .then(stream => {
    const options = {
      useSharedStrings: true,
      useStyles: true
    };

    return Workbook.writeFile(stream, "./out/template-out.xlsx", options).then(() => {
      console.log("Done.");
    });
  })
  .catch(error => {
    console.error(error.message);
    console.error(error.stack);
  });
