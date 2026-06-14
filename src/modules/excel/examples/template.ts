import { Workbook } from "@excel/index";
import { getXlsxIo } from "@excel/workbook";
const workbook = Workbook.create();
getXlsxIo(workbook)
  .readFile("./out/template.xlsx")
  .then(stream => {
    const options = {
      useSharedStrings: true,
      useStyles: true
    };

    return Workbook.writeXlsx(stream, "./out/template-out.xlsx", options).then(() => {
      console.log("Done.");
    });
  })
  .catch(error => {
    console.error(error.message);
    console.error(error.stack);
  });
