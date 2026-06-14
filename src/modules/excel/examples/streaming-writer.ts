import { cellSetValue } from "@excel/cell";
import { rowSetFont } from "@excel/row";

import { WorkbookWriter } from "../../../index";

const filename = process.argv[2];
const styles = {
  filename,
  useStyles: true
};
const wb = new WorkbookWriter(styles);
const ws = wb.addWorksheet("blort");

const style = {
  font: { name: "Comic Sans MS", underline: true, bold: true, size: 16 },
  alignment: { vertical: "middle" as const, horizontal: "center" as const }
};
ws.columns = [
  { header: "A1", width: 10 },
  { header: "B1", width: 20, style },
  { header: "C1", width: 30 }
];

rowSetFont(ws.getRow(2), {
  name: "Broadway",
  color: { argb: "FFFF0000" },
  outline: true,
  size: 20
});

cellSetValue(ws.getCell("A2"), "A2");
cellSetValue(ws.getCell("B2"), "B2");
cellSetValue(ws.getCell("C2"), "C2");
cellSetValue(ws.getCell("A3"), "A3");
cellSetValue(ws.getCell("B3"), "B3");
cellSetValue(ws.getCell("C3"), "C3");

wb.commit().then(() => {
  console.log("Done");
  // var wb2 = new Workbook();
  // return Workbook.readXlsxFile(wb2, './wb.test2.xlsx');
});
