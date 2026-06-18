import { Stream } from "@excel/index";

const filename = process.argv[2];
const styles = {
  filename,
  useStyles: true
};
const wb = new Stream.WorkbookWriter(styles);
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

Stream.setRowFont(ws.getRow(2), {
  name: "Broadway",
  color: { argb: "FFFF0000" },
  outline: true,
  size: 20
});

Stream.setCellValue(ws.getCell("A2"), "A2");
Stream.setCellValue(ws.getCell("B2"), "B2");
Stream.setCellValue(ws.getCell("C2"), "C2");
Stream.setCellValue(ws.getCell("A3"), "A3");
Stream.setCellValue(ws.getCell("B3"), "B3");
Stream.setCellValue(ws.getCell("C3"), "C3");

wb.commit().then(() => {
  console.log("Done");
  // var wb2 = new Workbook();
  // return Workbook.readFile(wb2, './wb.test2.xlsx');
});
