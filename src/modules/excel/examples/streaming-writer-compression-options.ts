import { Stream } from "@excel/index";

const filename = process.argv[2];
console.log(filename);
const optionsBestCompression = {
  filename,
  useStyles: true,
  zip: {
    zlib: { level: 9 } // Sets the compression level.
  }
};
const wb = new Stream.WorkbookWriter(optionsBestCompression);
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

const filename2 = process.argv[3];
console.log(filename2);
const optionsBestSpeed = {
  filename: filename2,
  useStyles: true,
  zip: {
    zlib: { level: 1 } // Sets the compression level.
  }
};
const wb2 = new Stream.WorkbookWriter(optionsBestSpeed);
const ws2 = wb2.addWorksheet("blort");

ws2.columns = [
  { header: "A1", width: 10 },
  { header: "B1", width: 20, style },
  { header: "C1", width: 30 }
];

Stream.setRowFont(ws2.getRow(2), {
  name: "Broadway",
  color: { argb: "FFFF0000" },
  outline: true,
  size: 20
});

Stream.setCellValue(ws2.getCell("A2"), "A2");
Stream.setCellValue(ws2.getCell("B2"), "B2");
Stream.setCellValue(ws2.getCell("C2"), "C2");
Stream.setCellValue(ws2.getCell("A3"), "A3");
Stream.setCellValue(ws2.getCell("B3"), "B3");
Stream.setCellValue(ws2.getCell("C3"), "C3");

wb2.commit().then(() => {
  console.log("Done");
  // var wb2 = new Workbook();
  // return Workbook.readFile(wb2, './wb.test2.xlsx');
});

const filename3 = process.argv[4];
console.log(filename3);
const options = {
  filename: filename3,
  useStyles: true
};
const wb3 = new Stream.WorkbookWriter(options);
const ws3 = wb3.addWorksheet("blort");

ws3.columns = [
  { header: "A1", width: 10 },
  { header: "B1", width: 20, style },
  { header: "C1", width: 30 }
];

Stream.setRowFont(ws3.getRow(2), {
  name: "Broadway",
  color: { argb: "FFFF0000" },
  outline: true,
  size: 20
});

Stream.setCellValue(ws3.getCell("A2"), "A2");
Stream.setCellValue(ws3.getCell("B2"), "B2");
Stream.setCellValue(ws3.getCell("C2"), "C2");
Stream.setCellValue(ws3.getCell("A3"), "A3");
Stream.setCellValue(ws3.getCell("B3"), "B3");
Stream.setCellValue(ws3.getCell("C3"), "C3");

wb3.commit().then(() => {
  console.log("Done");
  // var wb2 = new Workbook();
  // return Workbook.readFile(wb2, './wb.test2.xlsx');
});
