/**
 * Every value kind and style the streaming writer can put down, in both containers.
 *
 * Written through `streamBothFormats`, which runs the build once per format — a streamed workbook is not a value
 * that can be written twice, so the description is what gets reused. `format: "xlsb"` sends each committed row
 * through the same record encoder the buffered binary writer uses.
 *
 * Run: pnpm example --filter streaming-writer-basic
 */
import path from "node:path";
import { fileURLToPath } from "node:url";

import { streamBothFormats } from "@excel/examples/utils/stream-both";
import { Cell, Column, Row, Stream, Workbook, Worksheet } from "@excel/index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/excel-examples"
);
const filename = process.argv[2] ?? path.join(outDir, "streaming-writer-basic.xlsx");

await streamBothFormats(filename, { useSharedStrings: false, useStyles: false }, wb => {
  const ws = Workbook.addWorksheet(wb, "blort");

  const fonts = {
    arialBlackUI14: {
      name: "Arial Black",
      family: 2,
      size: 14,
      underline: true,
      italic: true
    },
    comicSansUdB16: {
      name: "Comic Sans MS",
      family: 4,
      size: 16,
      underline: "double" as const,
      bold: true
    }
  };

  const alignments = [
    { text: "Top Left", alignment: { horizontal: "left" as const, vertical: "top" as const } },
    {
      text: "Middle Centre",
      alignment: { horizontal: "center" as const, vertical: "middle" as const }
    },
    {
      text: "Bottom Right",
      alignment: { horizontal: "right" as const, vertical: "bottom" as const }
    },
    { text: "Wrap Text", alignment: { wrapText: true } },
    { text: "Indent 1", alignment: { indent: 1 } },
    { text: "Indent 2", alignment: { indent: 2 } },
    {
      text: "Rotate 15",
      alignment: { horizontal: "right" as const, vertical: "bottom" as const, textRotation: 15 }
    },
    {
      text: "Rotate 30",
      alignment: { horizontal: "right" as const, vertical: "bottom" as const, textRotation: 30 }
    },
    {
      text: "Rotate 45",
      alignment: { horizontal: "right" as const, vertical: "bottom" as const, textRotation: 45 }
    },
    {
      text: "Rotate 60",
      alignment: { horizontal: "right" as const, vertical: "bottom" as const, textRotation: 60 }
    },
    {
      text: "Rotate 75",
      alignment: { horizontal: "right" as const, vertical: "bottom" as const, textRotation: 75 }
    },
    {
      text: "Rotate 90",
      alignment: { horizontal: "right" as const, vertical: "bottom" as const, textRotation: 90 }
    },
    {
      text: "Rotate -15",
      alignment: { horizontal: "right" as const, vertical: "bottom" as const, textRotation: -55 }
    },
    {
      text: "Rotate -30",
      alignment: { horizontal: "right" as const, vertical: "bottom" as const, textRotation: -30 }
    },
    {
      text: "Rotate -45",
      alignment: { horizontal: "right" as const, vertical: "bottom" as const, textRotation: -45 }
    },
    {
      text: "Rotate -60",
      alignment: { horizontal: "right" as const, vertical: "bottom" as const, textRotation: -60 }
    },
    {
      text: "Rotate -75",
      alignment: { horizontal: "right" as const, vertical: "bottom" as const, textRotation: -75 }
    },
    {
      text: "Rotate -90",
      alignment: { horizontal: "right" as const, vertical: "bottom" as const, textRotation: -90 }
    },
    {
      text: "Vertical Text",
      alignment: {
        horizontal: "right" as const,
        vertical: "bottom" as const,
        textRotation: "vertical" as const
      }
    }
  ];
  // const badAlignments = [
  //   { text: 'Rotate -91', alignment: { textRotation: -91 } },
  //   { text: 'Rotate 91', alignment: { textRotation: 91 } },
  //   { text: 'Indent -1', alignment: { indent: -1 } },
  //   { text: 'Blank', alignment: {} },
  // ];

  const borders = {
    thin: {
      top: { style: "thin" as const },
      left: { style: "thin" as const },
      bottom: { style: "thin" as const },
      right: { style: "thin" as const }
    },
    doubleRed: {
      color: { argb: "FFFF0000" },
      top: { style: "double" as const },
      left: { style: "double" as const },
      bottom: { style: "double" as const },
      right: { style: "double" as const }
    },
    thickRainbow: {
      top: { style: "double" as const, color: { argb: "FFFF00FF" } },
      left: { style: "double" as const, color: { argb: "FF00FFFF" } },
      bottom: { style: "double" as const, color: { argb: "FF00FF00" } },
      right: { style: "double" as const, color: { argb: "FF00FF" } },
      diagonal: {
        style: "double" as const,
        color: { argb: "FFFFFF00" },
        up: true,
        down: true
      }
    }
  };

  const fills = {
    redDarkVertical: {
      type: "pattern" as const,
      pattern: "darkVertical" as const,
      fgColor: { argb: "FFFF0000" }
    },
    redGreenDarkTrellis: {
      type: "pattern" as const,
      pattern: "darkTrellis" as const,
      fgColor: { argb: "FFFF0000" },
      bgColor: { argb: "FF00FF00" }
    },
    blueWhiteHGrad: {
      type: "gradient" as const,
      gradient: "angle" as const,
      degree: 0,
      stops: [
        { position: 0, color: { argb: "FF0000FF" } },
        { position: 1, color: { argb: "FFFFFFFF" } }
      ]
    },
    rgbPathGrad: {
      type: "gradient" as const,
      gradient: "path" as const,
      center: { left: 0.5, top: 0.5 },
      stops: [
        { position: 0, color: { argb: "FFFF0000" } },
        { position: 0.5, color: { argb: "FF00FF00" } },
        { position: 1, color: { argb: "FF0000FF" } }
      ]
    }
  };

  Column.setDefinitions(ws, [
    { header: "Col 1", key: "key", width: 25 },
    { header: "Col 2", key: "name", width: 32 },
    { header: "Col 3", key: "age", width: 21 },
    { header: "Col 4", key: "addr1", width: 18 },
    { header: "Col 5", key: "addr2", width: 8 },
    { header: "Col 6", width: 8 },
    { header: "Col 7", width: 8 },
    {
      header: "Col 8",
      width: 8,
      style: { font: fonts.comicSansUdB16, alignment: alignments[1].alignment }
    }
  ]);

  Cell.setValue(ws, "A2", 7);
  Cell.setValue(ws, "B2", "Hello, World!");
  Cell.setFont(ws, "B2", fonts.comicSansUdB16);
  Cell.setBorder(ws, "B2", borders.thin);

  Cell.setValue(ws, "C2", -5.55);
  Cell.setNumFmt(ws, "C2", "'£'#,##0.00;[Red]-'£'#,##0.00");
  Cell.setFont(ws, "C2", fonts.arialBlackUI14);

  Cell.setValue(ws, "D2", 3.14);
  Cell.setValue(ws, "D2", new Date());
  Cell.setNumFmt(ws, "D2", "d-mmm-yyyy");
  Cell.setFont(ws, "D2", fonts.comicSansUdB16);
  Cell.setBorder(ws, "D2", borders.doubleRed);

  Cell.setValue(ws, "E2", `${["Hello", "World"].join(", ")}!`);
  Stream.commitRow(ws.getRow(2));

  Cell.setValue(ws, "A3", {
    text: "www.google.com",
    hyperlink: "http://www.google.com"
  });
  Cell.setValue(ws, "A4", "Boo!");
  Cell.setValue(ws, "C4", "Hoo!");
  Worksheet.merge(ws, "A4", "C4");
  Stream.commitRow(ws.getRow(4));

  Cell.setValue(ws, "A5", 1);
  Cell.setValue(ws, "B5", 2);
  Cell.setValue(ws, "C5", { formula: "A5+B5", result: 3 });
  Stream.commitRow(ws.getRow(5));

  Cell.setValue(ws, "A6", "Hello");
  Cell.setValue(ws, "B6", "World");
  Cell.setValue(ws, "C6", {
    formula: 'CONCATENATE(A6,", ",B6,"!")',
    result: "Hello, World!"
  });
  Cell.setBorder(ws, "C6", borders.thickRainbow);
  Stream.commitRow(ws.getRow(6));

  Cell.setValue(ws, "A7", 1);
  Cell.setValue(ws, "B7", 2);
  Cell.setValue(ws, "C7", { formula: "A7+B7" });
  Stream.commitRow(ws.getRow(7));

  const now = new Date();
  Cell.setValue(ws, "A8", now);
  Cell.setValue(ws, "B8", 0);
  Cell.setValue(ws, "C8", { formula: "A8+B8", result: now });
  Stream.commitRow(ws.getRow(8));

  Cell.setValue(ws, "A9", 1.6);
  Cell.setNumFmt(ws, "A9", "# ?/?");
  Cell.setValue(ws, "B9", 1.6);
  Cell.setNumFmt(ws, "B9", "h:mm:ss");
  Cell.setValue(ws, "C9", 0.016);
  Cell.setNumFmt(ws, "C9", "0.00%");
  Cell.setValue(ws, "D9", 1.6);
  Cell.setNumFmt(ws, "D9", "[Green]#,##0 ;[Red](#,##0)");
  Cell.setValue(ws, "E9", 1.6);
  Cell.setNumFmt(ws, "E9", "#0.000");
  Cell.setValue(ws, "F9", 0.016);
  Cell.setNumFmt(ws, "F9", "# ?/?%");
  Stream.commitRow(ws.getRow(9));

  Cell.setValue(ws, "A10", "<");
  Cell.setValue(ws, "B10", ">");
  Cell.setValue(ws, "C10", "<a>");
  Cell.setValue(ws, "D10", "><");
  Stream.commitRow(ws.getRow(10));

  ws.getRow(11).height = 40;
  alignments.forEach((alignment, index) => {
    const rowNumber = 11;
    const colNumber = index + 1;
    const cell = ws.getCell(rowNumber, colNumber);
    Stream.setCellValue(cell, alignment.text);
    Stream.setCellAlignment(cell, alignment.alignment);
  });
  Stream.commitRow(ws.getRow(11));

  const row12 = ws.getRow(12);
  row12.height = 40;
  Stream.setCellValue(Stream.getRowCell(row12, 1), "Blue White Horizontal Gradient");
  Stream.setCellFill(Stream.getRowCell(row12, 1), fills.blueWhiteHGrad);
  Stream.setCellValue(Stream.getRowCell(row12, 2), "Red Dark Vertical");
  Stream.setCellFill(Stream.getRowCell(row12, 2), fills.redDarkVertical);
  Stream.setCellValue(Stream.getRowCell(row12, 3), "Red Green Dark Trellis");
  Stream.setCellFill(Stream.getRowCell(row12, 3), fills.redGreenDarkTrellis);
  Stream.setCellValue(Stream.getRowCell(row12, 4), "RGB Path Gradient");
  Stream.setCellFill(Stream.getRowCell(row12, 4), fills.rgbPathGrad);

  // row and column styles
  Row.setFont(ws, 13, fonts.arialBlackUI14);
  Cell.setValue(ws, "H12", "Foo");
  Cell.setValue(ws, "G13", "Foo");
  Cell.setValue(ws, "H13", "Bar");
  Cell.setValue(ws, "I13", "Baz");
  Cell.setValue(ws, "H14", "Baz");
  // Stream.commitRow(ws.getRow(13));

  ws.commit();
});

console.log("Done.");
