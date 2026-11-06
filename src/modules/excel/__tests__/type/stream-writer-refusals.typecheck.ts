// This file is typechecked by `pnpm type` but is NOT executed by Vitest.
//
// A namespace function that needs rows already written, or the whole document, refuses a streaming writer at
// runtime with `ExcelUnsupportedFeatureError` (`refuseStreamWorksheet`, `refuseStreamWorkbook`, `workbookOnly`).
// The runtime refusal is for JavaScript callers and `any`; a typed caller must hear it from the compiler first.
// Each function below is refused at runtime, so its first parameter must not accept a writer — and the controls,
// which do take one, prove the check can tell the two apart.

import type { Column, Image, Row, Stream, Workbook, Worksheet } from "@excel/index";

type Assert<T extends true> = T;
type First<F> = F extends (first: infer A, ...rest: never[]) => unknown ? A : never;
type Refuses<F, W> = [W] extends [First<F>] ? false : true;

type WsW = Stream.WorksheetWriter;
type WbW = Stream.WorkbookWriter;

export type RefusedOnWorksheetWriter = [
  Assert<Refuses<typeof Row.insert, WsW>>,
  Assert<Refuses<typeof Row.insertMany, WsW>>,
  Assert<Refuses<typeof Row.duplicate, WsW>>,
  Assert<Refuses<typeof Row.splice, WsW>>,
  Assert<Refuses<typeof Row.getCount, WsW>>,
  Assert<Refuses<typeof Row.getActualCount, WsW>>,
  Assert<Refuses<typeof Row.getLast, WsW>>,
  Assert<Refuses<typeof Row.each, WsW>>,
  Assert<Refuses<typeof Row.autoFit, WsW>>,
  Assert<Refuses<typeof Column.getCount, WsW>>,
  Assert<Refuses<typeof Column.getActualCount, WsW>>,
  Assert<Refuses<typeof Column.has, WsW>>,
  Assert<Refuses<typeof Column.getLast, WsW>>,
  Assert<Refuses<typeof Column.splice, WsW>>,
  Assert<Refuses<typeof Column.each, WsW>>,
  Assert<Refuses<typeof Column.eachAddress, WsW>>,
  Assert<Refuses<typeof Column.autoFit, WsW>>,
  Assert<Refuses<typeof Column.getValues, WsW>>,
  Assert<Refuses<typeof Column.setValues, WsW>>,
  Assert<Refuses<typeof Image.addShape, WsW>>,
  Assert<Refuses<typeof Image.removeShape, WsW>>,
  Assert<Refuses<typeof Image.listShapes, WsW>>,
  Assert<Refuses<typeof Worksheet.unmerge, WsW>>
];

export type RefusedOnWorkbookWriter = [
  Assert<Refuses<typeof Workbook.registerPerson, WbW>>,
  Assert<Refuses<typeof Workbook.addExternalLink, WbW>>,
  Assert<Refuses<typeof Workbook.removeWorksheet, WbW>>,
  Assert<Refuses<typeof Workbook.addChartsheet, WbW>>,
  Assert<Refuses<typeof Workbook.addPivotChartsheet, WbW>>,
  Assert<Refuses<typeof Workbook.getChartsheet, WbW>>,
  Assert<Refuses<typeof Workbook.listChartsheets, WbW>>,
  Assert<Refuses<typeof Workbook.removeChartsheet, WbW>>,
  Assert<Refuses<typeof Workbook.copyChartsheet, WbW>>,
  Assert<Refuses<typeof Workbook.replaceChartsheetChart, WbW>>,
  Assert<Refuses<typeof Workbook.renameChartsheet, WbW>>,
  Assert<Refuses<typeof Workbook.importWorksheet, WbW>>,
  Assert<Refuses<typeof Workbook.listUnrecognizedValues, WbW>>,
  Assert<Refuses<typeof Workbook.discardUnrecognizedValues, WbW>>,
  Assert<Refuses<typeof Workbook.getExternalLink, WbW>>,
  Assert<Refuses<typeof Workbook.listExternalLinks, WbW>>,
  Assert<Refuses<typeof Workbook.listPersons, WbW>>,
  Assert<Refuses<typeof Workbook.registerFunction, WbW>>,
  Assert<Refuses<typeof Workbook.unregisterFunction, WbW>>,
  Assert<Refuses<typeof Workbook.getModel, WbW>>,
  Assert<Refuses<typeof Workbook.setModel, WbW>>,
  Assert<Refuses<typeof Workbook.removeCellStyle, WbW>>,
  Assert<Refuses<typeof Workbook.toBuffer, WbW>>,
  Assert<Refuses<typeof Workbook.toStream, WbW>>,
  Assert<Refuses<typeof Workbook.writeFile, WbW>>,
  Assert<Refuses<typeof Workbook.writeStream, WbW>>
];

// Controls: these take a writer, so `Refuses` must answer `false` for them.
export type AcceptedControls = [
  Assert<[Refuses<typeof Row.setHeight, WsW>] extends [false] ? true : false>,
  Assert<[Refuses<typeof Column.setWidth, WsW>] extends [false] ? true : false>
];
