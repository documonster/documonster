// This file is typechecked by `pnpm type` but is NOT executed by Vitest.
//
// `WorksheetWriter#workbook` hands back the writer that made the sheet. It was typed as the
// structural protocol the sheet uses internally — shared strings, styles, `_openStream`,
// `packageSink` — so a caller holding `ws.workbook` could reach those and not the workbook's
// own `addWorksheet` or `commit`.

import { Stream } from "@excel/index";
import type { Stream as BrowserStream } from "@excel/index.browser";

type IsEqual<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2 ? true : false;
type Assert<T extends true> = T;

declare const ws: Stream.WorksheetWriter;
type _Workbook = Assert<IsEqual<typeof ws.workbook, Stream.WorkbookWriter>>;
ws.workbook.addWorksheet("Next");

const wb = new Stream.WorkbookWriter();
type _Added = Assert<IsEqual<ReturnType<typeof wb.addWorksheet>, Stream.WorksheetWriter>>;

declare const browserWs: BrowserStream.WorksheetWriter;
type _BrowserWorkbook = Assert<IsEqual<typeof browserWs.workbook, BrowserStream.WorkbookWriter>>;
