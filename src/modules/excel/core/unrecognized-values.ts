/**
 * The sheet settings a read kept because they are outside their setting's vocabulary, listed — what
 * `Workbook.readWithDiagnostics` reports as `unrecognizedValues` and `Workbook.listUnrecognizedValues` returns.
 *
 * A module of its own, with nothing but the workbook core beneath it, because `Workbook.listUnrecognizedValues` is part
 * of the workbook core that every bundle reaching Excel imports — the Word and PDF bridges included. It used to live in
 * `workbook-io-types.ts`, beside `readWorkbookStream` and its `import()` of the XLSX reader: importing it from the core
 * put that `import()` into every such bundle, where the reader's chunk outlived the tree-shaken caller and made a
 * single-file (IIFE) bundle impossible.
 */
import type { WorkbookData } from "@excel/core/workbook-core";
import { getWorksheets } from "@excel/core/workbook-core";
import { getSheetName } from "@excel/core/worksheet-core";
import { findKeptSheetSettings } from "@excel/utils/sheet-settings-checks";

/**
 * A value a read kept because it is outside its setting's vocabulary: the sheet it is on (absent for a workbook-level
 * setting), the setting's path (`pageSetup.orientation`, `views[0].zoomScale`) and the text the file gave.
 */
export interface UnrecognizedValue {
  readonly sheet?: string;
  readonly path: string;
  readonly value: string;
}

/**
 * The settings a read kept on `wb`'s sheets, as they are kept now — the `unrecognizedValues` of a report, and
 * `Workbook.listUnrecognizedValues`. Fresh objects.
 */
export function keptSettingsReport(wb: WorkbookData): UnrecognizedValue[] {
  const entries: UnrecognizedValue[] = [];
  for (const ws of getWorksheets(wb)) {
    for (const { path, value } of findKeptSheetSettings(ws)) {
      entries.push({ sheet: getSheetName(ws), path, value });
    }
  }
  return entries;
}
