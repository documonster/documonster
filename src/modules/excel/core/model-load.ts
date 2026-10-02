import type { DefinedNameModel, DefinedNamesData } from "@excel/core/defined-names";
import { definedNamesSetModel } from "@excel/core/defined-names";
import type { WorkbookData } from "@excel/core/workbook-core";
import type { WorkbookModel } from "@excel/core/workbook-model";
/**
 * Applying a model whose defined names may arrive as raw text — the one place the default
 * formula-syntax probe is bound.
 *
 * Classifying a defined name ("is `OFFSET(Sheet1!$A$1,0,0,3,1)` a formula or opaque text?") needs the
 * formula tokenizer and parser. Only *loading* a model does that, so the probe is supplied by whoever
 * loads — the XLSX and XLSB readers, and the public `Workbook.setModel` / `DefinedNames.setModel` —
 * rather than imported by `defined-names.ts` and `workbook.browser.ts`, which every writer reaches.
 *
 * Applying a whole workbook model lives here for the same reason. A webpack-family bundler (rspack)
 * decides which exports of a module are used across the *whole* build, so an export the lazily loaded
 * reader calls keeps its dependencies wherever its module is — and `workbook.browser.ts` is in the entry
 * chunk of every write-only export. While `setWorkbookModel` lived there it kept the tokenizer, the parser
 * and, through `setSheetModel`, the chart handle in that chunk.
 */
import { createWorksheet } from "@excel/core/worksheet";
import { setSheetModel } from "@excel/core/worksheet-load";
import { parse } from "@formula/syntax/parser";
import { tokenize } from "@formula/syntax/tokenizer";

/**
 * The built-in formula-syntax probe, backed by the real tokenizer + parser. Used whenever a defined-names
 * record was created without an explicit `formulaSyntaxProbe`.
 */
export function defaultFormulaSyntaxProbe(text: string): boolean {
  try {
    const tokens = tokenize(text);
    if (tokens.length === 0) {
      return false;
    }
    parse(tokens);
    return true;
  } catch {
    return false;
  }
}

/** `definedNamesSetModel` with the built-in probe as the fallback — `DefinedNames.setModel`. */
export function loadDefinedNamesModel(dn: DefinedNamesData, value: DefinedNameModel[]): void {
  definedNamesSetModel(dn, value, defaultFormulaSyntaxProbe);
}

/**
 * Replace the workbook's contents with `value` — `Workbook.setModel`, and the last step of both readers.
 *
 * Defined names are classified with the workbook's own probe when it was created with one, and with
 * {@link defaultFormulaSyntaxProbe} otherwise.
 */
export function loadWorkbookModel(wb: WorkbookData, value: WorkbookModel): void {
  wb.creator = value.creator;
  wb.lastModifiedBy = value.lastModifiedBy;
  wb.lastPrinted = value.lastPrinted;
  wb.created = value.created;
  wb.modified = value.modified;
  wb.company = value.company;
  wb.manager = value.manager;
  wb.title = value.title;
  wb.subject = value.subject;
  wb.keywords = value.keywords;
  wb.category = value.category;
  wb.description = value.description;
  wb.language = value.language;
  wb.revision = value.revision;
  wb.contentStatus = value.contentStatus;
  wb.workbookContentType = value.workbookContentType;

  wb.properties = value.properties;
  wb.protection = value.protection;
  wb.calcProperties = value.calcProperties;
  wb._worksheets = [];
  wb._tableNames.clear();
  value.worksheets.forEach(worksheetModel => {
    const { id, name, state } = worksheetModel;
    // API invariant: `_worksheets` is keyed by a positive integer
    // sheet id. A worksheet model with a missing or non-integer id
    // would be stored under a string pseudo key like `"undefined"`
    // or `"NaN"`, making it unreachable via `getWorksheet(name)`.
    // The xlsx reconciler enforces the same invariant
    // before reaching this point; programmatic callers assigning
    // `model` directly with a malformed payload land here instead.
    if (!Number.isInteger(id) || (id as number) <= 0) {
      return;
    }
    const orderNo = value.sheets && value.sheets.findIndex(ws => ws.id === id);
    const worksheet = (wb._worksheets[id] = createWorksheet({
      id,
      name,
      orderNo: orderNo !== -1 ? orderNo : undefined,
      state,
      workbook: wb
    }));
    setSheetModel(worksheet, worksheetModel);
  });

  definedNamesSetModel(wb._definedNames, value.definedNames, defaultFormulaSyntaxProbe);
  wb.views = value.views;
  wb._themes = value.themes;
  wb.media = value.media || [];

  // Handle pivot tables - either newly created or loaded from file
  // Loaded pivot tables come from loadedPivotTables after reconciliation
  wb.pivotTables = value.pivotTables || value.loadedPivotTables || [];

  // Preserve default font for round-trip fidelity
  wb._defaultFont = value.defaultFont;
  // Restore named cell styles for round-trip fidelity
  wb._cellStyles = value.cellStyles
    ? new Map(value.cellStyles.map(cs => [cs.name, cs]))
    : undefined;
  wb._dxfs = value.dxfs;
  // Restore chart entries
  wb._chartEntries = value.chartEntries || {};
  wb._chartRels = value.chartRels || {};
  wb._chartStyles = value.chartStyles || {};
  wb._chartColors = value.chartColors || {};
  wb._chartExStyles = (value as { chartExStyles?: Record<number, Uint8Array> }).chartExStyles || {};
  wb._chartExColors = (value as { chartExColors?: Record<number, Uint8Array> }).chartExColors || {};
  wb._chartExEntries = value.chartExEntries || {};
  wb._chartExRels = value.chartExRels || {};
  wb._chartExStructuredEntries = value.chartExStructuredEntries || {};
  // Restore chartsheets. Populate each chartsheet's `orderNo` from
  // the position in `value.sheets` (workbook.xml tab order) so the
  // writer's `prepare()` can sort interleaved worksheets +
  // chartsheets back into the author's layout. Matches the
  // equivalent loop above for worksheets.
  wb._chartsheets = value.chartsheets || [];
  if (value.sheets) {
    for (const cs of wb._chartsheets) {
      const idx = value.sheets.findIndex((s: { id?: number }) => s.id === cs.id);
      if (idx !== -1) {
        cs.orderNo = idx;
      }
    }
  }
  // Restore threaded-comment person directory. Always assign a new
  // list so callers editing the previous value don't mutate the
  // newly-loaded workbook by accident.
  wb._persons = value.persons ? [...value.persons] : [];
  // Restore raw-passthrough slicer/timeline parts so dashboards
  // survive round-trip. The maps are stored by reference — loaders
  // and writers treat them as read-only; mutating them between
  // load and save is not supported.
  wb._slicerParts = value.slicerParts ?? {};
  wb._slicerCacheParts = value.slicerCacheParts ?? {};
  wb._timelineParts = value.timelineParts ?? {};
  wb._timelineCacheParts = value.timelineCacheParts ?? {};
  wb._opaqueParts = value.opaqueParts ?? [];
  wb._xlsbPivotCaches = value.xlsbPivotCaches;
  wb._opaqueContentTypeDefaults = value.opaqueContentTypeDefaults ?? {};
  wb._opaqueDrops = value.opaqueDrops ?? [];
  // Preserve external workbook references (empty array if none)
  wb.externalLinks = value.externalLinks ? [...value.externalLinks] : [];
  // Reset the writer-scoped auto-discovery cache — loading a fresh
  // workbook replaces any accumulated state from previous writes.
  wb._writerExternalLinkCache = new Map();
}
