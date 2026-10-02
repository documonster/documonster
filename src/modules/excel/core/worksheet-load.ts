import { createChart } from "@excel/chart/chart-handle";
import type { ChartAnchorModel } from "@excel/chart/model/types";
import { createDataValidations } from "@excel/core/data-validations";
import { formCheckboxFromModel } from "@excel/core/form-control";
import { imageCreate } from "@excel/core/image";
import { rowCreate } from "@excel/core/row";
// Chart runtime is imported directly (static). The chart modules depend only
// on the `*-core` data layer (never on this heavy `worksheet.ts`), so the
// dependency graph stays acyclic: `worksheet → chart → *-core`. A consumer
// that never references a chart API gets the entire chart implementation
// tree-shaken out by the bundler — no host registry / install step required.
import type { TableData, TableModel } from "@excel/core/table";
import { createTable, tableSetModel } from "@excel/core/table";
/**
 * Applying a worksheet model — the load half of `getSheetModel` / `setSheetModel`.
 *
 * A module of its own, not a member of `worksheet.ts`, because what it reaches is what a loaded sheet can
 * hold: charts (`chart-handle`), form controls, merges. `worksheet.ts` is in every writer's bundle, and the
 * XLSX and XLSB readers call this from a lazily loaded chunk. A webpack-family bundler (rspack) decides
 * which exports of a module are used across the *whole* build, so while this lived in `worksheet.ts` the
 * reader's use of it kept the chart handle in the entry chunk of a write-only export.
 */
import type { WorksheetModel } from "@excel/core/worksheet";
import { mergeCellsWithoutStyle, setSheetName } from "@excel/core/worksheet";
import type { WorksheetData } from "@excel/core/worksheet-core";
import {
  columnFromModel,
  getSheetName,
  getSheetWorkbook,
  rowSetModel
} from "@excel/core/worksheet-core";
import type { HeaderFooter } from "@excel/types";

// Type for data validation model - maps address to validation

export function setSheetModel(ws: WorksheetData, value: WorksheetModel): void {
  if (value.styledBlankRanges !== undefined) {
    (ws as { _styledBlankRanges?: readonly unknown[] })._styledBlankRanges =
      value.styledBlankRanges;
  }
  setSheetName(ws, value.name);
  ws.state = value.state;
  ws._columns = columnFromModel(ws, value.cols ?? []);
  _parseRows(ws, value);

  _parseMergeCells(ws, value);
  ws.dataValidations = createDataValidations(value.dataValidations);
  ws.properties = value.properties;
  ws.pageSetup = value.pageSetup;
  // Re-apply the OOXML schema defaults: `scaleWithDoc` / `alignWithMargins`
  // default to true and are omitted from the XML when unset, so a parsed
  // model carries only the fields the file actually wrote. Filling them here
  // keeps `HeaderFooter` fully populated for every source (fresh, cloned,
  // parsed) instead of leaking `undefined` through the public shape.
  const headerFooter = (value.headerFooter ?? {}) as Partial<HeaderFooter>;
  ws.headerFooter = {
    differentFirst: headerFooter.differentFirst ?? false,
    differentOddEven: headerFooter.differentOddEven ?? false,
    scaleWithDoc: headerFooter.scaleWithDoc ?? true,
    alignWithMargins: headerFooter.alignWithMargins ?? true,
    oddHeader: headerFooter.oddHeader ?? null,
    oddFooter: headerFooter.oddFooter ?? null,
    evenHeader: headerFooter.evenHeader ?? null,
    evenFooter: headerFooter.evenFooter ?? null,
    firstHeader: headerFooter.firstHeader ?? null,
    firstFooter: headerFooter.firstFooter ?? null
  };
  ws.rowBreaks = value.rowBreaks ?? [];
  ws.colBreaks = value.colBreaks ?? [];
  ws.views = value.views;
  ws.autoFilter = value.autoFilter;
  ws._autoFilterCriteria = value.autoFilterCriteria;
  ws._sortStateXml = value.sortStateXml;
  ws._opaqueRels = value.opaqueRels;
  ws._chartsheetPlaceholder = value.chartsheetPlaceholder;
  ws._xlsbDrawingRelationshipId = value.xlsbDrawingRelationshipId;
  ws._worksheetNamespaceAttributes = value.worksheetNamespaceAttributes;
  ws._worksheetMcIgnorable = value.worksheetMcIgnorable;
  ws._sortStateAutoFilterRef = value.sortStateAutoFilterRef;
  ws._media = value.media.map(medium => imageCreate(ws, medium));
  ws._shapes = value.shapes ? value.shapes.slice() : [];
  // Restore watermark state from media entries
  ws._watermark = value.watermark ?? null;
  ws._watermarkMedia = null;
  if (!ws._watermark) {
    for (const medium of ws._media) {
      if (medium.type === "watermark") {
        ws._watermark = {
          imageId: medium.imageId ?? "",
          mode: "overlay",
          opacity: medium.opacity
        };
        break;
      }
    }
  }
  ws.sheetProtection = value.sheetProtection;
  ws.tables = value.tables.reduce((tables: { [key: string]: TableData }, table: TableModel) => {
    const t = createTable(ws, table);
    tableSetModel(t, table);
    tables[table.name] = t;
    getSheetWorkbook(ws)._tableNames.add(table.name.toLowerCase());
    return tables;
  }, {});
  ws.pivotTables = value.pivotTables;
  for (const pivotTable of ws.pivotTables ?? []) {
    pivotTable.worksheetName ??= getSheetName(ws);
    pivotTable.name ??= `PivotTable${pivotTable.tableNumber}`;
  }
  ws.conditionalFormattings = value.conditionalFormattings;
  ws.ignoredErrors = value.ignoredErrors ?? [];
  ws.threadedComments = value.threadedComments ?? [];
  // **`getModel` has always written this out and nothing read it back**, which made `getModel`/`setModel`
  // asymmetric for sparklines and is not a defect of any one container: every path that carries a worksheet
  // through its model dropped them. `Workbook.read` is one such path — it builds sheets internally and
  // transfers them to the caller's workbook as a model — so an XLSB *and* an XLSX read both produced a sheet
  // whose sparklines had been read correctly and then discarded one step later. Copied rather than aliased so
  // a caller mutating the model it passed in cannot reach inside the sheet.
  ws._sparklineGroups = value.sparklineGroups ? [...value.sparklineGroups] : [];
  // Rebuild form controls from the serialised model so importSheet() and any
  // other model round-trip preserves checkbox state, position, and links.
  ws.formControls = (value.formControls ?? []).map(fcModel => formCheckboxFromModel(ws, fcModel));
  // Preserve loaded drawing data (charts, etc.)
  ws._drawing = value.drawing;
  // Restore chart handles from the model (explicit `charts` array) or from
  // drawing anchors. `createChart` is imported statically, so this works
  // unconditionally — pure load-save pass-through still flows through the same
  // path, and a consumer that never references any chart API gets the chart
  // implementation tree-shaken out regardless.
  if (value.charts && value.charts.length > 0) {
    ws._charts = value.charts.map((c: ChartAnchorModel) =>
      createChart(ws, { chartNumber: c.chartNumber, chartExNumber: c.chartExNumber }, c.range)
    );
  } else if ((value.drawing as { anchors?: unknown } | undefined)?.anchors) {
    // Extract chart anchors from drawing (loaded from XLSX)
    type DrawingChartAnchor = {
      chartNumber?: number;
      chartExNumber?: number;
      range?: Parameters<typeof createChart>[2];
    };
    const anchors = (value.drawing as { anchors: DrawingChartAnchor[] }).anchors;
    ws._charts = anchors
      .filter(a => a.chartNumber || a.chartExNumber)
      .map(a =>
        createChart(
          ws,
          { chartNumber: a.chartNumber ?? 0, chartExNumber: a.chartExNumber ?? 0 },
          a.range!
        )
      );
  } else {
    ws._charts = [];
  }
}

function _parseRows(ws: WorksheetData, model: WorksheetModel): void {
  ws._rows = [];
  if (model.rows) {
    model.rows.forEach(rowModel => {
      const row = rowCreate(ws, rowModel.number);
      ws._rows[row.number - 1] = row;
      rowSetModel(row, rowModel);
    });
  }
}

function _parseMergeCells(ws: WorksheetData, model: WorksheetModel): void {
  if (model.mergeCells) {
    model.mergeCells.forEach((merge: string) => {
      // Do not merge styles when importing an Excel file
      // since each cell may have different styles intentionally.
      mergeCellsWithoutStyle(ws, merge);
    });
  }
}
