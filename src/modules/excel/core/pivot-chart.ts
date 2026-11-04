import type {
  AddChartOptions,
  AddComboChartOptions,
  PivotChartOptions as PivotChartMetadataOptions,
  PivotChartSource
} from "@excel/chart/model/types";
import type { PivotTable } from "@excel/core/pivot-table";
import type { WorkbookData } from "@excel/core/workbook-core";
import { quoteSheetName } from "@excel/utils/address";
import { excelNameKey } from "@utils/name-key";
import { xmlDecode } from "@utils/xml-encode";

const DEFAULT_PIVOT_CHART_PIVOT_AREA_XML =
  '<pivotArea type="data" outline="0" fieldPosition="0"><references count="1"><reference field="4294967294" count="1" selected="0"><x v="0"/></reference></references></pivotArea>';

type PivotChartCreationOptions = AddChartOptions | AddComboChartOptions;

export function withPivotChartSource<T extends PivotChartCreationOptions>(
  pivotTable: PivotTable,
  options: T,
  fmtId?: number,
  pivotChartOptions?: PivotChartMetadataOptions
): T & { pivotSource: PivotChartSource } {
  // Auto-assign the next available fmtId when the caller didn't pass
  // one. Before this fix, a second `addPivotChart(pivotTable, …)`
  // against the same pivot would reuse `fmtId=0`, collide with the
  // first chart's `chartFormat` entry (the `exists` guard in
  // `ensurePivotChartFormat` would no-op), and produce two charts
  // pointing at a single pivotArea declaration — Excel renders only
  // one of them correctly.
  const existingFormats = pivotTable.chartFormats ?? [];
  const allocatedFmtId =
    fmtId ??
    (existingFormats.length === 0 ? 0 : Math.max(...existingFormats.map(f => f.format ?? 0)) + 1);
  const resolvedPivotOpts = pivotChartOptions ?? options.pivotChartOptions;
  ensurePivotChartFormat(pivotTable, allocatedFmtId);
  applyPivotChartOptions(pivotTable, resolvedPivotOpts);
  return {
    ...options,
    pivotSource: {
      name: buildPivotChartSourceName(pivotTable),
      fmtId: allocatedFmtId,
      options: resolvedPivotOpts
    }
  };
}

function applyPivotChartOptions(
  pivotTable: PivotTable,
  options: PivotChartMetadataOptions | undefined
): void {
  if (!options) {
    return;
  }
  pivotTable.pivotChartOptions = options;
  if (options.refreshOnOpen && pivotTable.cacheDefinition) {
    pivotTable.cacheDefinition.refreshOnLoad = "1";
  }
}

function ensurePivotChartFormat(pivotTable: PivotTable, fmtId: number): void {
  pivotTable.chartFormat ??= 1;
  pivotTable.chartFormats ??= [];
  const exists = pivotTable.chartFormats.some(
    format => format.chart === 0 && format.format === fmtId
  );
  if (!exists) {
    pivotTable.chartFormats.push({
      chart: 0,
      format: fmtId,
      series: true,
      pivotAreaXml: DEFAULT_PIVOT_CHART_PIVOT_AREA_XML
    });
  }
}

function buildPivotChartSourceName(pivotTable: PivotTable): string {
  const tableName = pivotTable.name ?? `PivotTable${pivotTable.tableNumber}`;
  return pivotTable.worksheetName
    ? `${quoteSheetName(pivotTable.worksheetName)}!${tableName}`
    : tableName;
}

const PIVOT_SOURCE_NAME =
  /<(?:[\w.-]+:)?pivotSource\b[^>]*>[\s\S]*?<(?:[\w.-]+:)?name\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?name>/g;
const BOOK_PREFIX = /^\[[^\]]*\]/;
const CHART_PART = /(?:^|\/)charts\/chart\d+\.xml$/i;

/**
 * Whether any chart in the workbook names `pivotTable` in its `pivotSource` — a chart the library
 * parsed, the original bytes it keeps for one, or a chart part it preserved verbatim without
 * parsing it at all.
 *
 * Every one of those holds the reference as XML: the chart model keeps `pivotSource` as the raw
 * `<c:pivotSource>` element (there is no structured form of it to compare), so the reference is
 * read out of the element's `<c:name>` in every case.
 *
 * A loaded chart may prefix the reference with the workbook it was authored in
 * (`[Book1.xlsx]Sheet1!PivotTable1`), and may or may not quote the sheet name, so both spellings
 * are compared unquoted after that prefix is dropped — case-insensitively, as Excel resolves sheet
 * and pivot-table names.
 */
export function isPivotChartSource(wb: WorkbookData, pivotTable: PivotTable): boolean {
  const tableName = pivotTable.name ?? `PivotTable${pivotTable.tableNumber}`;
  const target = (
    pivotTable.worksheetName ? `${pivotTable.worksheetName}!${tableName}` : tableName
  ).toLowerCase();
  const names = (xml: string): string[] =>
    [...xml.matchAll(PIVOT_SOURCE_NAME)].map(match =>
      // `[Book1.xlsx]'My Sheet'!T` and `'[Book1.xlsx]My Sheet'!T` are both written.
      unquoteReference(xmlDecode(match[1]).trim().replace(BOOK_PREFIX, ""))
        .replace(BOOK_PREFIX, "")
        .toLowerCase()
    );
  const decoder = new TextDecoder();
  const sources: string[] = [];
  for (const entry of Object.values(wb._chartEntries)) {
    if (entry.model.pivotSource) {
      sources.push(entry.model.pivotSource);
    }
    if (entry.rawData) {
      sources.push(decoder.decode(entry.rawData));
    }
  }
  for (const part of wb._opaqueParts) {
    if (CHART_PART.test(part.path)) {
      sources.push(decoder.decode(part.data));
    }
  }
  return sources.some(xml => names(xml).includes(target));
}

/** A `'Sheet ''1'''!Table` reference with its sheet unquoted; an unquoted one is returned as it is. */
function unquoteReference(reference: string): string {
  const match = /^'((?:[^']|'')*)'!(.*)$/.exec(reference);
  return match ? `${match[1].replace(/''/g, "'")}!${match[2]}` : reference;
}

const FILTER_PIVOT_TABLE = /<(?:[\w.-]+:)?pivotTable\b([^>]*)\/?>/g;
const ATTRIBUTE = (name: string) => new RegExp(`\\b${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`);
const TAB_ID = ATTRIBUTE("tabId");
const NAME = ATTRIBUTE("name");

/**
 * Whether a slicer or a timeline in the workbook filters `pivotTable` — its cache part lists the
 * pivot table as `<pivotTable tabId="…" name="…"/>`, by the id of the sheet it is on (`sheetId`)
 * and its name. Those parts are preserved from a loaded file as bytes; removing the pivot table
 * would leave them pointing at nothing, which Excel repairs by deleting the slicer.
 */
export function isPivotFilterTarget(
  wb: WorkbookData,
  pivotTable: PivotTable,
  sheetId: number
): boolean {
  const tableName = excelNameKey(pivotTable.name ?? `PivotTable${pivotTable.tableNumber}`);
  const decoder = new TextDecoder();
  for (const parts of [wb._slicerCacheParts, wb._timelineCacheParts]) {
    for (const bytes of Object.values(parts ?? {})) {
      for (const [, attributes] of decoder.decode(bytes).matchAll(FILTER_PIVOT_TABLE)) {
        const tabId = TAB_ID.exec(attributes);
        const name = NAME.exec(attributes);
        if (
          tabId &&
          name &&
          Number(tabId[1] ?? tabId[2]) === sheetId &&
          excelNameKey(xmlDecode(name[1] ?? name[2])) === tableName
        ) {
          return true;
        }
      }
    }
  }
  return false;
}
