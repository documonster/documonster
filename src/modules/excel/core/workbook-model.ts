/**
 * The workbook serialisation model, and the one function that builds it.
 *
 * Split out of `workbook.browser.ts` because **both** writers need it and only one of them can
 * import that module. The buffered path reaches it through `core/workbook-format.ts` and the
 * streaming path from `stream/workbook-writer.browser.ts`, which `workbook.browser.ts` itself
 * imports — so a static import there would close a cycle. Both used to work around that with
 * `await import("@excel/core/workbook.browser")`, and that is worse than a cycle: the module is
 * statically imported by four other files, so the dynamic import could never move it into a chunk
 * of its own (rolldown says `INEFFECTIVE_DYNAMIC_IMPORT`) and instead made it a chunk *entry*,
 * whose every export counts as used. That defeated tree-shaking well outside this module — a
 * `Cell`-only consumer was paying for the formula tokenizer, because `setWorkbookModel` could no
 * longer be dropped and it reaches the defined-name syntax probe. Sitting below both writers, the
 * model needs no dynamic import from either.
 *
 * The model types live here too rather than beside their producer, so that `workbook-core.ts` can
 * name them without importing upward into `workbook.browser` — which its own header says nothing
 * below it does.
 */

import type { ChartExEntry } from "@excel/chart/model/chart-ex-types";
import type { ChartEntry } from "@excel/chart/model/types";
import type { DefinedNameModel } from "@excel/core/defined-names";
import { definedNamesModel } from "@excel/core/defined-names";
import type { OpaqueDrop, OpaquePart } from "@excel/core/opaque-part";
import type { PivotTable } from "@excel/core/pivot-table";
import { sheetsInTabOrder } from "@excel/core/sheet-order";
import type { NamedStyleEntry, WorkbookData } from "@excel/core/workbook-core";
import { getWorksheets } from "@excel/core/workbook-core";
import type { WorksheetModel } from "@excel/core/worksheet";
import { getSheetModel } from "@excel/core/worksheet";
import type {
  CalculationProperties,
  Font,
  ThreadedCommentPerson,
  WorkbookProperties,
  WorkbookView
} from "@excel/types";
import type { RelationshipModel } from "@excel/xlsx/xform/core/relationship-xform";
import type { ChartsheetModel } from "@excel/xlsx/xform/sheet/chartsheet-xform";

/** Workbook-level media entry (looser than the public `ImageData` input shape). */
export interface WorkbookMedia {
  type: string;
  extension: string;
  filename?: string;
  buffer?: Uint8Array;
  base64?: string;
  name?: string;
  /** External link target — when set, the image is referenced, not embedded. */
  link?: string;
  /**
   * Media index of the SVG companion for this raster image. When set, the
   * picture is written as a raster `a:blip` plus an `asvg:svgBlip` extension
   * referencing the SVG media at this index. Internal bookkeeping only.
   */
  svgMediaId?: number;
}

/** Internal model type for serialization */
export interface WorkbookModel {
  creator?: string;
  lastModifiedBy?: string;
  lastPrinted?: Date;
  created: Date;
  modified: Date;
  properties: Partial<WorkbookProperties>;
  protection?: WorkbookProtectionModel;
  worksheets: WorksheetModel[];
  sheets?: WorksheetModel[];
  definedNames: DefinedNameModel[];
  /**
   * Live `DefinedNames` instance — used by the write-time chartEx
   * transform (`prepareChartExSidecars`) which registers hidden
   * `_xlchart.vN.M` defined names on the fly and needs an object
   * it can mutate in place. The serialised `definedNames` array
   * above is re-materialised from this instance after the
   * transform runs. Optional because the model is also used for
   * input paths that don't carry the live instance.
   */
  definedNamesInstance?: unknown;
  views: WorkbookView[];
  company: string;
  manager: string;
  title: string;
  subject: string;
  keywords: string;
  category: string;
  description: string;
  language?: string;
  revision?: number;
  contentStatus?: string;
  /**
   * Override ContentType for `/xl/workbook.xml`, captured from the source
   * file's `[Content_Types].xml` on read (templates and macro-enabled
   * workbooks declare a different value than a plain .xlsx). Undefined for a
   * freshly created workbook, in which case the writer emits the plain type.
   */
  workbookContentType?: string;
  themes?: unknown;
  media: WorkbookMedia[];
  pivotTables: PivotTable[];
  /** Loaded pivot tables from file - used during reconciliation */
  loadedPivotTables?: PivotTable[];
  calcProperties: Partial<CalculationProperties>;
  /** Default font preserved from the original file for round-trip fidelity */
  defaultFont?: Partial<Font>;
  /** Workbook-level named cell styles (OOXML cellStyles), in definition order. */
  cellStyles?: NamedStyleEntry[];
  /** Chart entries indexed by 1-based chart number */
  chartEntries?: Record<number, ChartEntry>;
  /** Chart rels indexed by chart number — preserved for round-trip */
  chartRels?: Record<number, RelationshipModel[]>;
  /** Chart style XML raw bytes indexed by style number — preserved for round-trip */
  chartStyles?: Record<number, Uint8Array>;
  /** Chart colors XML raw bytes indexed by colors number — preserved for round-trip */
  chartColors?: Record<number, Uint8Array>;
  chartExStyles?: Record<number, Uint8Array>;
  chartExColors?: Record<number, Uint8Array>;
  /** ChartEx raw bytes (Office 2016+ extended charts) indexed by chartEx number */
  chartExEntries?: Record<number, Uint8Array>;
  /** ChartEx rels indexed by chartEx number */
  chartExRels?: Record<number, RelationshipModel[]>;
  /** Structured chartEx entries (loaded or programmatically built) indexed by chartEx number */
  chartExStructuredEntries?: Record<number, ChartExEntry>;
  /** Chartsheets parsed from the XLSX file — preserved for round-trip */
  chartsheets?: ChartsheetModel[];
  /**
   * Office 365 threaded-comment person directory, hydrated from
   * `xl/persons/person.xml` on load and serialised back on save when
   * non-empty. See {@link Workbook.persons}.
   */
  persons?: ThreadedCommentPerson[];
  /**
   * Raw-passthrough slicer parts keyed by zip-relative path. Documonster
   * does not structurally model slicers yet but preserves the bytes on
   * round-trip so dashboards continue to work.
   */
  slicerParts?: Record<string, Uint8Array>;
  slicerCacheParts?: Record<string, Uint8Array>;
  timelineParts?: Record<string, Uint8Array>;
  timelineCacheParts?: Record<string, Uint8Array>;
  /**
   * Package parts this library does not model, preserved verbatim together with
   * their content type and the relationships that reach them.
   *
   * Without this the loader drained unrecognised entries and dropped the bytes,
   * so `read` followed by `write` silently deleted a VBA project, custom
   * document properties, data connections, query tables and printer settings.
   * The macro case was the sharpest: `workbookContentType` *is* round-tripped,
   * so the output kept declaring itself macro-enabled with no macros left in it.
   */
  opaqueParts?: OpaquePart[];
  xlsbPivotCaches?: readonly { readonly cacheId: number; readonly relationshipId: string }[];
  /**
   * `Default` content-type declarations the preserved parts rely on, keyed by
   * lower-case extension. Separate from {@link OpaquePart.contentType} because a
   * Default is a property of the package, not of one part.
   */
  opaqueContentTypeDefaults?: Record<string, string>;
  /**
   * Preserved parts that were deliberately not written back, and why.
   *
   * Reported rather than discarded because two of the reasons are things a caller
   * may need to act on: a digital signature is removed on any write that
   * re-serialises a modelled part, and a part becomes unreachable when the sheet
   * that referenced it is deleted. Populated on read and topped up at write time,
   * when reachability is finally known.
   */
  opaqueDrops?: OpaqueDrop[];
  /**
   * What an XLSX write had to drop because it belongs to the other container's sheet family.
   *
   * Written by `_resolveOpaqueReachability` and read back by `writeToZip` in the same call, which is what distinguishes
   * it from `opaqueDrops` — that one is a *read-time* report and a write never reaches its readers. This exists so
   * `unsupported` can govern the one loss an XLSX write has; see `refuseXlsxUnsupported`.
   */
  opaqueForeignSheetParts?: readonly string[];
  /**
   * External workbook references in declaration order. Matches the on-disk
   * `[N]Sheet!Ref` indexing (1-based). Empty or undefined when the workbook
   * has no external references.
   */
  externalLinks?: ExternalLinkModel[];
}

/** Internal model for workbook-level protection (serialized to <workbookProtection>) */
export interface WorkbookProtectionModel {
  lockStructure?: boolean;
  lockWindows?: boolean;
  lockRevision?: boolean;
  workbookPassword?: string;
  revisionsPassword?: string;
  algorithmName?: string;
  hashValue?: string;
  saltValue?: string;
  spinCount?: number;
}

// =============================================================================
// External Workbook Link Types
// =============================================================================

/**
 * Cached values for a single sheet of an external workbook. Keys are the
 * A1-notation cell addresses *in uppercase* (e.g. `"A1"`, `"B12"`). Values
 * are the cached primitives Excel displays when the external file is not
 * currently available — must be JSON primitives: string, number, boolean, or
 * null for an explicitly blank cell.
 */
export type ExternalLinkCachedSheet = Record<string, string | number | boolean | null>;

/**
 * A single external workbook reference. Each entry corresponds to one
 * `xl/externalLinks/externalLink{N}.xml` part in the output file, and to
 * one `<externalReference r:id="...">` entry in `xl/workbook.xml`.
 *
 * The on-disk formula syntax for referring to this workbook is `[N]Sheet!A1`
 * where `N` is the 1-based `index` below.
 */
export interface ExternalLinkModel {
  /**
   * The 1-based index used in `[N]Sheet!A1` formulas. This is the position
   * in the workbook's `<externalReferences>` list (in declaration order).
   * Assigned automatically on read/write; treat as read-only when produced
   * by the library.
   */
  index: number;
  /**
   * The rel Target that will be written into
   * `xl/externalLinks/_rels/externalLink{N}.xml.rels`. For relative paths
   * (which is what users almost always want), pass the bare filename or a
   * path relative to the current workbook: `"测试.xlsx"`, `"data/ref.xlsx"`.
   * Office resolves bare relative paths from the current workbook's
   * directory — *that* is the fix for the "Office goes to the Documents
   * folder" problem with external links.
   *
   * Absolute `file:///` or `http(s)://` URIs are accepted and written
   * through unchanged.
   */
  target: string;
  /**
   * Almost always `"External"`. `"Internal"` is for embedded workbooks
   * (rare) and is preserved on round-trip when present in the source file.
   */
  targetMode: "External" | "Internal";
  /**
   * The relationship id inside `xl/_rels/workbook.xml.rels` pointing to this
   * external link's XML part. Populated automatically on read and
   * re-assigned on write. Callers should leave this undefined.
   */
  rId?: string;
  /**
   * The sheet names exposed by the external workbook, in declaration order.
   * Excel writes one `<sheetName val="..."/>` per entry under
   * `<sheetNames>` inside the externalLink part.
   *
   * At minimum you must declare every sheet that appears in a formula
   * targeting this external workbook, otherwise Excel will fail to link
   * the cached values and show `#REF!`.
   */
  sheetNames: string[];
  /**
   * Cached primitive values per sheet. Key is the *sheet name* (matching an
   * entry in `sheetNames`), value is a map from A1 address to primitive.
   *
   * Cached values are what Excel displays when the referenced external file
   * is not available (e.g. freshly-downloaded workbook on another machine).
   * Writing them turns your file from "opens with errors" into "opens,
   * shows values, offers to update links".
   */
  cachedValues?: Record<string, ExternalLinkCachedSheet>;
}

export function getWorkbookModel(wb: WorkbookData): WorkbookModel {
  // Built once and used for both `worksheets` and `sheets`, which used to call `getSheetModel` on every sheet twice.
  const worksheetModels = getWorksheets(wb).map(worksheet => getSheetModel(worksheet));
  return {
    creator: wb.creator || "Unknown",
    lastModifiedBy: wb.lastModifiedBy || "Unknown",
    lastPrinted: wb.lastPrinted,
    created: wb.created,
    modified: wb.modified,
    properties: wb.properties,
    protection: wb.protection,
    worksheets: worksheetModels,
    // **The full tab bar, chartsheets included — which is what every other producer of this field means by it.**
    //
    // This was `getWorksheets(wb).map(getSheetModel)`: a second copy of `worksheets`, with the chartsheets missing and
    // `getSheetModel` run twice per sheet. Both of the other places that set `sheets` include chartsheets — the XLSX
    // writer builds it with `sheetsInTabOrder`, and the XLSX reader assigns the parsed `<sheets>` list — so one field
    // carried two meanings, and `setWorkbookModel` reads it under the other one.
    //
    // What that cost: `setModel(getModel(wb))` moved a chartsheet to the end of the tab bar. Measured on a workbook
    // ordered worksheet `A`, chartsheet `C`, worksheet `B` — written directly it is `A, C, B` in both containers, and
    // after a model round trip `A, B, C`. It also gave `B` and `C` the same `orderNo`, so what came out depended on
    // sort stability rather than on the author's layout.
    // `?? []` because this function is also called on the *streaming* writer, through
    // `getWorkbookModel(this as never)` — a structurally similar object that carries no `_chartsheets` at all. Reading
    // the field without a fallback threw inside `sheetsInTabOrder` and took out every streamed XLSB commit.
    sheets: sheetsInTabOrder(worksheetModels, (wb._chartsheets ?? []) as never) as never,
    definedNames: definedNamesModel(wb._definedNames),
    // Live `DefinedNames` instance — required by the write-time
    // chartEx transform `prepareChartExSidecars`, which registers
    // hidden `_xlchart.vN.M` names on the fly and needs an object
    // that can mutate in place. The serialised `definedNames`
    // array above is re-materialised after the transform runs.
    definedNamesInstance: wb._definedNames,
    views: wb.views,
    company: wb.company,
    manager: wb.manager,
    title: wb.title,
    subject: wb.subject,
    keywords: wb.keywords,
    category: wb.category,
    description: wb.description,
    language: wb.language,
    revision: wb.revision,
    contentStatus: wb.contentStatus,
    workbookContentType: wb.workbookContentType,
    themes: wb._themes,
    media: wb.media,
    pivotTables: wb.pivotTables,
    calcProperties: wb.calcProperties,
    defaultFont: wb._defaultFont,
    cellStyles: wb._cellStyles ? [...wb._cellStyles.values()] : undefined,
    externalLinks: wb.externalLinks,
    chartEntries: wb._chartEntries,
    chartRels: wb._chartRels,
    chartStyles: wb._chartStyles,
    chartColors: wb._chartColors,
    chartExStyles: wb._chartExStyles,
    chartExColors: wb._chartExColors,
    chartExEntries: wb._chartExEntries,
    chartExRels: wb._chartExRels,
    chartExStructuredEntries: wb._chartExStructuredEntries,
    chartsheets: wb._chartsheets,
    persons: wb._persons,
    slicerParts: wb._slicerParts,
    slicerCacheParts: wb._slicerCacheParts,
    timelineParts: wb._timelineParts,
    timelineCacheParts: wb._timelineCacheParts,
    opaqueParts: wb._opaqueParts,
    xlsbPivotCaches: wb._xlsbPivotCaches,
    opaqueContentTypeDefaults: wb._opaqueContentTypeDefaults,
    opaqueDrops: wb._opaqueDrops
  };
}
