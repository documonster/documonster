import type { ChartExEntry } from "@excel/chart/model/chart-ex-types";
/**
 * workbook-core — the low-level Workbook container layer.
 *
 * Holds the plain-data `WorkbookData` record interface and the low-level
 * registry accessors (defined-names, chart/chartEx entries, images, sheet-name
 * validation, sheet removal) that the worksheet module reaches up into.
 *
 * Exists to break the worksheet ↔ workbook import cycle: worksheet needs these
 * workbook-level registries, and the heavy `workbook.browser` module needs to
 * create worksheets. Keeping the registries here — above worksheet, below the
 * heavy workbook module — yields a clean one-directional graph:
 *
 *     workbook.browser.ts  (addWorksheet / importWorksheet / chartsheet / IO)
 *        ↓
 *     workbook-model.ts  (WorkbookModel + getWorkbookModel — also reached by both writers)
 *        ↓
 *     workbook-core.ts  (WorkbookData + registry accessors)
 *        ↓
 *     worksheet.ts / worksheet-core.ts / cell / row / column
 *
 * No file below this one imports the heavy `workbook.browser`. This one used to name three of its
 * types, which contradicted the sentence above; they live in `workbook-model.ts` now, and the only
 * edge back up is `workbook-model.ts`'s type-level reach for `WorkbookData` — erased at compile
 * time, so the emitted graph stays one-directional.
 */
import type { ChartEntry } from "@excel/chart/model/types";
import type { BuiltInCellStyle } from "@excel/core/builtin-cell-styles";
import { BUILTIN_CELL_STYLES } from "@excel/core/builtin-cell-styles";
import type { DefinedNamesData } from "@excel/core/defined-names";
import type { OpaqueDrop, OpaquePart } from "@excel/core/opaque-part";
import type { PivotTable } from "@excel/core/pivot-table";
import type {
  WorkbookMedia,
  WorkbookProtectionModel,
  ExternalLinkModel
} from "@excel/core/workbook-model";
import type { Worksheet } from "@excel/core/worksheet";
import { describeHandle, getSheetName } from "@excel/core/worksheet-core";
import { ExcelArgumentError, ExcelImageError, ExcelWorksheetNameError } from "@excel/errors";
import type {
  Font,
  ImageData,
  NamedStyle,
  Style,
  WorkbookProperties,
  CalculationProperties,
  WorkbookView,
  ThreadedCommentPerson
} from "@excel/types";
import { argumentChecks } from "@excel/utils/arguments";
import { copyPlain } from "@excel/utils/copy-plain";
import { checkStyle } from "@excel/utils/style-checks";
import { RelType } from "@excel/xlsx/rel-type";
import type { RelationshipModel } from "@excel/xlsx/xform/core/relationship-xform";
import type { ChartsheetModel } from "@excel/xlsx/xform/sheet/chartsheet-xform";
import { isBase64 } from "@utils/base64";
import { excelNameKey, sheetNameKey } from "@utils/name-key";
/**
 * A user function as the workbook *stores* it.
 *
 * Deliberately not the engine's `FormulaFunction`, whose `invoke` must return a `RuntimeValue`:
 * building one means calling `rvNumber` / `rvString` / `BLANK`, and that made
 * `core/workbook.browser.ts` — the module every `Workbook` member goes through — import the
 * formula runtime eagerly, so a consumer who merely created a workbook carried it. The conversion
 * belongs to the engine boundary and lives in `core/formula-adapter.ts`, which is opt-in through
 * `documonster/excel/formula`; the registry keeps the caller's own function until then.
 */
export interface WorkbookFunctionDescriptor {
  readonly minArity: number;
  readonly maxArity: number;
  readonly volatile?: boolean;
  /** {@link RegisterFunctionOptions.returnsArray}. */
  readonly returnsArray?: boolean;
  /** The caller's function, unconverted — see the note above. */
  readonly invoke: (args: never[]) => unknown;
}

/**
 * A named cell style stored on the workbook: a {@link NamedStyle} plus its name.
 */
export interface NamedStyleEntry extends NamedStyle {
  name: string;
}

export interface WorkbookData {
  category: string;
  company: string;
  created: Date;
  description: string;
  keywords: string;
  manager: string;
  modified: Date;
  subject: string;
  title: string;
  creator?: string;
  lastModifiedBy?: string;
  lastPrinted?: Date;
  language?: string;
  revision?: number;
  contentStatus?: string;
  /**
   * Source file path remembered by Node `readFile`; absent for buffers/browser inputs.
   * @deprecated Internal state, not part of the API. Removed from the published type in 1.0.
   */
  sourceFilePath?: string;
  /**
   * Override ContentType for `/xl/workbook.xml` captured from the source file's
   * `[Content_Types].xml` (templates / macro-enabled workbooks differ from a
   * plain .xlsx). Undefined for a freshly created workbook.
   * @deprecated Internal state, not part of the API. Removed from the published type in 1.0.
   */
  workbookContentType?: string;
  properties: Partial<WorkbookProperties>;
  calcProperties: Partial<CalculationProperties>;
  views: WorkbookView[];
  media: WorkbookMedia[];
  /** @deprecated Use `Pivot.list`. Removed from the published type in 1.0. */
  pivotTables: PivotTable[];
  protection?: WorkbookProtectionModel;
  externalLinks: ExternalLinkModel[];
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  _worksheets: Worksheet[];
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  _definedNames: DefinedNamesData;
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  _themes?: unknown;
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  _defaultFont?: Partial<Font>;
  /**
   * Workbook-level named cell styles (OOXML `cellStyle`), keyed by `excelNameKey(name)`: Excel compares cell style
   * names without regard to case, so `Brand Title` and `brand title` are one style. Each entry keeps its spelling in
   * `name`. Populated
   * by {@link addCellStyle} and on read from an existing workbook's
   * `cellStyles`/`cellStyleXfs`.
   * @deprecated Internal state, not part of the API. Removed from the published type in 1.0.
   */
  _cellStyles?: Map<string, NamedStyleEntry>;
  /**
   * The workbook-level differential-format table (`<dxfs>`) as it was read, in index order.
   *
   * Carried because a `dxfId` is an index and not every holder of one is modelled — a pivot table's
   * `<formats>` and a `<colorFilter>` are preserved as XML with the number they were read with. The writer
   * starts its table from this one so those numbers stay valid; modelled rules and table columns hold these
   * same objects and are resolved back to the same index.
   * @deprecated Internal state, not part of the API. Removed from the published type in 1.0.
   */
  _dxfs?: (Partial<Style> | undefined)[];
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  _writerExternalLinkCache: Map<string, ExternalLinkModel>;
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  _tableNames: Set<string>;
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  _chartEntries: Record<number, ChartEntry>;
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  _chartRels: Record<number, RelationshipModel[]>;
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  _chartStyles: Record<number, Uint8Array>;
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  _chartColors: Record<number, Uint8Array>;
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  _chartExStyles: Record<number, Uint8Array>;
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  _chartExColors: Record<number, Uint8Array>;
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  _chartExEntries: Record<number, Uint8Array>;
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  _chartExRels: Record<number, RelationshipModel[]>;
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  _chartExStructuredEntries: Record<number, ChartExEntry>;
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  _chartsheets: ChartsheetModel[];
  /**
   * The highest id of a sheet removed in this session — worksheets and chartsheets share the pool — so a removed
   * sheet's id is never handed out again, as Excel does it (see `getNextId`).
   * @internal
   */
  _lastSheetId?: number;
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  _persons: ThreadedCommentPerson[];
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  _slicerParts: Record<string, Uint8Array>;
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  _slicerCacheParts: Record<string, Uint8Array>;
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  _timelineParts: Record<string, Uint8Array>;
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  _timelineCacheParts: Record<string, Uint8Array>;
  /**
   * Preserved unmodelled package parts — see `WorkbookModel.opaqueParts`.
   * @deprecated Internal state, not part of the API. Removed from the published type in 1.0.
   */
  _opaqueParts: OpaquePart[];
  /**
   * Media ids an API call has stopped using in this session — see {@link releaseMedia}.
   * @deprecated Internal state, not part of the API. Removed from the published type in 1.0.
   */
  _releasedMedia?: Set<number>;
  /**
   * `BrtBeginPivotCacheID` bindings an XLSB read preserved.
   *
   * XLSB-only, and here rather than in the pivot model because the parts they name are carried through as
   * opaque bytes: the reader does not build a `PivotTable` from them, but `workbook.bin` is rebuilt on write
   * and would otherwise drop the record that announces each cache.
   * @deprecated Internal state, not part of the API. Removed from the published type in 1.0.
   */
  _xlsbPivotCaches?: readonly { readonly cacheId: number; readonly relationshipId: string }[];
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  _opaqueContentTypeDefaults: Record<string, string>;
  /**
   * Preserved parts not written back — see `WorkbookModel.opaqueDrops`.
   * @deprecated Internal state, not part of the API. Removed from the published type in 1.0.
   */
  _opaqueDrops: OpaqueDrop[];
  /** @deprecated Internal state, not part of the API. Removed from the published type in 1.0. */
  userFunctions?: Map<string, WorkbookFunctionDescriptor>;
}

/**
 * Refuse something that is not a workbook handle where one is required — `42`, `null`, `{}` — with an
 * `ExcelArgumentError` naming `field`, rather than a `TypeError` from inside. One property read: a handle carries its
 * sheet list, and nothing else that reaches here does.
 */
export function checkWorkbook(wb: unknown, field = "wb"): asserts wb is WorkbookData {
  if ((wb as { _worksheets?: unknown } | null | undefined)?._worksheets === undefined) {
    throw new ExcelArgumentError(field, `${field} must be a workbook, got ${describeHandle(wb)}.`);
  }
}

/**
 * Validate a named cell style name. Shared by the standard workbook API and the
 * streaming writer so both reject the same inputs. Throws on an empty name or
 * the reserved built-in "Normal" style (builtinId 0), which is always emitted
 * implicitly as `cellStyleXfs[0]`.
 */
export function validateCellStyleName(name: string): void {
  if (typeof name !== "string" || name === "") {
    throw new ExcelWorksheetNameError("Named cell style name must be a non-empty string.");
  }
  if (isNormalCellStyleName(name)) {
    throw new ExcelWorksheetNameError('The built-in "Normal" cell style cannot be redefined.');
  }
}

/** Whether `name` is the built-in "Normal" style, in any case — Excel compares cell style names that way. */
export function isNormalCellStyleName(name: string): boolean {
  return excelNameKey(name) === NORMAL_STYLE_KEY;
}

const NORMAL_STYLE_KEY = excelNameKey("Normal");

/** The stored named cell style `name` refers to, compared as Excel compares them (without regard to case). */
export function findCellStyleEntry(wb: WorkbookData, name: string): NamedStyleEntry | undefined {
  return typeof name === "string" ? wb._cellStyles?.get(excelNameKey(name)) : undefined;
}

/**
 * Named cell styles keyed as Excel compares their names, without regard to case. Of two that differ only in case the
 * first is kept and the later dropped, which is what Excel does with such a file.
 */
export function namedStyleMap(styles: Iterable<NamedStyleEntry>): Map<string, NamedStyleEntry> {
  const map = new Map<string, NamedStyleEntry>();
  for (const style of styles) {
    const key = excelNameKey(style.name);
    if (!map.has(key)) {
      map.set(key, style);
    }
  }
  return map;
}

/**
 * Add (or replace) a workbook-level named cell style such as "Heading 1", and return the stored entry.
 * Cells reference it via `Style.styleName` (e.g. `Cell.applyStyle`).
 *
 * Names compare as Excel compares them, without regard to case: adding `brand title` when `Brand Title` exists
 * replaces that style's definition, as re-adding `Brand Title` would, and the stored name keeps its first spelling
 * (`Brand Title`) — the one Excel keeps when it merges two such styles.
 *
 * The reserved built-in "Normal" style (builtinId 0), in any case, cannot be redefined; it is
 * always emitted implicitly as `cellStyleXfs[0]`.
 */
export function addCellStyle(wb: WorkbookData, name: string, style: NamedStyle): NamedStyleEntry {
  checkWorkbook(wb);
  validateCellStyleName(name);
  argumentChecks("Workbook.addCellStyle").required("style", style);
  argumentChecks("Workbook.addCellStyle").record("style", style);
  checkStyle("Workbook.addCellStyle", style);
  if (!wb._cellStyles) {
    wb._cellStyles = new Map();
  }
  const key = excelNameKey(name);
  const entry: NamedStyleEntry = {
    ...copyPlain(style),
    name: wb._cellStyles.get(key)?.name ?? name
  };
  wb._cellStyles.set(key, entry);
  return copyPlain(entry);
}

/**
 * Get a workbook-level named cell style by name (in any case), or `undefined` if absent. It is a copy: editing it changes nothing —
 * change the style with `addCellStyle`, which checks what it is given.
 */
export function getCellStyle(wb: WorkbookData, name: string): NamedStyleEntry | undefined {
  checkWorkbook(wb);
  const entry = findCellStyleEntry(wb, name);
  return entry && copyPlain(entry);
}

/** List all workbook-level named cell styles in definition order, as copies (see {@link getCellStyle}). */
export function listCellStyles(wb: WorkbookData): NamedStyleEntry[] {
  checkWorkbook(wb);
  return wb._cellStyles ? [...wb._cellStyles.values()].map(copyPlain) : [];
}

/** Remove a named cell style (named in any case). Returns `true` if it existed. */
export function removeCellStyle(wb: WorkbookData, name: string): boolean {
  checkWorkbook(wb);
  return wb._cellStyles && typeof name === "string"
    ? wb._cellStyles.delete(excelNameKey(name))
    : false;
}

/**
 * Define one of the built-in Excel cell-style presets (e.g. "Heading1") on the
 * workbook and return its OOXML style name (e.g. "Heading 1") for use with
 * `Cell.applyStyle`. Built-in presets carry the correct `builtinId`, which
 * accessibility software relies on to recognise the style's semantic role. A custom style of the same name in
 * another case is that style, and is replaced; the name returned is the stored spelling.
 */
export function useBuiltInCellStyle(wb: WorkbookData, id: BuiltInCellStyle): string {
  checkWorkbook(wb);
  if (typeof id !== "string" || !Object.hasOwn(BUILTIN_CELL_STYLES, id)) {
    argumentChecks("Workbook.useBuiltInCellStyle").fail(
      "id",
      `one of ${Object.keys(BUILTIN_CELL_STYLES).join(", ")}`,
      id
    );
  }
  const { name, ...style } = BUILTIN_CELL_STYLES[id];
  return addCellStyle(wb, name, style).name;
}

export function removeWorksheetEx(wb: WorkbookData, worksheet: Worksheet): void {
  // Everything the sheet drew stops being a use of its media: pictures, background, watermark and
  // header/footer images, and the picture fills of its charts. The charts go with the sheet — left
  // registered, each was still written as an `xl/charts/chartN.xml` nothing pointed at.
  for (const medium of worksheet._media ?? []) {
    releaseMedia(wb, medium.imageId);
  }
  for (const chart of worksheet._charts ?? []) {
    if (chart.chartNumber > 0) {
      releaseChartMedia(wb, chart.chartNumber);
      removeChartEntry(wb, chart.chartNumber);
    }
    if (chart.chartExNumber > 0) {
      removeChartExStructuredEntry(wb, chart.chartExNumber);
    }
  }
  // Release any workbook-wide table names this sheet held so the names can
  // be reused by future tables on other sheets without spurious "name
  // already exists" errors.
  const tables = worksheet.tables;
  if (tables) {
    for (const tableName of Object.keys(tables)) {
      wb._tableNames.delete(excelNameKey(tableName));
    }
  }
  worksheet._removed = true;
  // Its id is not handed out again — see `getNextId`.
  wb._lastSheetId = Math.max(wb._lastSheetId ?? 0, worksheet.id);
  wb._worksheets[worksheet.id] = undefined!;
}

export function getImage(wb: WorkbookData, id: number | string): WorkbookMedia | undefined {
  return wb.media[Number(id)];
}

/**
 * Record that an API call removed a use of medium `id`. The id stays valid; the writer leaves the
 * medium out of the package only if it is one of these *and* nothing written references it. A
 * medium that arrived unreferenced in a file that was read, or that was added and never placed,
 * is written as it always was.
 *
 * A picture with an SVG companion is two media — the raster fallback the blip embeds and the SVG its
 * `asvg:svgBlip` extension names — so releasing one releases the other too. Releasing only the raster
 * left the SVG written with nothing referencing it.
 */
export function releaseMedia(wb: WorkbookData, id: number | string | undefined): void {
  const index = id === undefined ? NaN : Number(id);
  if (Number.isInteger(index) && index >= 0) {
    const released = (wb._releasedMedia ??= new Set());
    if (released.has(index)) {
      return;
    }
    released.add(index);
    const svg = (wb.media[index] as { svgMediaId?: number } | undefined)?.svgMediaId;
    if (svg !== undefined) {
      releaseMedia(wb, svg);
    }
  }
}

/**
 * Release the media a chart's relationships point at — its picture fills. Call it before the chart's
 * entry is removed. A chart image is written as `xl/media/image{id + 1}`, which is how the
 * relationship names it — a chart read from a package that named its media otherwise has its targets
 * rewritten to that form on read (`reconcileChartMediaTargets`), so the name is the id.
 */
export function releaseChartMedia(wb: WorkbookData, chartNumber: number): void {
  if (!(chartNumber > 0)) {
    return;
  }
  const rels = [
    ...((wb._chartEntries[chartNumber]?.rels ?? []) as { Target?: string }[]),
    ...((wb._chartRels[chartNumber] ?? []) as { Target?: string }[])
  ];
  for (const rel of rels) {
    const match = /(?:^|\/)media\/image(\d+)\.[^/]+$/.exec(rel.Target ?? "");
    if (match) {
      releaseMedia(wb, Number(match[1]) - 1);
    }
  }
}

/**
 * The next free `chartN` number, counting **preserved parts as taken**.
 *
 * A chart read from a package this library does not fully model arrives as an *opaque part* — `xl/charts/chart1.xml`
 * kept verbatim — and never enters `_chartEntries`. Numbering from that map alone therefore handed `1` straight back:
 * adding a chart to a workbook read from XLSB overwrote the chart it came with, so two charts left one behind and the
 * surviving drawing pointed at the wrong picture. The package writer's dedupe made it silent, because the newly
 * generated part is written before the preserved one and the preserved one is the copy that gets dropped.
 *
 * Reading the paths is the honest test: it asks what is actually in the package rather than what the model happens to
 * have recorded about it. `chartEx` shares the counter deliberately — `nextChartExNumber` reads the same space — so a
 * `chartEx1.xml` also reserves 1.
 */
/**
 * Chart numbers already taken by *preserved* parts.
 *
 * `chart1.xml` and `chartEx1.xml` are read together on purpose: both counters allocate into `xl/charts/`, so a name
 * taken by either form is taken for both.
 */
function preservedChartNumbers(wb: WorkbookData): readonly number[] {
  return wb._opaqueParts
    .map(part => /\/charts\/chart(?:Ex)?(\d+)\.xml$/i.exec(part.path)?.[1])
    .filter((digits): digits is string => digits !== undefined)
    .map(Number)
    .filter(number => Number.isFinite(number));
}

export function nextChartNumber(wb: WorkbookData): number {
  const existing = [...Object.keys(wb._chartEntries).map(Number), ...preservedChartNumbers(wb)];
  return existing.length > 0 ? Math.max(...existing) + 1 : 1;
}

export function addChartEntry(wb: WorkbookData, entry: ChartEntry): void {
  wb._chartEntries[entry.chartNumber] = entry;
}

export function setChartStyle(wb: WorkbookData, chartNumber: number, data: Uint8Array): void {
  wb._chartStyles[chartNumber] = data;
}

export function setChartColors(wb: WorkbookData, chartNumber: number, data: Uint8Array): void {
  wb._chartColors[chartNumber] = data;
}

export function getChartEntry(wb: WorkbookData, chartNumber: number): ChartEntry | undefined {
  return wb._chartEntries[chartNumber];
}

export function removeChartEntry(wb: WorkbookData, chartNumber: number): void {
  delete wb._chartEntries[chartNumber];
  delete wb._chartRels[chartNumber];
  delete wb._chartStyles[chartNumber];
  delete wb._chartColors[chartNumber];
}

export function nextChartExNumber(wb: WorkbookData): number {
  const rawKeys = Object.keys(wb._chartExEntries ?? {}).map(Number);
  const structKeys = Object.keys(wb._chartExStructuredEntries ?? {}).map(Number);
  // Preserved parts count as taken, for the reason given on `nextChartNumber`: a chart kept verbatim never enters the
  // entry maps, so numbering from them alone reissues a name that is already in the package.
  const combined = [...rawKeys, ...structKeys, ...preservedChartNumbers(wb)];
  return combined.length > 0 ? Math.max(...combined) + 1 : 1;
}

export function addChartExStructuredEntry(wb: WorkbookData, entry: ChartExEntry): void {
  if (!wb._chartExStructuredEntries) {
    wb._chartExStructuredEntries = {};
  }
  wb._chartExStructuredEntries[entry.chartExNumber] = entry;
}

export function getChartExStructuredEntry(
  wb: WorkbookData,
  chartExNumber: number
): ChartExEntry | undefined {
  return wb._chartExStructuredEntries?.[chartExNumber];
}

/**
 * Refuse a sheet name longer than Excel's 31 characters. The API entry points call this before
 * {@link validateSheetName}, which shortens such a name — what a read of another program's file needs, and what a
 * caller who chose the name should be told about instead: the sheet ended up named something they did not write, and
 * formulas they wrote with the full name pointed nowhere.
 */
export function refuseLongSheetName(name: unknown): void {
  if (typeof name === "string" && name.length > 31) {
    throw new ExcelWorksheetNameError(
      `Sheet name ${JSON.stringify(name)} is ${name.length} characters long; Excel allows at most 31.`
    );
  }
}

export function validateSheetName(
  wb: WorkbookData,
  name: string,
  existing?: Worksheet | { name: string }
): string {
  if (typeof name !== "string") {
    throw new ExcelWorksheetNameError("The name has to be a string.");
  }
  if (name === "") {
    throw new ExcelWorksheetNameError("The name can't be empty.");
  }
  if (name === "History") {
    throw new ExcelWorksheetNameError(
      'The name "History" is protected. Please use a different name.'
    );
  }
  // Illegal characters per Excel's own naming rules: asterisk (*),
  // question mark (?), colon (:), forward slash (/), backslash (\),
  // left bracket ([), right bracket (]). The chartsheet regex used
  // to omit `\\`; unified here so both families enforce the same
  // char set.
  if (/[*?:/\\[\]]/.test(name)) {
    throw new ExcelWorksheetNameError(
      `Sheet name ${name} cannot include any of the following characters: * ? : \\ / [ ]`
    );
  }
  if (/(^')|('$)/.test(name)) {
    throw new ExcelWorksheetNameError(
      `The first or last character of sheet name cannot be a single quotation mark: ${name}`
    );
  }
  if (name.length > 31) {
    if (process.env.NODE_ENV !== "production") {
      console.warn(`Sheet name ${name} exceeds 31 chars. This will be truncated`);
    }
    name = name.substring(0, 31);
  }
  const nameKey = sheetNameKey(name);
  const duplicateWorksheet = wb._worksheets.find(
    ws => ws && ws !== existing && sheetNameKey(ws._name) === nameKey
  );
  const duplicateChartsheet = wb._chartsheets.find(
    cs => cs && cs !== existing && sheetNameKey(cs.name) === nameKey
  );
  if (duplicateWorksheet || duplicateChartsheet) {
    throw new ExcelWorksheetNameError(`Sheet name already exists: ${name}`);
  }
  return name;
}

export function removeChartExStructuredEntry(wb: WorkbookData, chartExNumber: number): void {
  if (wb._chartExStructuredEntries) {
    delete wb._chartExStructuredEntries[chartExNumber];
  }
  delete wb._chartExEntries[chartExNumber];
  delete wb._chartExRels[chartExNumber];
  delete wb._chartExStyles[chartExNumber];
  delete wb._chartExColors[chartExNumber];
}

/** The workbook's defined names — the registry the `DefinedNames` namespace edits. */
export function getDefinedNames(wb: WorkbookData): DefinedNamesData {
  checkWorkbook(wb);
  return wb._definedNames;
}
/**
 * The extensions an added image may carry: those `ImageData` documents, and `jpg`, which people write anyway
 * and `imageContentTypeFor` types as `image/jpeg`. The extension names the media part and its content type,
 * so anything else — a MIME type such as `image/png` above all — produced a part named `image1.image/png`
 * that Excel reports as unreadable.
 */
const IMAGE_EXTENSIONS: readonly string[] = ["png", "jpeg", "gif", "jpg"];

/**
 * Throw {@link ExcelImageError} unless `extension` is one an added image may carry. An omitted extension is
 * left to the writers, which name such a part `.png`.
 */
export function assertImageExtension(extension: unknown): void {
  if (extension === undefined) {
    return;
  }
  if (typeof extension !== "string" || !IMAGE_EXTENSIONS.includes(extension.toLowerCase())) {
    throw new ExcelImageError(
      `Invalid image extension ${JSON.stringify(extension)}; expected one of ${IMAGE_EXTENSIONS.join(", ")} (a file extension, not a MIME type)`
    );
  }
}

export function addWorkbookImage(wb: WorkbookData, image: ImageData): number {
  checkImageData("Image.add", image);
  return registerWorkbookImage(wb, image);
}

/**
 * Refuse an image that cannot be written: no extension or one that is not an image's, or no source of bytes —
 * `buffer`, `base64` or `filename` — and no `link`. Each of those was accepted and failed when the workbook was written, with a
 * `TypeError` from inside the writer, or was written as a picture whose bytes were a string.
 */
export function checkImageData(where: string, image: unknown): void {
  if (image === null || typeof image !== "object") {
    throw new ExcelImageError(`${where}: the image must be an object, got ${String(image)}.`);
  }
  const record = image as Record<string, unknown>;
  // The type requires it, and the XLSX writer named the part `image1.undefined` and failed on its content type.
  if (record.extension === undefined) {
    throw new ExcelImageError(
      `${where}: the image needs an extension, one of ${IMAGE_EXTENSIONS.join(", ")}.`
    );
  }
  assertImageExtension(record.extension);
  const sources = ["buffer", "base64", "filename"].filter(key => record[key] != null);
  if (
    record.buffer != null &&
    !(
      ArrayBuffer.isView(record.buffer) &&
      (record.buffer as { BYTES_PER_ELEMENT?: number }).BYTES_PER_ELEMENT === 1
    )
  ) {
    throw new ExcelImageError(`${where}: the image's buffer must be bytes (a Uint8Array).`);
  }
  for (const key of ["base64", "filename", "link"]) {
    if (record[key] != null && typeof record[key] !== "string") {
      throw new ExcelImageError(`${where}: the image's ${key} must be a string.`);
    }
  }
  // Checked here rather than when the workbook is written, where it failed in a browser and was
  // written as whatever Node's lenient decoder made of it. The `data:` prefix is optional.
  if (typeof record.base64 === "string") {
    const payload = record.base64.slice(record.base64.indexOf(",") + 1);
    if (!isBase64(payload, "either")) {
      throw new ExcelImageError(`${where}: the image's base64 is not base64.`);
    }
  }
  if (sources.length === 0 && record.link == null) {
    throw new ExcelImageError(
      `${where}: the image has no bytes; give one of buffer, base64 or filename, or a link.`
    );
  }
}

/**
 * {@link addWorkbookImage} without the extension check, for a reader registering a medium a package
 * already holds under whatever extension it was stored with (`emf`, `svg`, …).
 */
export function registerWorkbookImage(wb: WorkbookData, image: ImageData): number {
  const { svg, ...raster } = image;
  if (
    svg &&
    raster.link &&
    raster.buffer == null &&
    raster.base64 == null &&
    raster.filename == null
  ) {
    // An SVG companion needs an embedded raster fallback; a *linked* (external)
    // raster has no package part to attach the svgBlip extension to.
    throw new ExcelImageError(
      "An SVG image requires an embedded raster fallback (buffer/base64/filename); it cannot be combined with an external link."
    );
  }
  const id = wb.media.length;
  const rasterMedia: WorkbookMedia = { ...raster, type: "image" };
  wb.media.push(rasterMedia);

  if (svg) {
    // Register the SVG as a second `type: "image"` media so it flows through
    // the existing media naming, content-types, and zip-writing paths. Link
    // it back to the raster blip so the drawing serializer can emit the
    // asvg:svgBlip extension.
    const svgId = wb.media.length;
    wb.media.push({ ...svg, type: "image", extension: "svg" });
    rasterMedia.svgMediaId = svgId;
  }

  return id;
}

/**
 * Look up a worksheet by its **sheet id** or by name, compared as Excel compares sheet names (or the first sheet when
 * `id` is omitted). The id is the stable number a sheet was given when it was added — not its position, so it does
 * not move when another sheet is removed, and a removed sheet's id is not given to a later one — whereas the
 * chartsheet lookups take a 0-based position among the chartsheets. Pure data accessor over `wb._worksheets` — kept in the core
 * layer so the chart module can resolve sheets for cache population without
 * forming a `workbook → chart → workbook` import cycle.
 */
export function getWorksheet(wb: WorkbookData, id?: number | string): Worksheet | undefined {
  checkWorkbook(wb);
  if (id === undefined) {
    return wb._worksheets.find(Boolean);
  }
  if (typeof id === "number") {
    return wb._worksheets[id];
  }
  if (typeof id === "string") {
    const idKey = sheetNameKey(id);
    return wb._worksheets.find(
      worksheet => worksheet && sheetNameKey(getSheetName(worksheet)) === idKey
    );
  }
  return undefined;
}

/** All worksheets in tab order (excludes the reserved index-0 slot). See {@link getWorksheet}. */
export function getWorksheets(wb: WorkbookData): Worksheet[] {
  checkWorkbook(wb);
  return wb._worksheets
    .slice(1)
    .sort((a, b) => a.orderNo - b.orderNo)
    .filter(Boolean);
}

/**
 * Copy a classic chart's style/colours/rels sidecars from `sourceChartNumber`
 * to `targetChartNumber` (optionally into a different `targetWorkbook`).
 *
 * Pure data operation over the workbook's sidecar records — lives in the core
 * layer so `chart`'s `copyTo`/`clone` handle ops can use it without a
 * `workbook → chart → workbook` import cycle.
 */
export function copyChartSidecars(
  wb: WorkbookData,
  sourceChartNumber: number,
  targetChartNumber: number,
  targetWorkbook: WorkbookData = wb
): void {
  const style = wb._chartStyles[sourceChartNumber];
  if (style) {
    setChartStyle(targetWorkbook, targetChartNumber, style.slice());
  }
  const colors = wb._chartColors[sourceChartNumber];
  if (colors) {
    setChartColors(targetWorkbook, targetChartNumber, colors.slice());
  }
  // Copy the full chart rels bag (`_chartRels`), not just the
  // style/colors pair. A classic chart can carry rels to embedded
  // images (pictureFill), external data links, and `<c:userShapes>`
  // drawing parts — without copying those the clone ends up with
  // dangling rIds. Deep-copy each rel so a later mutation on the
  // source doesn't leak into the clone.
  //
  // Rewrite style/colors Targets to the destination chart number —
  // verbatim copy would leave the rel pointing at the source's
  // `style{src}.xml`, while the writer emits `style{dst}.xml` and
  // produces a chart whose .rels references a non-existent file.
  //
  // For image rels on a cross-workbook copy (`targetWorkbook !==
  // wb`), re-register each referenced image in the destination
  // workbook and rewrite the Target to point at the new media
  // file. Without this, a pictureFill that round-tripped through
  // `importWorksheet` pointed at the source workbook's media array —
  // which the destination package doesn't ship, so Excel shows a
  // broken image icon.
  const srcRels = wb._chartRels[sourceChartNumber];
  if (Array.isArray(srcRels) && srcRels.length > 0) {
    const crossWorkbook = targetWorkbook !== wb;
    targetWorkbook._chartRels[targetChartNumber] = srcRels.map(rel => {
      if (typeof rel !== "object" || rel === null) {
        return rel;
      }
      const cloned = { ...rel } as { Type?: string; Target?: string; [k: string]: unknown };
      const target = typeof cloned.Target === "string" ? cloned.Target : undefined;
      if (target) {
        if (/^style\d+\.xml$/.test(target)) {
          cloned.Target = `style${targetChartNumber}.xml`;
        } else if (/^colors\d+\.xml$/.test(target)) {
          cloned.Target = `colors${targetChartNumber}.xml`;
        } else if (crossWorkbook && cloned.Type === RelType.Image) {
          const rewritten = _rewriteCrossWorkbookImageTarget(wb, target, targetWorkbook);
          if (rewritten !== undefined) {
            cloned.Target = rewritten;
          }
        }
      }
      return cloned;
    });
  }
}

function _rewriteCrossWorkbookImageTarget(
  wb: WorkbookData,
  target: string,
  targetWorkbook: WorkbookData
): string | undefined {
  const match = /\/media\/image(\d+)\.([a-zA-Z0-9]+)$/.exec(target);
  if (!match) {
    return undefined;
  }
  const sourceMediaIndex = parseInt(match[1], 10) - 1;
  if (!Number.isFinite(sourceMediaIndex) || sourceMediaIndex < 0) {
    return undefined;
  }
  const medium = getImage(wb, sourceMediaIndex) as
    | { extension?: string; buffer?: Uint8Array; base64?: string }
    | undefined;
  if (!medium) {
    return undefined;
  }
  const ext = medium.extension as "png" | "jpeg" | "gif" | undefined;
  if (ext !== "png" && ext !== "jpeg" && ext !== "gif") {
    return undefined;
  }
  const payload: { extension: "png" | "jpeg" | "gif"; buffer?: Uint8Array; base64?: string } = {
    extension: ext
  };
  // `instanceof Uint8Array` is realm-sensitive: buffers that crossed
  // a Worker / iframe / `structuredClone` boundary carry a different
  // `Uint8Array` prototype and fail the operator even though they
  // are byte-granular typed arrays. Duck-type via `ArrayBuffer.isView`
  // + `BYTES_PER_ELEMENT === 1` so cross-workbook copies from a
  // worker-loaded Workbook preserve the image bytes; otherwise the
  // copy path silently falls through to `return undefined`, dropping
  // every image from the chart. Matches `chart-images.ts`'s handling
  // of the same realm-crossing issue.
  const buf = medium.buffer as ArrayBufferView | undefined;
  if (
    buf &&
    ArrayBuffer.isView(buf) &&
    (buf as unknown as { BYTES_PER_ELEMENT?: number }).BYTES_PER_ELEMENT === 1
  ) {
    payload.buffer =
      buf instanceof Uint8Array
        ? buf.slice()
        : new Uint8Array(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
  } else if (typeof medium.base64 === "string") {
    payload.base64 = medium.base64;
  } else {
    return undefined;
  }
  const newId = addWorkbookImage(targetWorkbook, payload);
  return `../media/image${newId + 1}.${ext}`;
}

/**
 * Copy a ChartEx chart's styleEx/colorsEx/rels sidecars. See
 * {@link copyChartSidecars} for the layering rationale.
 */
export function copyChartExSidecars(
  wb: WorkbookData,
  sourceChartExNumber: number,
  targetChartExNumber: number,
  targetWorkbook: WorkbookData = wb
): void {
  const rels = wb._chartExRels[sourceChartExNumber];
  if (rels && rels.length > 0) {
    // Rewrite `Target` for rels that point at numbered sidecars
    // (styleEx / colorsEx / userShapes). Those files get different
    // on-disk numbers on the clone — copying the rel verbatim
    // leaves it pointing at the source's sidecar, so saving the
    // package produces a chartEx whose .rels references
    // `styleEx{src}.xml` while the writer emits `styleEx{dst}.xml`.
    // Strip the number from the source Target and re-stamp it with
    // the target's number.
    //
    // For image rels on a cross-workbook copy (`targetWorkbook !==
    // wb`), re-register each referenced image in the destination
    // workbook and rewrite the Target — same logic as classic chart
    // sidecars. Without this, a ChartEx with embedded images (e.g.
    // pictureFill or custom geometry) would reference media that
    // doesn't exist in the destination package.
    const crossWorkbook = targetWorkbook !== wb;
    targetWorkbook._chartExRels[targetChartExNumber] = rels.map(r => {
      if (typeof r !== "object" || r === null) {
        return r;
      }
      const cloned = { ...r };
      const target: string | undefined =
        typeof cloned.Target === "string" ? cloned.Target : undefined;
      if (target) {
        const styleExMatch = /^styleEx\d+\.xml$/.exec(target);
        if (styleExMatch) {
          cloned.Target = `styleEx${targetChartExNumber}.xml`;
        } else if (/^colorsEx\d+\.xml$/.exec(target)) {
          cloned.Target = `colorsEx${targetChartExNumber}.xml`;
        } else if (crossWorkbook && cloned.Type === RelType.Image) {
          const rewritten = _rewriteCrossWorkbookImageTarget(wb, target, targetWorkbook);
          if (rewritten !== undefined) {
            cloned.Target = rewritten;
          }
        }
      }
      return cloned;
    });
  }
  // ChartEx style / colors sidecars (matching `_chartStyles` /
  // `_chartColors` for classic charts). Previously only `_chartExRels`
  // was copied — a cloned chartEx lost its authored chartExStyle and
  // chartExColors bytes, so the saved package re-derived them from
  // defaults and the clone looked different from the source.
  const exStyle = wb._chartExStyles[sourceChartExNumber];
  if (exStyle) {
    targetWorkbook._chartExStyles[targetChartExNumber] = exStyle.slice();
  }
  const exColors = wb._chartExColors[sourceChartExNumber];
  if (exColors) {
    targetWorkbook._chartExColors[targetChartExNumber] = exColors.slice();
  }
}
