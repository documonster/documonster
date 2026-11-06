import type { BuiltInCellStyle } from "@excel/core/builtin-cell-styles";
import type { DefinedNamesData } from "@excel/core/defined-names";
import type { NamedStyleEntry, WorkbookData } from "@excel/core/workbook-core";
import {
  addCellStyle as coreAddCellStyle,
  checkWorkbook,
  getCellStyle as coreGetCellStyle,
  getDefinedNames as coreGetDefinedNames,
  getWorksheet as coreGetWorksheet,
  getWorksheets as coreListWorksheets,
  listCellStyles as coreListCellStyles,
  refuseLongSheetName,
  removeCellStyle as coreRemoveCellStyle,
  useBuiltInCellStyle as coreUseBuiltInCellStyle
} from "@excel/core/workbook-core";
import type {
  ExternalLinkCachedSheet,
  ExternalLinkModel,
  WorkbookModel
} from "@excel/core/workbook-model";
/**
 * The `Workbook` operations that take values ending up in the file, checked where they are passed. Shared by the
 * Node and browser `Workbook` surfaces.
 */
import type {
  AddExternalLinkOptions,
  ProtectWorkbookOptions,
  RegisterPersonOptions
} from "@excel/core/workbook.browser";
import {
  addChartsheet as coreAddChartsheet,
  addExternalLink as coreAddExternalLink,
  addPivotChartsheet as coreAddPivotChartsheet,
  addWorksheet as coreAddWorksheet,
  copyChartsheet as coreCopyChartsheet,
  discardUnrecognizedValues as coreDiscardUnrecognizedValues,
  eachWorksheet as coreEachWorksheet,
  getChartsheet as coreGetChartsheet,
  getChartsheets as coreListChartsheets,
  getExternalLink as coreGetExternalLink,
  getWorkbookModel as coreGetWorkbookModel,
  importWorksheet as coreImportWorksheet,
  listExternalLinks as coreListExternalLinks,
  listPersons as coreListPersons,
  listUnrecognizedValues as coreListUnrecognizedValues,
  protectWorkbook,
  registerFunction as coreRegisterFunction,
  registerPerson as coreRegisterPerson,
  removeChartsheet as coreRemoveChartsheet,
  removeWorksheet as coreRemoveWorksheet,
  renameChartsheet as coreRenameChartsheet,
  replaceChartsheetChart as coreReplaceChartsheetChart,
  unprotectWorkbook,
  unregisterFunction as coreUnregisterFunction,
  verifyWorkbookPassword
} from "@excel/core/workbook.browser";
import type { Worksheet } from "@excel/core/worksheet";
import { checkAutoFilter } from "@excel/surface/argument-checks";
import type { StreamWorkbook, StreamWorksheet } from "@excel/surface/stream-targets";
import {
  asStreamWorkbook,
  refuseStreamWorkbook,
  workbookOnly
} from "@excel/surface/stream-targets";
import { setModel as surfaceSetModel } from "@excel/surface/workbook-model";
import type {
  AddWorksheetOptions,
  AutoFilter,
  HeaderFooter,
  NamedStyle,
  PageSetup,
  WorkbookProtection,
  WorksheetProperties,
  WorksheetState,
  WorksheetView
} from "@excel/types";
import { argumentChecks, checkProtectCall, checkSchema } from "@excel/utils/arguments";
import { copyPlain } from "@excel/utils/copy-plain";
import { splitPasswordOptions } from "@excel/utils/sheet-protection";
import { checkSheetSettings } from "@excel/utils/sheet-settings-checks";
import { isRecordArgument } from "@utils/object";
import { s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";

// The sheet settings and the auto-filter are checked by `checkSheetSettings` and `checkAutoFilter`, which the
// readers share; the schema pins the option names and checks the rest.
const addWorksheetOptionsSchema = s.object({
  properties: s.unknown<Partial<WorksheetProperties>>().optional(),
  pageSetup: s.unknown<Partial<PageSetup>>().optional(),
  headerFooter: s.unknown<Partial<HeaderFooter>>().optional(),
  views: s.unknown<Array<Partial<WorksheetView>>>().optional(),
  state: s.unknown<WorksheetState>().optional(),
  useSharedStrings: s.boolean().optional(),
  autoFilter: s.unknown<AutoFilter>().optional()
});
const _addWorksheetOptionsMatch: SchemaMatches<
  typeof addWorksheetOptionsSchema,
  AddWorksheetOptions
> = true;
void _addWorksheetOptionsMatch;

/**
 * Add a worksheet and return it. The options are checked as the writers check the same fields, so a `pageSetup`,
 * `views` or `state` value the format cannot hold is refused here rather than when the workbook is written.
 */
export function addWorksheet(
  wb: WorkbookData,
  name?: string,
  options?: AddWorksheetOptions
): Worksheet;
/** Add a sheet to a streaming workbook writer, and return its `Stream.WorksheetWriter`. */
export function addWorksheet(
  wb: StreamWorkbook,
  name?: string,
  options?: AddWorksheetOptions
): StreamWorksheet;
// The document model's signature again, last: `ReturnType<typeof …>` reads the last overload, and code that names a
// worksheet that way must keep meaning the document model's. (It is also first, which is what an `any` argument picks.)
export function addWorksheet(
  wb: WorkbookData,
  name?: string,
  options?: AddWorksheetOptions
): Worksheet;
export function addWorksheet(
  wb: WorkbookData | StreamWorkbook,
  name?: string,
  options?: AddWorksheetOptions
): Worksheet | StreamWorksheet {
  refuseLongSheetName(name);
  checkSchema("Workbook.addWorksheet", "options", addWorksheetOptionsSchema.optional(), options);
  if (options) {
    checkSheetSettings("Workbook.addWorksheet", options, path => `options.${path}`);
    if (options.autoFilter !== undefined && options.autoFilter !== null) {
      checkAutoFilter("Workbook.addWorksheet", "options.autoFilter", options.autoFilter);
    }
  }
  const writer = asStreamWorkbook(wb);
  if (writer) {
    return writer._addWorksheet(name, options) as unknown as StreamWorksheet;
  }
  return coreAddWorksheet(wb as WorkbookData, name, options);
}

/**
 * The sheet with id `id` or name `name` (compared as Excel compares sheet names), or the first sheet when `id` is
 * omitted; `undefined` when there is none. A streaming workbook writer answers with its `Stream.WorksheetWriter`.
 */
export function getWorksheet(wb: WorkbookData, id?: number | string): Worksheet | undefined;
export function getWorksheet(wb: StreamWorkbook, id?: number | string): StreamWorksheet | undefined;
// The document model's signature again, last: `ReturnType<typeof …>` reads the last overload, and code that names a
// worksheet that way must keep meaning the document model's. (It is also first, which is what an `any` argument picks.)
export function getWorksheet(wb: WorkbookData, id?: number | string): Worksheet | undefined;
export function getWorksheet(
  wb: WorkbookData | StreamWorkbook,
  id?: number | string
): Worksheet | StreamWorksheet | undefined {
  const writer = asStreamWorkbook(wb);
  if (writer) {
    return writer._getWorksheet(id) as unknown as StreamWorksheet | undefined;
  }
  return coreGetWorksheet(wb as WorkbookData, id);
}

/** The workbook's worksheets in tab order — a streaming writer's sheets, in the order they were added. */
export function listWorksheets(wb: WorkbookData): Worksheet[];
export function listWorksheets(wb: StreamWorkbook): StreamWorksheet[];
// The document model's signature again, last: `ReturnType<typeof …>` reads the last overload, and code that names a
// worksheet that way must keep meaning the document model's. (It is also first, which is what an `any` argument picks.)
export function listWorksheets(wb: WorkbookData): Worksheet[];
export function listWorksheets(wb: WorkbookData | StreamWorkbook): Worksheet[] | StreamWorksheet[] {
  const writer = asStreamWorkbook(wb);
  if (writer) {
    return writer._listWorksheets() as unknown as StreamWorksheet[];
  }
  return coreListWorksheets(wb as WorkbookData);
}

/** Call `callback` with each worksheet and its id, in tab order — a streaming writer's sheets too. */
export function eachWorksheet(
  wb: WorkbookData,
  callback: (sheet: Worksheet, id: number) => void
): void;
export function eachWorksheet(
  wb: StreamWorkbook,
  callback: (sheet: StreamWorksheet, id: number) => void
): void;
// The document model's signature again, last: `ReturnType<typeof …>` reads the last overload, and code that names a
// worksheet that way must keep meaning the document model's. (It is also first, which is what an `any` argument picks.)
export function eachWorksheet(
  wb: WorkbookData,
  callback: (sheet: Worksheet, id: number) => void
): void;
export function eachWorksheet(
  wb: WorkbookData | StreamWorkbook,
  callback:
    | ((sheet: Worksheet, id: number) => void)
    | ((sheet: StreamWorksheet, id: number) => void)
): void {
  const writer = asStreamWorkbook(wb);
  if (!writer) {
    coreEachWorksheet(wb as WorkbookData, callback as (sheet: Worksheet, id: number) => void);
    return;
  }
  if (typeof callback !== "function") {
    argumentChecks("Workbook.eachWorksheet").fail("callback", "a function", callback);
  }
  for (const sheet of writer._listWorksheets() as unknown as StreamWorksheet[]) {
    (callback as (sheet: StreamWorksheet, id: number) => void)(sheet, sheet.id);
  }
}

/** The workbook's defined names — a streaming writer's too, which it writes when it is committed. */
export function getDefinedNames(wb: WorkbookData | StreamWorkbook): DefinedNamesData {
  const writer = asStreamWorkbook(wb);
  return writer ? writer._getDefinedNames() : coreGetDefinedNames(wb as WorkbookData);
}

/**
 * Add (or replace) a workbook-level named cell style such as "Heading 1", and return a copy of the stored entry. A
 * name equal to an earlier one ignoring case is that style: it is replaced, and keeps the earlier spelling. On a
 * streaming writer, add it before the rows that use it are committed; nothing is registered with `useStyles: false`.
 */
export function addCellStyle(
  wb: WorkbookData | StreamWorkbook,
  name: string,
  style: NamedStyle
): NamedStyleEntry {
  const writer = asStreamWorkbook(wb);
  return writer
    ? structuredClone(writer._addCellStyle("Workbook.addCellStyle", name, style))
    : coreAddCellStyle(wb as WorkbookData, name, style);
}

/**
 * Define one of Excel's built-in cell-style presets (e.g. "Heading1") and return its style name (e.g. "Heading 1")
 * for `Cell.applyStyle` or `Stream.applyCellStyle`.
 */
export function useBuiltInCellStyle(
  wb: WorkbookData | StreamWorkbook,
  id: BuiltInCellStyle
): string {
  const writer = asStreamWorkbook(wb);
  return writer
    ? writer._useBuiltInCellStyle("Workbook.useBuiltInCellStyle", id)
    : coreUseBuiltInCellStyle(wb as WorkbookData, id);
}

/** A copy of the named cell style `name` (in any case), or `undefined`. */
export function getCellStyle(
  wb: WorkbookData | StreamWorkbook,
  name: string
): NamedStyleEntry | undefined {
  const writer = asStreamWorkbook(wb);
  return writer ? writer._getCellStyle(name) : coreGetCellStyle(wb as WorkbookData, name);
}

/** Every workbook-level named cell style in definition order, as copies. */
export function listCellStyles(wb: WorkbookData | StreamWorkbook): NamedStyleEntry[] {
  const writer = asStreamWorkbook(wb);
  return writer ? writer._listCellStyles() : coreListCellStyles(wb as WorkbookData);
}

/** Protect the workbook's structure or windows: `protect(wb, { password: "pw", lockStructure: true })`. */
export function protect(
  wb: WorkbookData | StreamWorkbook,
  options?: ProtectWorkbookOptions
): Promise<void>;
/**
 * @deprecated Pass the password in the options object: `protect(wb, { password, ...options })`.
 * Removed in 1.0.
 */
export function protect(
  wb: WorkbookData | StreamWorkbook,
  password: string | undefined,
  options?: Partial<WorkbookProtection>
): Promise<void>;
export function protect(
  wb: WorkbookData | StreamWorkbook,
  passwordOrOptions?: string | ProtectWorkbookOptions,
  legacyOptions?: Partial<WorkbookProtection>
): Promise<void> {
  const writer = asStreamWorkbook(wb);
  try {
    if (!writer) {
      checkWorkbook(wb);
    }
    checkProtectCall("Workbook.protect", passwordOrOptions, legacyOptions, "workbook");
  } catch (error) {
    return Promise.reject(error);
  }
  if (writer) {
    const [password, options] = splitPasswordOptions(passwordOrOptions, legacyOptions);
    return writer._protect(password, options);
  }
  return protectWorkbook(
    wb as WorkbookData,
    passwordOrOptions as string | undefined,
    legacyOptions
  );
}

/** Remove the workbook's structure and window protection, as {@link protect} set it — a streaming writer's too. */
export function unprotect(wb: WorkbookData | StreamWorkbook): void {
  const writer = asStreamWorkbook(wb);
  if (writer) {
    writer.unprotect();
    return;
  }
  unprotectWorkbook(wb as WorkbookData);
}

/**
 * Check `password` against the workbook's protection, as {@link protect} set it. `false` when the workbook is not
 * protected, or was protected without a password.
 */
export function verifyPassword(
  wb: WorkbookData | StreamWorkbook,
  password: string
): Promise<boolean> {
  // The writer keeps its protection where a workbook does, so the one check serves both.
  return verifyWorkbookPassword(wb as WorkbookData, password);
}

const registerPersonOptionsSchema = s.object({
  userId: s.string().optional(),
  providerId: s.string().optional()
});
const _registerPersonOptionsMatch: SchemaMatches<
  typeof registerPersonOptionsSchema,
  RegisterPersonOptions
> = true;
void _registerPersonOptionsMatch;

const addExternalLinkOptionsSchema = s.object({
  target: s.string({ nonEmpty: true }),
  sheetNames: s.readonlyArray(s.string({ nonEmpty: true })).optional(),
  // Each sheet's cached values are kept as given; only the map's shape is checked.
  cachedValues: s.record(s.unknown<ExternalLinkCachedSheet>()).optional(),
  targetMode: s.literal("External", "Internal").optional()
});
const _addExternalLinkOptionsMatch: SchemaMatches<
  typeof addExternalLinkOptionsSchema,
  AddExternalLinkOptions
> = true;
void _addExternalLinkOptionsMatch;

/**
 * Register a threaded-comment author and return their person id. Registering the same display name and user id
 * again returns the existing id.
 */
export function registerPerson(
  wb: WorkbookData,
  displayName: string,
  options?: RegisterPersonOptions
): string;
/**
 * @deprecated Pass the identity in the options object: `registerPerson(wb, name, { userId, providerId })`.
 * Removed in 1.0.
 */
export function registerPerson(
  wb: WorkbookData,
  displayName: string,
  userId?: string,
  providerId?: string
): string;
export function registerPerson(
  wb: WorkbookData,
  displayName: string,
  userIdOrOptions?: string | RegisterPersonOptions,
  legacyProviderId?: string
): string {
  refuseStreamWorkbook(
    wb,
    "Workbook.registerPerson",
    "writes no threaded comments, so it has no people to register"
  );
  const check = argumentChecks("Workbook.registerPerson");
  check.required("displayName", displayName);
  check.string("displayName", displayName);
  // The deprecated positional form stores its identity as given, as it always did.
  if (isRecordArgument(userIdOrOptions)) {
    checkSchema("Workbook.registerPerson", "options", registerPersonOptionsSchema, userIdOrOptions);
  }
  return coreRegisterPerson(
    wb,
    displayName,
    userIdOrOptions as string | undefined,
    legacyProviderId
  );
}

/**
 * Declare a link to another workbook and return it. `target` is the file — a bare name or a path relative to this
 * workbook (`"data/ref.xlsx"`) — and must be a non-empty string: it becomes the relationship's target.
 *
 * The link returned is the workbook's own, as `getExternalLink` and `listExternalLinks` return it: an edit to it —
 * filling in `cachedValues`, say — reaches the file, and is checked when the workbook is written.
 */
export function addExternalLink(
  wb: WorkbookData,
  input: AddExternalLinkOptions
): ExternalLinkModel {
  refuseStreamWorkbook(wb, "Workbook.addExternalLink", "writes no external-link parts");
  argumentChecks("Workbook.addExternalLink").required("input", input);
  checkSchema("Workbook.addExternalLink", "input", addExternalLinkOptionsSchema, input);
  return coreAddExternalLink(wb, input);
}

// --- what only a workbook in memory can do ---

const KEEPS_NOTHING = "keeps only what it has not written yet, and has no";

/** {@link coreRemoveWorksheet}, refusing a streaming writer: a sheet it has added may already be in the file. */
export const removeWorksheet = /* @__PURE__ */ workbookOnly(
  "Workbook.removeWorksheet",
  "cannot remove a sheet: it may already have been written",
  coreRemoveWorksheet
);
/** A streaming writer has no chartsheets: they draw from cells it has already written. */
export const addChartsheet = /* @__PURE__ */ workbookOnly(
  "Workbook.addChartsheet",
  "cannot add a chartsheet: its chart reads cells that have already been written",
  coreAddChartsheet
);
/** A streaming writer has no chartsheets: a pivot chart needs a pivot table it cannot hold. */
export const addPivotChartsheet = /* @__PURE__ */ workbookOnly(
  "Workbook.addPivotChartsheet",
  "cannot add a pivot chartsheet: it holds no pivot tables",
  coreAddPivotChartsheet
);
/** A streaming writer has no chartsheets. */
export const getChartsheet = /* @__PURE__ */ workbookOnly(
  "Workbook.getChartsheet",
  "has no chartsheets",
  coreGetChartsheet
);
/** A streaming writer has no chartsheets. */
export const listChartsheets = /* @__PURE__ */ workbookOnly(
  "Workbook.listChartsheets",
  "has no chartsheets",
  coreListChartsheets
);
/** A streaming writer has no chartsheets. */
export const removeChartsheet = /* @__PURE__ */ workbookOnly(
  "Workbook.removeChartsheet",
  "has no chartsheets",
  coreRemoveChartsheet
);
/** A streaming writer has no chartsheets. */
export const copyChartsheet = /* @__PURE__ */ workbookOnly(
  "Workbook.copyChartsheet",
  "has no chartsheets",
  coreCopyChartsheet
);
/** A streaming writer has no chartsheets. */
export const replaceChartsheetChart = /* @__PURE__ */ workbookOnly(
  "Workbook.replaceChartsheetChart",
  "has no chartsheets",
  coreReplaceChartsheetChart
);
/** @deprecated Use `Chartsheet.setName(Workbook.getChartsheet(wb, nameOrIndex), name)`. Removed in 1.0. */
export const renameChartsheet = /* @__PURE__ */ workbookOnly(
  "Workbook.renameChartsheet",
  "has no chartsheets",
  coreRenameChartsheet
);
/** Copying a whole sheet into a streaming writer would write it out of row order; add its rows instead. */
export const importWorksheet = /* @__PURE__ */ workbookOnly(
  "Workbook.importWorksheet",
  "cannot import a sheet: add its rows to a sheet of the writer instead",
  coreImportWorksheet
);
/** A streaming writer was not read from a file, so it holds no unrecognised values. */
export const listUnrecognizedValues = /* @__PURE__ */ workbookOnly(
  "Workbook.listUnrecognizedValues",
  "was not read from a file, so it holds no unrecognized values",
  coreListUnrecognizedValues
);
/** A streaming writer was not read from a file, so it holds no unrecognised values. */
export const discardUnrecognizedValues = /* @__PURE__ */ workbookOnly(
  "Workbook.discardUnrecognizedValues",
  "was not read from a file, so it holds no unrecognized values",
  coreDiscardUnrecognizedValues
);
/** A streaming writer writes no external-link parts. */
export const getExternalLink = /* @__PURE__ */ workbookOnly(
  "Workbook.getExternalLink",
  `${KEEPS_NOTHING} external links`,
  coreGetExternalLink
);
/** A streaming writer writes no external-link parts. */
export const listExternalLinks = /* @__PURE__ */ workbookOnly(
  "Workbook.listExternalLinks",
  `${KEEPS_NOTHING} external links`,
  coreListExternalLinks
);
/** A streaming writer writes no threaded comments, so it has no people to list. */
export const listPersons = /* @__PURE__ */ workbookOnly(
  "Workbook.listPersons",
  "writes no threaded comments, so it has no people",
  coreListPersons
);
/** A streaming writer has no formula engine to register a function with. */
export const registerFunction = /* @__PURE__ */ workbookOnly(
  "Workbook.registerFunction",
  "does not calculate formulas, so it has no functions to register",
  coreRegisterFunction
);
/** A streaming writer has no formula engine to register a function with. */
export const unregisterFunction = /* @__PURE__ */ workbookOnly(
  "Workbook.unregisterFunction",
  "does not calculate formulas, so it has no functions to register",
  coreUnregisterFunction
);
/**
 * A copy of the workbook's model: changing it changes nothing; apply a changed copy with {@link setModel}. The bytes it
 * carries (media, chart parts) are the workbook's own, not copies: read them, do not write into them. A streaming writer cannot describe what it has already written; it has no model.
 *
 * Copied as `Worksheet.getModel` is, and for the same reason: the model is built for the writer and shares the
 * workbook's own records — its properties, protection, views, media and charts — so an edit to it changed the workbook
 * with none of the setters' checks, and `setModel` on a second workbook left the two sharing them.
 */
export const getModel = /* @__PURE__ */ workbookOnly(
  "Workbook.getModel",
  `${KEEPS_NOTHING} model of what it has written`,
  publicWorkbookModel
);

function publicWorkbookModel(wb: WorkbookData): WorkbookModel {
  // The defined-names registry rides along for the writer alone; handed out, its name set was the workbook's own.
  const { definedNamesInstance: _registry, ...model } = coreGetWorkbookModel(wb);
  return copyPlain(model);
}
/** A streaming writer cannot be replaced by a model: what it has written stays written. */
export const setModel = /* @__PURE__ */ workbookOnly(
  "Workbook.setModel",
  "cannot take a model: what it has written stays written",
  surfaceSetModel
);
/** A named style a committed row uses cannot be taken back, so a streaming writer keeps every style it was given. */
export const removeCellStyle = /* @__PURE__ */ workbookOnly(
  "Workbook.removeCellStyle",
  "cannot remove a cell style: rows already written may use it",
  coreRemoveCellStyle
);

/** @deprecated Use {@link listWorksheets}. Removed in 1.0. */
export const getWorksheets = listWorksheets;
/** @deprecated Use {@link listChartsheets}. Removed in 1.0. */
export const getChartsheets = listChartsheets;
/** @deprecated Use {@link eachWorksheet}. Removed in 1.0. */
export const eachSheet = eachWorksheet;
/** @deprecated Use {@link importWorksheet}. Removed in 1.0. */
export const importSheet = importWorksheet;
/** @deprecated Use {@link addCellStyle}, which returns (a copy of) the stored entry. Removed in 1.0. */
export const defineCellStyle = addCellStyle;
/** @deprecated Use {@link useBuiltInCellStyle}. Removed in 1.0. */
export const useBuiltinCellStyle = useBuiltInCellStyle;
