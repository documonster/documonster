/**
 * The vocabulary of the worksheet settings that are plain assignable fields — `pageSetup` (with its `margins`),
 * `headerFooter`, `properties`, `views` and `state` — and the two things done with it.
 *
 * A field a caller assigns directly (`ws.pageSetup.orientation = "sideways"`) passes through no function that could
 * check it, so a value outside the file format's vocabulary was written as given and Excel repaired the file. The
 * writers therefore run {@link checkSheetSettings} before they write a sheet, and the failure is a typed error naming
 * the sheet and the field; `Workbook.addWorksheet` runs it on its options, so a value passed there fails at the call.
 *
 * A value read from a file is not the caller's to fix, and a vocabulary narrower than the file's must not lose what
 * the file said, so the readers run {@link setAsideInvalidSheetSettings} on what they parsed: an out-of-vocabulary
 * value leaves its field absent and is kept, as read, on the object that held it (`@utils/unrecognized-values`).
 * The writers put it back with {@link withKeptSheetSettings} while the field is still absent, and
 * {@link listKeptSheetSettings} reports it. One definition serves all three, so a read can never produce a value the
 * next write refuses, and never loses one either.
 *
 * Only what the format itself rules out is refused: enumerations, booleans, and the integers and references the
 * schema types require. Ranges Excel's dialogs enforce but its reader may tolerate are left alone.
 *
 * This is a visitor rather than an `@utils/schema` declaration because a read needs every value outside the vocabulary
 * — each one set aside in place, on the object that held it — where a schema reports the first. Its wording is the
 * schema's, so a refusal here reads as one from a schema-checked option does.
 */
import { ExcelInvalidAddressError } from "@excel/errors";
import { argumentFailure } from "@excel/utils/arguments";
import { describeValue as describe } from "@utils/describe-value";
import {
  UNRECOGNIZED_KEY,
  discardUnrecognizedValues,
  listUnrecognizedValues,
  setAsideValue
} from "@utils/unrecognized-values";

/** The settings a sheet carries as plain fields. Every member is optional and may be read from a file. */
export interface SheetSettings {
  readonly state?: unknown;
  readonly pageSetup?: unknown;
  readonly headerFooter?: unknown;
  readonly properties?: unknown;
  readonly views?: unknown;
}

/** ST_SheetState. */
export const SHEET_STATES: readonly string[] = ["visible", "hidden", "veryHidden"];
/** ST_Orientation. */
export const ORIENTATIONS: readonly string[] = ["default", "portrait", "landscape"];
/** ST_PageOrder. */
export const PAGE_ORDERS: readonly string[] = ["downThenOver", "overThenDown"];
/**
 * ST_CellComments, plus `"None"` — the model's default — and `"asDisplyed"`, the spelling this library once used.
 */
export const CELL_COMMENTS: readonly string[] = [
  "None",
  "none",
  "asDisplayed",
  "asDisplyed",
  "atEnd"
];
/** ST_PrintError. */
export const PRINT_ERRORS: readonly string[] = ["displayed", "blank", "dash", "NA"];
/** The model's view states: ST_PaneState's `frozenSplit` is read as `frozen`, and a sheet without a pane is `normal`. */
export const VIEW_STATES: readonly string[] = ["normal", "frozen", "split"];
/** ST_SheetViewType. */
export const VIEW_STYLES: readonly string[] = ["normal", "pageBreakPreview", "pageLayout"];
/** ST_Pane. */
export const PANES: readonly string[] = ["topLeft", "topRight", "bottomLeft", "bottomRight"];

const ARGB = /^(?:[0-9A-Fa-f]{2})?[0-9A-Fa-f]{6}$/;
const CELL = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/;
const MAX_COLUMN_SPLIT = 16383;
const MAX_ROW_SPLIT = 1048575;

type Fields = Record<string, unknown>;

/** What happens to a value outside the vocabulary: the writer refuses it, a reader drops it. */
interface Sink {
  invalid(container: Fields, key: string, path: string, expected: string, address?: boolean): void;
}

function isRecord(value: unknown): value is Fields {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isCell(value: string): boolean {
  const m = CELL.exec(value);
  if (!m) {
    return false;
  }
  let col = 0;
  for (const ch of m[1].toUpperCase()) {
    col = col * 26 + ch.charCodeAt(0) - 64;
  }
  const row = Number(m[2]);
  return col <= 16384 && row >= 1 && row <= 1048576;
}

/** The field checks, each skipping an absent (`undefined`) field. */
function fieldChecks(sink: Sink) {
  return {
    oneOf(c: Fields, key: string, path: string, allowed: readonly string[]): void {
      const v = c[key];
      if (v !== undefined && (typeof v !== "string" || !allowed.includes(v))) {
        sink.invalid(c, key, path, `one of ${allowed.map(a => JSON.stringify(a)).join(", ")}`);
      }
    },
    bool(c: Fields, key: string, path: string): void {
      const v = c[key];
      if (v !== undefined && typeof v !== "boolean") {
        sink.invalid(c, key, path, "true or false");
      }
    },
    str(c: Fields, key: string, path: string, nullable = false): void {
      const v = c[key];
      if (v !== undefined && !(nullable && v === null) && typeof v !== "string") {
        sink.invalid(c, key, path, "a string");
      }
    },
    num(
      c: Fields,
      key: string,
      path: string,
      { min, max, integer }: { min?: number; max?: number; integer?: boolean } = {}
    ): void {
      const v = c[key];
      if (v === undefined) {
        return;
      }
      if (
        typeof v !== "number" ||
        !Number.isFinite(v) ||
        (integer && !Number.isInteger(v)) ||
        (min !== undefined && v < min) ||
        (max !== undefined && v > max)
      ) {
        const noun = integer ? "an integer" : "a finite number";
        const range =
          min !== undefined && max !== undefined
            ? ` from ${min} to ${max}`
            : min !== undefined
              ? ` >= ${min}`
              : "";
        sink.invalid(c, key, path, `${noun}${range}`);
      }
    },
    cell(c: Fields, key: string, path: string): void {
      const v = c[key];
      // An empty string has always meant "none" here.
      if (v !== undefined && v !== "" && (typeof v !== "string" || !isCell(v))) {
        sink.invalid(c, key, path, "a cell such as B2 inside A1:XFD1048576", true);
      }
    },
    /** The field as a record to check further, or `undefined` when it is absent, `null`, or refused. */
    record(c: Fields, key: string, path: string): Fields | undefined {
      const v = c[key];
      if (v === undefined || v === null) {
        return undefined;
      }
      if (!isRecord(v)) {
        sink.invalid(c, key, path, "an object");
        return undefined;
      }
      return v;
    }
  };
}

function visitSheetSettings(sheet: Fields, sink: Sink): void {
  const f = fieldChecks(sink);
  f.oneOf(sheet, "state", "state", SHEET_STATES);

  const setup = f.record(sheet, "pageSetup", "pageSetup");
  if (setup) {
    const margins = f.record(setup, "margins", "pageSetup.margins");
    if (margins) {
      for (const side of ["left", "right", "top", "bottom", "header", "footer"]) {
        f.num(margins, side, `pageSetup.margins.${side}`);
      }
    }
    f.oneOf(setup, "orientation", "pageSetup.orientation", ORIENTATIONS);
    f.oneOf(setup, "pageOrder", "pageSetup.pageOrder", PAGE_ORDERS);
    f.oneOf(setup, "cellComments", "pageSetup.cellComments", CELL_COMMENTS);
    f.oneOf(setup, "errors", "pageSetup.errors", PRINT_ERRORS);
    for (const count of [
      "horizontalDpi",
      "verticalDpi",
      "fitToWidth",
      "fitToHeight",
      "scale",
      "paperSize",
      "copies"
    ]) {
      f.num(setup, count, `pageSetup.${count}`, { min: 0, integer: true });
    }
    f.num(setup, "firstPageNumber", "pageSetup.firstPageNumber", { integer: true });
    for (const flag of [
      "fitToPage",
      "blackAndWhite",
      "draft",
      "showRowColHeaders",
      "showGridLines",
      "horizontalCentered",
      "verticalCentered",
      "useFirstPageNumber",
      "usePrinterDefaults"
    ]) {
      f.bool(setup, flag, `pageSetup.${flag}`);
    }
  }

  const headerFooter = f.record(sheet, "headerFooter", "headerFooter");
  if (headerFooter) {
    for (const flag of ["differentFirst", "differentOddEven", "scaleWithDoc", "alignWithMargins"]) {
      f.bool(headerFooter, flag, `headerFooter.${flag}`);
    }
    for (const text of [
      "oddHeader",
      "oddFooter",
      "evenHeader",
      "evenFooter",
      "firstHeader",
      "firstFooter"
    ]) {
      f.str(headerFooter, text, `headerFooter.${text}`, true);
    }
  }

  const properties = f.record(sheet, "properties", "properties");
  if (properties) {
    const color = f.record(properties, "tabColor", "properties.tabColor");
    if (color) {
      if (color.argb !== undefined && (typeof color.argb !== "string" || !ARGB.test(color.argb))) {
        sink.invalid(
          color,
          "argb",
          "properties.tabColor.argb",
          'six or eight hex digits (alpha first), such as "FFFF0000"'
        );
      }
      f.num(color, "theme", "properties.tabColor.theme", { min: 0, integer: true });
      f.num(color, "indexed", "properties.tabColor.indexed", { min: 0, integer: true });
      f.num(color, "tint", "properties.tabColor.tint", { min: -1, max: 1 });
    }
    // Excel's outline is seven levels deep, but both writers clamp a deeper one, so only the type is the format's.
    for (const level of ["outlineLevelCol", "outlineLevelRow"]) {
      f.num(properties, level, `properties.${level}`, { min: 0, integer: true });
    }
    f.num(properties, "defaultRowHeight", "properties.defaultRowHeight", { min: 0 });
    f.num(properties, "defaultColWidth", "properties.defaultColWidth", { min: 0 });
    f.num(properties, "dyDescent", "properties.dyDescent");
    f.bool(properties, "showGridLines", "properties.showGridLines");
    f.str(properties, "codeName", "properties.codeName");
    const outline = f.record(properties, "outlineProperties", "properties.outlineProperties");
    if (outline) {
      f.bool(outline, "summaryBelow", "properties.outlineProperties.summaryBelow");
      f.bool(outline, "summaryRight", "properties.outlineProperties.summaryRight");
    }
  }

  const views = sheet.views;
  if (views !== undefined && views !== null && !Array.isArray(views)) {
    sink.invalid(sheet, "views", "views", "an array");
  } else if (Array.isArray(views)) {
    views.forEach((view: unknown, i) => {
      const at = `views[${i}]`;
      if (!isRecord(view)) {
        sink.invalid(views as unknown as Fields, String(i), at, "an object");
        return;
      }
      f.oneOf(view, "state", `${at}.state`, VIEW_STATES);
      f.oneOf(view, "style", `${at}.style`, VIEW_STYLES);
      f.oneOf(view, "activePane", `${at}.activePane`, PANES);
      // A frozen pane's splits count columns and rows; a split pane's are a distance in twentieths of a point.
      if (view.state === "frozen") {
        f.num(view, "xSplit", `${at}.xSplit`, { min: 0, max: MAX_COLUMN_SPLIT, integer: true });
        f.num(view, "ySplit", `${at}.ySplit`, { min: 0, max: MAX_ROW_SPLIT, integer: true });
      } else {
        f.num(view, "xSplit", `${at}.xSplit`, { min: 0 });
        f.num(view, "ySplit", `${at}.ySplit`, { min: 0 });
      }
      f.cell(view, "topLeftCell", `${at}.topLeftCell`);
      f.cell(view, "activeCell", `${at}.activeCell`);
      for (const zoom of ["zoomScale", "zoomScaleNormal", "workbookViewId"]) {
        f.num(view, zoom, `${at}.${zoom}`, { min: 0, integer: true });
      }
      for (const flag of [
        "rightToLeft",
        "showRuler",
        "showRowColHeaders",
        "showGridLines",
        "tabSelected"
      ]) {
        f.bool(view, flag, `${at}.${flag}`);
      }
    });
  }
}

/**
 * Check a sheet's plain settings, refusing a value the file format cannot hold with an `ExcelArgumentError` (or an
 * `ExcelInvalidAddressError` for a pane's cell) whose message begins with `where`.
 */
export function checkSheetSettings(
  where: string,
  sheet: SheetSettings,
  field: (path: string) => string = path => path
): void {
  visitSheetSettings(sheet as Fields, {
    invalid(container, key, settingPath, expected, address) {
      // The setting's path as the caller named it: `options.pageSetup.orientation`, or `value.orientation` for a setter
      // that takes the one setting.
      const path = field(settingPath);
      if (address) {
        throw new ExcelInvalidAddressError(
          String(container[key]),
          `${where}: ${path} must be ${expected}, got ${describe(container[key])}`
        );
      }
      throw argumentFailure(where, path, expected, container[key]);
    }
  });
}

/**
 * Set aside, in place, every setting a reader parsed that the file format cannot hold: the field is removed and its
 * text kept on the object that held it, so the next write puts it back unchanged. A value with no text to keep — a
 * view that is not an object, a number that is not finite — is dropped. Returns one `Sheet: …` line per value dropped.
 */
export function setAsideInvalidSheetSettings(sheetName: string, sheet: object): string[] {
  const dropped: string[] = [];
  visitSheetSettings(sheet as Fields, {
    invalid(container, key, path, expected) {
      const value = container[key];
      if (Array.isArray(container)) {
        // A view that is not an object: keep the list dense, with a default view in its place.
        container[Number(key)] = {};
      } else {
        delete container[key];
      }
      const text =
        typeof value === "string"
          ? value
          : typeof value === "boolean" || (typeof value === "number" && Number.isFinite(value))
            ? String(value)
            : undefined;
      // The sheet's own fields (its state) are not carried as a container, so there is nowhere to keep one.
      if (text !== undefined && !Array.isArray(container) && container !== sheet) {
        setAsideValue(container, key, text);
      } else {
        dropped.push(
          `${sheetName}: ${path} ${describe(value)} is not ${expected}, and was read as absent`
        );
      }
    }
  });
  return dropped;
}

/** The containers a sheet's plain settings live in, by path. */
function settingContainers(sheet: Fields): [path: string, container: Fields][] {
  const found: [string, Fields][] = [];
  const add = (path: string, value: unknown): Fields | undefined => {
    if (isRecord(value)) {
      found.push([path, value]);
      return value;
    }
    return undefined;
  };
  const setup = add("pageSetup", sheet.pageSetup);
  add("pageSetup.margins", setup?.margins);
  add("headerFooter", sheet.headerFooter);
  const properties = add("properties", sheet.properties);
  add("properties.tabColor", properties?.tabColor);
  add("properties.outlineProperties", properties?.outlineProperties);
  if (Array.isArray(sheet.views)) {
    sheet.views.forEach((view: unknown, i) => add(`views[${i}]`, view));
  }
  return found;
}

/** A setting a reader kept as read: where it is, its text, and what the field takes instead. */
export interface KeptSheetSetting {
  /** `pageSetup.orientation`, `views[0].xSplit`, … */
  readonly path: string;
  readonly value: string;
  /** What the field takes, as the write check words it: `one of "default", "portrait", "landscape"`. */
  readonly expected: string;
}

/** Every setting a reader set aside on `sheet` and no one has set since. */
export function findKeptSheetSettings(sheet: object): KeptSheetSetting[] {
  const kept = new Map<string, string>();
  for (const [path, container] of settingContainers(sheet as Fields)) {
    for (const { field, value } of listUnrecognizedValues(container)) {
      kept.set(`${path}.${field}`, value);
    }
  }
  if (kept.size === 0) {
    return [];
  }
  const found: KeptSheetSetting[] = [];
  // Checked again with the text in place, which words each the way a write would.
  visitSheetSettings(withKeptSheetSettings(sheet as SheetSettings) as Fields, {
    invalid(_container, _key, path, expected) {
      const value = kept.get(path);
      if (value !== undefined) {
        found.push({ path, value, expected });
        kept.delete(path);
      }
    }
  });
  // A kept value the check now accepts — the vocabulary has grown since it was set aside — is still kept.
  for (const [path, value] of kept) {
    found.push({ path, value, expected: "a value of this setting" });
  }
  return found;
}

/**
 * Drop the settings a reader set aside on `sheet` — all, or those `filter` returns `true` for, given the path
 * {@link findKeptSheetSettings} reports (`pageSetup.orientation`) — and return how many were still in effect.
 */
export function discardKeptSheetSettings(
  sheet: object,
  filter?: (setting: { readonly path: string; readonly value: string }) => boolean
): number {
  let discarded = 0;
  for (const [path, container] of settingContainers(sheet as Fields)) {
    discarded += discardUnrecognizedValues(container, filter, path);
  }
  return discarded;
}

/**
 * `sheet`'s settings as a writer should emit them: where a field is still absent and a reader set its value aside,
 * a copy of the container with the kept text in the field. Containers with nothing kept are returned as they are, and
 * `sheet` itself is not changed. Run after {@link checkSheetSettings}, which never sees the kept text.
 */
export function withKeptSheetSettings<T extends SheetSettings>(sheet: T): T {
  const restore = (container: unknown): unknown => {
    if (!isRecord(container)) {
      return container;
    }
    const kept = listUnrecognizedValues(container);
    if (kept.length === 0) {
      return container;
    }
    const copy: Fields = { ...container };
    for (const { field, value } of kept) {
      copy[field] = value;
    }
    delete copy[UNRECOGNIZED_KEY];
    return copy;
  };
  const source = sheet as Fields;
  let setup = source.pageSetup;
  if (isRecord(setup)) {
    const margins = restore(setup.margins);
    setup = restore(setup);
    if (margins !== (setup as Fields).margins) {
      setup = { ...(setup as Fields), margins };
    }
  }
  let properties = source.properties;
  if (isRecord(properties)) {
    const tabColor = restore(properties.tabColor);
    const outline = restore(properties.outlineProperties);
    properties = restore(properties);
    if (
      tabColor !== (properties as Fields).tabColor ||
      outline !== (properties as Fields).outlineProperties
    ) {
      properties = { ...(properties as Fields), tabColor, outlineProperties: outline };
    }
  }
  const views = Array.isArray(source.views) ? source.views.map(restore) : source.views;
  const headerFooter = restore(source.headerFooter);
  if (
    setup === source.pageSetup &&
    properties === source.properties &&
    headerFooter === source.headerFooter &&
    (!Array.isArray(views) || views.every((view, i) => view === (source.views as unknown[])[i]))
  ) {
    return sheet;
  }
  return { ...sheet, pageSetup: setup, properties, headerFooter, views } as T;
}
