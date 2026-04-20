import type { Column } from "@excel/column";
import type { DataValidations } from "@excel/data-validations";
import type { ValueType, FormulaType } from "@excel/enums";
import { Enums } from "@excel/enums";
import { ExcelError, InvalidValueTypeError } from "@excel/errors";
import { Note } from "@excel/note";
import type { Row } from "@excel/row";
import type {
  Style,
  NumFmt,
  Font,
  Alignment,
  Protection,
  Borders,
  Fill,
  CellRichTextValue,
  CellErrorValue,
  DataValidation,
  CellValue,
  CellValueInput,
  CellHyperlinkValue,
  CellHyperlinkValueInput,
  CellFormulaHyperlinkValue,
  CellCheckboxValue,
  RichText
} from "@excel/types";
import { getCellDisplayText } from "@excel/utils/cell-format";
import { colCache } from "@excel/utils/col-cache";
import { copyStyle } from "@excel/utils/copy-style";
import { slideFormula } from "@excel/utils/shared-formula";
import { escapeHtml } from "@excel/utils/under-dash";
import type { Workbook } from "@excel/workbook";
import type { Worksheet } from "@excel/worksheet";

export type FormulaResult = string | number | boolean | Date | CellErrorValue;

// Extended formula type for internal use (includes shared formula fields)
export interface FormulaValueData {
  shareType?: string;
  ref?: string;
  formula?: string;
  sharedFormula?: string;
  result?: FormulaResult;
  date1904?: boolean;
  isDynamicArray?: boolean;
}

// FullAddress for Cell - only needs basic fields for defined names
interface FullAddress {
  sheetName: string;
  address: string;
  row: number;
  col: number;
}

export interface CellAddress {
  address: string;
  row: number;
  col: number;
  $col$row?: string;
}

export interface NoteText {
  text: string;
  font?: Partial<Font>;
}

export interface NoteConfig {
  texts?: NoteText[];
  margins?: { insetmode?: string; inset?: number[] };
  protection?: { locked?: string; lockText?: string };
  editAs?: string;
  anchor?: string;
}

export interface NoteModel {
  type: string;
  note: NoteConfig;
}

export interface CellModel {
  address: string;
  type: ValueType;
  // Internal value storage - type depends on cell type
  value?:
    | number
    | string
    | boolean
    | Date
    | CellRichTextValue
    | CellErrorValue
    | CellHyperlinkValue;
  style?: Partial<Style>;
  comment?: NoteModel;
  text?: string;
  hyperlink?: string;
  tooltip?: string;
  master?: string;
  shareType?: string;
  ref?: string;
  formula?: string;
  sharedFormula?: string;
  result?: FormulaResult;
  /**
   * Rich-text runs associated with this cell.
   *
   * - When `type === RichText`, this holds a CellRichTextValue (object wrapping
   *   the runs array) — this is the historical, cell-level rich-text payload.
   * - When `type === Hyperlink`, this holds a plain RichText[] — the runs used
   *   for formatted display of the hyperlink text.
   *
   * Callers should branch on `type` before accessing the shape.
   */
  richText?: CellRichTextValue | RichText[];
  sharedString?: number;
  error?: CellErrorValue;
  rawValue?: unknown;
  isDynamicArray?: boolean;
}

// Internal interface for Value type objects
interface ICellValue {
  model: CellModel;
  value: CellValueType;
  type: ValueType;
  effectiveType: ValueType;
  address: string;
  formula?: string;
  result?: FormulaResult;
  formulaType?: FormulaType;
  hyperlink?: string;
  master?: Cell;
  text?: string;
  release(): void;
  toCsvString(): string;
  toString(): string;
  isMergedTo?(master: Cell): boolean;
}

// Type for cell values returned to callers (read side)
export type CellValueType = CellValue;
// Type for cell values accepted from callers (write side)
export type CellValueInputType = CellValueInput;

// Returns true if the value is a non-empty object (has at least one own key),
// or any truthy non-object value. Returns false for undefined, null, false, 0,
// empty string, and empty objects `{}`. This is used to prevent an empty `{}`
// style property on a row from shadowing a real style property on a column.
const hasOwnKeys = (v: unknown): boolean =>
  !!v && (typeof v !== "object" || Object.keys(v as object).length > 0);

/**
 * Flatten a rich-text array into its plain-text representation by
 * concatenating each run's `text`. Missing/null runs contribute "".
 */
function flattenRichText(runs: readonly RichText[]): string {
  let out = "";
  for (const run of runs) {
    if (run && typeof run.text === "string") {
      out += run.text;
    }
  }
  return out;
}

/**
 * A `CellHyperlinkValue` with its invariants established:
 *   - `text` is always a string (no longer optional)
 *   - `hyperlink` is always a string (no longer optional)
 *   - when `richText` is present, it is a non-empty array and
 *     `text === flattenRichText(richText)`
 */
interface NormalizedHyperlink extends CellHyperlinkValue {
  text: string;
  hyperlink: string;
}

/**
 * Normalize a CellHyperlinkValue so the {@link NormalizedHyperlink} invariants
 * hold.
 *
 * - If the caller supplied `richText` but no `text`, text is derived.
 * - If the caller supplied `text` but no `richText`, richText stays absent.
 * - If both are supplied, `richText` wins and `text` is regenerated
 *   (to keep `text === flatten(richText)`).
 * - An empty `richText: []` is dropped (treated as "no rich text").
 */
function normalizeHyperlinkValue(value: CellHyperlinkValueInput): NormalizedHyperlink {
  let text: string;
  let richText: RichText[] | undefined;
  if (Array.isArray(value.richText) && value.richText.length > 0) {
    richText = value.richText;
    text = flattenRichText(richText);
  } else {
    text = typeof value.text === "string" ? value.text : "";
  }
  const out: NormalizedHyperlink = {
    text,
    hyperlink: typeof value.hyperlink === "string" ? value.hyperlink : ""
  };
  if (richText) {
    out.richText = richText;
  }
  if (typeof value.tooltip === "string" && value.tooltip.length > 0) {
    out.tooltip = value.tooltip;
  }
  return out;
}

// Cell requirements
//  Operate inside a worksheet
//  Store and retrieve a value with a range of types: text, number, date, hyperlink, reference, formula, etc.
//  Manage/use and manipulate cell format either as local to cell or inherited from column or row.

class Cell {
  static Types = Enums.ValueType;

  // Type declarations only - no runtime overhead
  declare private _row: Row;
  declare private _column: Column;
  declare private _address: string;

  declare private _value: ICellValue;
  declare public style: Partial<Style>;
  declare private _mergeCount: number;

  declare private _comment?: Note;

  constructor(row: Row, column: Column, address: string) {
    if (!row || !column) {
      throw new ExcelError("A Cell needs a Row");
    }

    this._row = row;
    this._column = column;

    colCache.validateAddress(address);
    this._address = address;

    // TODO: lazy evaluation of this._value
    this._value = Value.create(Cell.Types.Null, this);

    this.style = this._mergeStyle(row.style, column.style, {});

    this._mergeCount = 0;
  }

  get worksheet(): Worksheet {
    return this._row.worksheet;
  }

  get workbook(): Workbook {
    return this._row.worksheet.workbook;
  }

  // help GC by removing cyclic (and other) references
  destroy(): void {
    this.style = undefined!;
    this._value = undefined!;
    this._row = undefined!;
    this._column = undefined!;
    this._address = undefined!;
  }

  // =========================================================================
  // Styles stuff
  get numFmt(): string | NumFmt | undefined {
    return this.style.numFmt;
  }

  set numFmt(value: string | undefined) {
    this.style.numFmt = value;
  }

  get font(): Partial<Font> | undefined {
    return this.style.font;
  }

  set font(value: Partial<Font> | undefined) {
    this.style.font = value;
  }

  get alignment(): Partial<Alignment> | undefined {
    return this.style.alignment;
  }

  set alignment(value: Partial<Alignment> | undefined) {
    this.style.alignment = value;
  }

  get border(): Partial<Borders> | undefined {
    return this.style.border;
  }

  set border(value: Partial<Borders> | undefined) {
    this.style.border = value;
  }

  get fill(): Fill | undefined {
    return this.style.fill;
  }

  set fill(value: Fill | undefined) {
    this.style.fill = value;
  }

  get protection(): Partial<Protection> | undefined {
    return this.style.protection;
  }

  set protection(value: Partial<Protection> | undefined) {
    this.style.protection = value;
  }

  private _mergeStyle(
    rowStyle: Partial<Style>,
    colStyle: Partial<Style>,
    style: Partial<Style>
  ): Partial<Style> {
    const numFmt = (rowStyle && rowStyle.numFmt) || (colStyle && colStyle.numFmt);
    if (numFmt) {
      style.numFmt = numFmt;
    }

    const font =
      (rowStyle && hasOwnKeys(rowStyle.font) && rowStyle.font) ||
      (colStyle && hasOwnKeys(colStyle.font) && colStyle.font);
    if (font) {
      style.font = structuredClone(font);
    }

    const alignment =
      (rowStyle && hasOwnKeys(rowStyle.alignment) && rowStyle.alignment) ||
      (colStyle && hasOwnKeys(colStyle.alignment) && colStyle.alignment);
    if (alignment) {
      style.alignment = structuredClone(alignment);
    }

    const border =
      (rowStyle && hasOwnKeys(rowStyle.border) && rowStyle.border) ||
      (colStyle && hasOwnKeys(colStyle.border) && colStyle.border);
    if (border) {
      style.border = structuredClone(border);
    }

    const fill =
      (rowStyle && hasOwnKeys(rowStyle.fill) && rowStyle.fill) ||
      (colStyle && hasOwnKeys(colStyle.fill) && colStyle.fill);
    if (fill) {
      style.fill = structuredClone(fill);
    }

    const protection =
      (rowStyle && hasOwnKeys(rowStyle.protection) && rowStyle.protection) ||
      (colStyle && hasOwnKeys(colStyle.protection) && colStyle.protection);
    if (protection) {
      style.protection = structuredClone(protection);
    }

    return style;
  }

  // =========================================================================
  // return the address for this cell
  get address(): string {
    return this._address;
  }

  get row(): number {
    return this._row.number;
  }

  get col(): number {
    return this._column.number;
  }

  get $col$row(): string {
    return `$${this._column.letter}$${this.row}`;
  }

  // =========================================================================
  // Value stuff

  get type(): ValueType {
    return this._value.type;
  }

  get effectiveType(): ValueType {
    return this._value.effectiveType;
  }

  toCsvString(): string {
    return this._value.toCsvString();
  }

  // =========================================================================
  // Merge stuff

  addMergeRef(): void {
    this._mergeCount++;
  }

  releaseMergeRef(): void {
    this._mergeCount--;
  }

  get isMerged(): boolean {
    return this._mergeCount > 0 || this.type === Cell.Types.Merge;
  }

  merge(master: Cell, ignoreStyle?: boolean): void {
    this._value.release();
    this._value = Value.create(Cell.Types.Merge, this, master);
    if (!ignoreStyle) {
      // Deep-copy so each cell has an independent style object.
      // Without this, all cells in the merge share the same reference,
      // and setting a property (e.g. border) on any cell mutates all of them.
      this.style = (copyStyle(master.style) as Partial<Style>) ?? {};
    }
  }

  unmerge(): void {
    if (this.type === Cell.Types.Merge) {
      this._value.release();
      this._value = Value.create(Cell.Types.Null, this);
      this.style = this._mergeStyle(this._row.style, this._column.style, { ...this.style });
    }
  }

  isMergedTo(master: Cell): boolean {
    if (this._value.type !== Cell.Types.Merge) {
      return false;
    }
    return this._value.isMergedTo ? this._value.isMergedTo(master) : false;
  }

  get master(): Cell {
    if (this.type === Cell.Types.Merge) {
      return this._value.master!;
    }
    return this; // an unmerged cell is its own master
  }

  get isHyperlink(): boolean {
    return this._value.type === Cell.Types.Hyperlink;
  }

  get hyperlink(): string | undefined {
    return this._value.hyperlink;
  }

  // return the value
  get value(): CellValueType {
    return this._value.value;
  }

  // set the value - can be number, string or raw
  set value(v: CellValueInputType) {
    // special case - merge cells set their master's value
    if (this.type === Cell.Types.Merge) {
      this._value.master!.value = v;
      return;
    }

    this._value.release();

    // assign value
    this._value = Value.create(Value.getType(v), this, v);
  }

  get note(): string | NoteConfig | undefined {
    if (!this._comment) {
      return undefined;
    }
    const noteValue = this._comment.note;
    return noteValue;
  }

  set note(note: string | NoteConfig) {
    this._comment = new Note(note);
  }

  // Internal comment accessor for row operations
  get comment(): Note | undefined {
    return this._comment;
  }

  set comment(comment: Note | NoteConfig | undefined) {
    if (comment === undefined) {
      this._comment = undefined;
    } else if (comment instanceof Note) {
      this._comment = comment;
    } else {
      this._comment = new Note(comment);
    }
  }

  get text(): string {
    return this._value.toString();
  }

  /**
   * The cell's display text — the value formatted the way Excel would render
   * it, applying the cell's `numFmt`. For a Date cell with `numFmt` `"mm-dd-yy"`,
   * this returns e.g. `"04-12-19"` rather than the JS `Date.prototype.toString()`
   * output you'd get from `cell.text`.
   *
   * Handles primitive values, dates, and formula results. For rich text,
   * hyperlinks, errors, and other complex types, falls back to `cell.text`.
   *
   * Note: numFmt codes that are locale-dependent in Excel (e.g. built-in
   * numFmtId 14 renders as `dd.mm.yyyy` under German locale but is stored
   * as `mm-dd-yy`) are applied literally — excelts does not perform
   * Excel's locale-based format substitution. If you need a specific date
   * style across cells regardless of per-cell numFmts, call the exported
   * {@link getCellDisplayText} helper with a `dateFormat` argument, or use
   * `worksheet.toJSON({ dateFormat })`.
   */
  get displayText(): string {
    return getCellDisplayText(this);
  }

  get html(): string {
    return escapeHtml(this.text);
  }

  toString(): string {
    return this.text;
  }

  /** @internal */
  _upgradeToHyperlink(hyperlink: string): void {
    // Upgrade this cell to a Hyperlink while preserving the existing display
    // text. Supports promotion from both plain String cells and RichText cells.
    // For RichText cells, the runs are preserved on the new hyperlink value.
    switch (this.type) {
      case Cell.Types.String: {
        this._value = Value.create(Cell.Types.Hyperlink, this, {
          text: String(this._value.value),
          hyperlink
        });
        break;
      }
      case Cell.Types.RichText: {
        const current = this._value.value as CellRichTextValue | undefined;
        const runs = current && Array.isArray(current.richText) ? current.richText : [];
        this._value = Value.create(Cell.Types.Hyperlink, this, {
          text: flattenRichText(runs),
          richText: runs.length > 0 ? runs : undefined,
          hyperlink
        });
        break;
      }
      default:
        // Other cell types (Number, Date, Formula, ...) are not auto-upgraded.
        break;
    }
  }

  // =========================================================================
  // Formula stuff
  get formula(): string | undefined {
    return this._value.formula;
  }

  get result(): FormulaResult | undefined {
    return this._value.result;
  }

  set result(value: FormulaResult | undefined) {
    if (this.type === Cell.Types.Formula) {
      this._value.result = value;
    }
  }

  get formulaType(): FormulaType {
    return this._value.formulaType ?? Enums.FormulaType.None;
  }

  // =========================================================================
  // Name stuff
  get fullAddress(): FullAddress {
    const { worksheet } = this._row;
    return {
      sheetName: worksheet.name,
      address: this.address,
      row: this.row,
      col: this.col
    };
  }

  get name(): string {
    return this.names[0];
  }

  set name(value: string) {
    this.names = [value];
  }

  get names(): string[] {
    return this.workbook.definedNames.getNamesEx(this.fullAddress);
  }

  set names(value: string[]) {
    const { definedNames } = this.workbook;
    definedNames.removeAllNames(this.fullAddress);
    value.forEach(name => {
      definedNames.addEx(this.fullAddress, name);
    });
  }

  addName(name: string): void {
    this.workbook.definedNames.addEx(this.fullAddress, name);
  }

  removeName(name: string): void {
    this.workbook.definedNames.removeEx(this.fullAddress, name);
  }

  removeAllNames(): void {
    this.workbook.definedNames.removeAllNames(this.fullAddress);
  }

  // =========================================================================
  // Data Validation stuff
  private get _dataValidations(): DataValidations {
    return this.worksheet.dataValidations;
  }

  get dataValidation(): DataValidation | undefined {
    return this._dataValidations.find(this.address);
  }

  set dataValidation(value: DataValidation) {
    this._dataValidations.add(this.address, value);
  }

  // =========================================================================
  // Model stuff

  get model(): CellModel {
    const { model } = this._value;
    model.style = this.style;
    if (this._comment) {
      model.comment = this._comment.model;
    }
    return model;
  }

  set model(value: CellModel) {
    this._value.release();
    this._value = Value.create(value.type, this);
    this._value.model = value;

    if (value.comment) {
      switch (value.comment.type) {
        case "note":
          this._comment = Note.fromModel(value.comment);
          break;
      }
    }

    if (value.style) {
      this.style = (copyStyle(value.style) as Partial<Style>) ?? {};
    } else {
      this.style = {};
    }
  }
}

// =============================================================================
// Internal Value Types

// Internal model interfaces for type safety within Value classes
interface NullValueModel {
  address: string;
  type: number;
}

interface NumberValueModel {
  address: string;
  type: number;
  value: number;
}

interface StringValueModel {
  address: string;
  type: number;
  value: string;
}

interface DateValueModel {
  address: string;
  type: number;
  value: Date;
}

interface BooleanValueModel {
  address: string;
  type: number;
  value: boolean;
}

interface HyperlinkValueModel {
  address: string;
  type: number;
  text?: string;
  richText?: RichText[];
  hyperlink?: string;
  tooltip?: string;
}

interface MergeValueModel {
  address: string;
  type: number;
  master?: string;
}

interface FormulaValueModel {
  address: string;
  type: number;
  shareType?: string;
  ref?: string;
  formula?: string;
  sharedFormula?: string;
  result?: FormulaResult;
  isDynamicArray?: boolean;
}

interface SharedStringValueModel {
  address: string;
  type: number;
  value: number;
}

interface RichTextValueModel {
  address: string;
  type: number;
  value: CellRichTextValue;
}

interface ErrorValueModel {
  address: string;
  type: number;
  value: CellErrorValue;
}

interface JSONValueModel {
  address: string;
  type: number;
  value: string;
  rawValue: unknown;
}

class NullValue {
  declare public model: NullValueModel;

  constructor(cell: Cell) {
    this.model = {
      address: cell.address,
      type: Cell.Types.Null
    };
  }

  get value(): null {
    return null;
  }

  set value(_value: unknown) {
    // nothing to do
  }

  get type(): ValueType {
    return Cell.Types.Null;
  }

  get effectiveType(): ValueType {
    return Cell.Types.Null;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  toCsvString(): string {
    return "";
  }

  release(): void {}

  toString(): string {
    return "";
  }
}

class NumberValue {
  declare public model: NumberValueModel;

  constructor(cell: Cell, value: number) {
    this.model = {
      address: cell.address,
      type: Cell.Types.Number,
      value
    };
  }

  get value(): number {
    return this.model.value;
  }

  set value(value: number) {
    this.model.value = value;
  }

  get type(): ValueType {
    return Cell.Types.Number;
  }

  get effectiveType(): ValueType {
    return Cell.Types.Number;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  toCsvString(): string {
    return this.model.value.toString();
  }

  release(): void {}

  toString(): string {
    return this.model.value.toString();
  }
}

class StringValue {
  declare public model: StringValueModel;

  constructor(cell: Cell, value: string) {
    this.model = {
      address: cell.address,
      type: Cell.Types.String,
      value
    };
  }

  get value(): string {
    return this.model.value;
  }

  set value(value: string) {
    this.model.value = value;
  }

  get type(): ValueType {
    return Cell.Types.String;
  }

  get effectiveType(): ValueType {
    return Cell.Types.String;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  toCsvString(): string {
    return `"${this.model.value.replace(/"/g, '""')}"`;
  }

  release(): void {}

  toString(): string {
    return this.model.value;
  }
}

class RichTextValue {
  declare public model: RichTextValueModel;

  constructor(cell: Cell, value: CellRichTextValue) {
    this.model = {
      address: cell.address,
      type: Cell.Types.String,
      value
    };
  }

  get value(): CellRichTextValue {
    return this.model.value;
  }

  set value(value: CellRichTextValue) {
    this.model.value = value;
  }

  toString(): string {
    return this.model.value.richText.map(t => t.text).join("");
  }

  get type(): ValueType {
    return Cell.Types.RichText;
  }

  get effectiveType(): ValueType {
    return Cell.Types.RichText;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  get text(): string {
    return this.toString();
  }

  toCsvString(): string {
    return `"${this.text.replace(/"/g, '""')}"`;
  }

  release(): void {}
}

class DateValue {
  declare public model: DateValueModel;

  constructor(cell: Cell, value: Date) {
    this.model = {
      address: cell.address,
      type: Cell.Types.Date,
      value
    };
  }

  get value(): Date {
    return this.model.value;
  }

  set value(value: Date) {
    this.model.value = value;
  }

  get type(): ValueType {
    return Cell.Types.Date;
  }

  get effectiveType(): ValueType {
    return Cell.Types.Date;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  toCsvString(): string {
    return this.model.value.toISOString();
  }

  release(): void {}

  toString(): string {
    return this.model.value.toString();
  }
}

class HyperlinkValue {
  declare public model: HyperlinkValueModel;

  constructor(cell: Cell, value?: CellHyperlinkValueInput | CellFormulaHyperlinkValue) {
    this.model = {
      address: cell.address,
      type: Cell.Types.Hyperlink
    };
    if (value) {
      // Formula + hyperlink: surface as a Hyperlink whose display is the
      // formula's evaluated result, but persist the formula on the model so
      // it round-trips on write. The cell value getter ignores the formula
      // fields, keeping the public Hyperlink shape clean.
      if ("formula" in value && typeof value.formula === "string") {
        const fh = value as CellFormulaHyperlinkValue;
        const display = fh.result === undefined || fh.result === null ? "" : String(fh.result);
        this.model.text = display;
        this.model.hyperlink = fh.hyperlink ?? "";
        if (fh.tooltip !== undefined) {
          this.model.tooltip = fh.tooltip;
        }
        // Internal-only fields preserved for write-time re-emission.
        (this.model as HyperlinkValueModel & { formula?: string; result?: FormulaResult }).formula =
          fh.formula;
        if (fh.result !== undefined) {
          (this.model as HyperlinkValueModel & { result?: FormulaResult }).result = fh.result;
        }
        return;
      }
      const normalized = normalizeHyperlinkValue(value as CellHyperlinkValueInput);
      this.model.text = normalized.text;
      this.model.hyperlink = normalized.hyperlink;
      if (normalized.richText) {
        this.model.richText = normalized.richText;
      }
      if (normalized.tooltip !== undefined) {
        this.model.tooltip = normalized.tooltip;
      }
    }
  }

  get value(): CellHyperlinkValue {
    const out: CellHyperlinkValue = {
      text: this.model.text ?? "",
      hyperlink: this.model.hyperlink ?? ""
    };
    if (this.model.richText && this.model.richText.length > 0) {
      out.richText = this.model.richText;
    }
    if (this.model.tooltip !== undefined) {
      out.tooltip = this.model.tooltip;
    }
    return out;
  }

  set value(value: CellHyperlinkValueInput) {
    const normalized = normalizeHyperlinkValue(value);
    this.model.text = normalized.text;
    this.model.hyperlink = normalized.hyperlink;
    if (normalized.richText) {
      this.model.richText = normalized.richText;
    } else {
      delete this.model.richText;
    }
    if (normalized.tooltip !== undefined) {
      this.model.tooltip = normalized.tooltip;
    } else {
      delete this.model.tooltip;
    }
  }

  get text(): string | undefined {
    return this.model.text;
  }

  set text(value: string | undefined) {
    // Setting text while richText is present would break the invariant
    // (text must equal flattenRichText(richText)). Dropping richText is the
    // only safe resolution.
    if (this.model.richText) {
      delete this.model.richText;
    }
    this.model.text = value;
  }

  get richText(): RichText[] | undefined {
    return this.model.richText;
  }

  set richText(value: RichText[] | undefined) {
    if (Array.isArray(value) && value.length > 0) {
      this.model.richText = value;
      this.model.text = flattenRichText(value);
    } else {
      delete this.model.richText;
      // leave model.text untouched — caller may still want the plain text
    }
  }

  get hyperlink(): string | undefined {
    return this.model.hyperlink;
  }

  set hyperlink(value: string | undefined) {
    this.model.hyperlink = value;
  }

  get type(): ValueType {
    return Cell.Types.Hyperlink;
  }

  get effectiveType(): ValueType {
    return Cell.Types.Hyperlink;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  toCsvString(): string {
    return this.model.hyperlink ?? "";
  }

  release(): void {}

  toString(): string {
    return this.model.text ?? "";
  }
}

class MergeValue {
  declare public model: MergeValueModel;
  declare private _master: Cell;

  constructor(cell: Cell, master?: Cell) {
    this.model = {
      address: cell.address,
      type: Cell.Types.Merge,
      master: master ? master.address : undefined
    };
    this._master = master as Cell;
    if (master) {
      master.addMergeRef();
    }
  }

  get value(): CellValueType {
    return this._master.value;
  }

  set value(value: CellValueInputType | Cell) {
    if (value instanceof Cell) {
      if (this._master) {
        this._master.releaseMergeRef();
      }
      value.addMergeRef();
      this._master = value;
    } else {
      this._master.value = value;
    }
  }

  isMergedTo(master: Cell): boolean {
    return master === this._master;
  }

  get master(): Cell {
    return this._master;
  }

  get type(): ValueType {
    return Cell.Types.Merge;
  }

  get effectiveType(): ValueType {
    return this._master.effectiveType;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  toCsvString(): string {
    return "";
  }

  release(): void {
    this._master.releaseMergeRef();
  }

  toString(): string {
    return this.value != null ? this.value.toString() : "";
  }
}

class FormulaValue {
  declare public cell: Cell;
  declare public model: FormulaValueModel;
  declare private _translatedFormula?: string;

  constructor(cell: Cell, value?: FormulaValueData) {
    this.cell = cell;

    this.model = {
      address: cell.address,
      type: Cell.Types.Formula,
      shareType: value ? value.shareType : undefined,
      ref: value ? value.ref : undefined,
      formula: value ? value.formula : undefined,
      sharedFormula: value ? value.sharedFormula : undefined,
      result: value ? value.result : undefined,
      isDynamicArray: value ? value.isDynamicArray : undefined
    };
  }

  private _copyModel(model: FormulaValueModel): FormulaValueData {
    const copy: FormulaValueData = {};
    if (model.formula) {
      copy.formula = model.formula;
    }
    if (model.result !== undefined) {
      copy.result = model.result;
    }
    if (model.ref) {
      copy.ref = model.ref;
    }
    if (model.shareType) {
      copy.shareType = model.shareType;
    }
    if (model.sharedFormula) {
      copy.sharedFormula = model.sharedFormula;
    }
    if (model.isDynamicArray) {
      copy.isDynamicArray = model.isDynamicArray;
    }
    return copy;
  }

  get value(): FormulaValueData {
    return this._copyModel(this.model);
  }

  set value(value: FormulaValueData) {
    if (value.formula) {
      this.model.formula = value.formula;
    }
    if (value.result !== undefined) {
      this.model.result = value.result;
    }
    if (value.ref) {
      this.model.ref = value.ref;
    }
    if (value.shareType) {
      this.model.shareType = value.shareType;
    }
    if (value.sharedFormula) {
      this.model.sharedFormula = value.sharedFormula;
    }
    if (value.isDynamicArray !== undefined) {
      this.model.isDynamicArray = value.isDynamicArray;
    }
  }

  validate(value: CellValueType): void {
    switch (Value.getType(value)) {
      case Cell.Types.Null:
      case Cell.Types.String:
      case Cell.Types.Number:
      case Cell.Types.Date:
        break;
      case Cell.Types.Hyperlink:
      case Cell.Types.Formula:
      default:
        throw new InvalidValueTypeError(
          String(Value.getType(value)),
          "Cannot process that type of result value"
        );
    }
  }

  get dependencies(): { ranges: string[] | null; cells: string[] | null } {
    // find all the ranges and cells mentioned in the formula
    const ranges = this.formula.match(/([a-zA-Z0-9]+!)?[A-Z]{1,3}\d{1,4}:[A-Z]{1,3}\d{1,4}/g);
    const cells = this.formula
      .replace(/([a-zA-Z0-9]+!)?[A-Z]{1,3}\d{1,4}:[A-Z]{1,3}\d{1,4}/g, "")
      .match(/([a-zA-Z0-9]+!)?[A-Z]{1,3}\d{1,4}/g);
    return {
      ranges,
      cells
    };
  }

  get formula(): string {
    return this.model.formula || this._getTranslatedFormula() || "";
  }

  set formula(value: string) {
    this.model.formula = value;
  }

  get formulaType(): FormulaType {
    if (this.model.formula) {
      return Enums.FormulaType.Master;
    }
    if (this.model.sharedFormula) {
      return Enums.FormulaType.Shared;
    }
    return Enums.FormulaType.None;
  }

  get result(): FormulaResult | undefined {
    return this.model.result;
  }

  set result(value: FormulaResult | undefined) {
    this.model.result = value;
  }

  get type(): ValueType {
    return Cell.Types.Formula;
  }

  get effectiveType(): ValueType {
    const v = this.model.result;
    if (v === null || v === undefined) {
      return Enums.ValueType.Null;
    }
    if (typeof v === "string") {
      return Enums.ValueType.String;
    }
    if (typeof v === "number") {
      return Enums.ValueType.Number;
    }
    if (v instanceof Date) {
      return Enums.ValueType.Date;
    }
    if (typeof v === "object" && "error" in v) {
      return Enums.ValueType.Error;
    }

    return Enums.ValueType.Null;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  _getTranslatedFormula(): string | undefined {
    if (!this._translatedFormula && this.model.sharedFormula) {
      const { worksheet } = this.cell;
      const master = worksheet.findCell(this.model.sharedFormula);
      this._translatedFormula =
        master && master.formula
          ? slideFormula(master.formula, master.address, this.model.address)
          : undefined;
    }
    return this._translatedFormula;
  }

  toCsvString(): string {
    return `${this.model.result ?? ""}`;
  }

  release(): void {}

  toString(): string {
    return this.model.result ? this.model.result.toString() : "";
  }
}

class SharedStringValue {
  declare public model: SharedStringValueModel;

  constructor(cell: Cell, value: number) {
    this.model = {
      address: cell.address,
      type: Cell.Types.SharedString,
      value
    };
  }

  get value(): number {
    return this.model.value;
  }

  set value(value: number) {
    this.model.value = value;
  }

  get type(): ValueType {
    return Cell.Types.SharedString;
  }

  get effectiveType(): ValueType {
    return Cell.Types.SharedString;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  toCsvString(): string {
    return this.model.value.toString();
  }

  release(): void {}

  toString(): string {
    return this.model.value.toString();
  }
}

class BooleanValue {
  declare public model: BooleanValueModel;

  constructor(cell: Cell, value: boolean) {
    this.model = {
      address: cell.address,
      type: Cell.Types.Boolean,
      value
    };
  }

  get value(): boolean {
    return this.model.value;
  }

  set value(value: boolean) {
    this.model.value = value;
  }

  get type(): ValueType {
    return Cell.Types.Boolean;
  }

  get effectiveType(): ValueType {
    return Cell.Types.Boolean;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  toCsvString(): number {
    return this.model.value ? 1 : 0;
  }

  release(): void {}

  toString(): string {
    return this.model.value.toString();
  }
}

interface CheckboxValueModel extends CellModel {
  type: typeof Cell.Types.Checkbox;
  value: boolean;
}

class CheckboxValue {
  declare public model: CheckboxValueModel;

  constructor(cell: Cell, value: CellCheckboxValue) {
    this.model = {
      address: cell.address,
      type: Cell.Types.Checkbox,
      value: value.checkbox
    };
  }

  get value(): CellCheckboxValue {
    return { checkbox: this.model.value };
  }

  set value(value: CellCheckboxValue) {
    this.model.value = value.checkbox;
  }

  get type(): ValueType {
    return Cell.Types.Checkbox;
  }

  get effectiveType(): ValueType {
    return Cell.Types.Boolean;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  toCsvString(): number {
    return this.model.value ? 1 : 0;
  }

  release(): void {}

  toString(): string {
    return this.model.value.toString();
  }
}

class ErrorValue {
  declare public model: ErrorValueModel;

  constructor(cell: Cell, value: CellErrorValue) {
    this.model = {
      address: cell.address,
      type: Cell.Types.Error,
      value
    };
  }

  get value(): CellErrorValue {
    return this.model.value;
  }

  set value(value: CellErrorValue) {
    this.model.value = value;
  }

  get type(): ValueType {
    return Cell.Types.Error;
  }

  get effectiveType(): ValueType {
    return Cell.Types.Error;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  toCsvString(): string {
    return this.toString();
  }

  release(): void {}

  toString(): string {
    return this.model.value.error.toString();
  }
}

class JSONValue {
  declare public model: JSONValueModel;

  constructor(cell: Cell, value: unknown) {
    this.model = {
      address: cell.address,
      type: Cell.Types.String,
      value: JSON.stringify(value),
      rawValue: value
    };
  }

  get value(): unknown {
    return this.model.rawValue;
  }

  set value(value: unknown) {
    this.model.rawValue = value;
    this.model.value = JSON.stringify(value);
  }

  get type(): ValueType {
    return Cell.Types.String;
  }

  get effectiveType(): ValueType {
    return Cell.Types.String;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  toCsvString(): string {
    return this.model.value;
  }

  release(): void {}

  toString(): string {
    return this.model.value;
  }
}

// Value is a place to hold common static Value type functions
const Value = {
  getType(value: CellValueInputType): number {
    if (value === null || value === undefined) {
      return Cell.Types.Null;
    }
    if (typeof value === "string") {
      return Cell.Types.String;
    }
    if (typeof value === "number") {
      return Cell.Types.Number;
    }
    if (typeof value === "boolean") {
      return Cell.Types.Boolean;
    }
    if (value instanceof Date) {
      return Cell.Types.Date;
    }
    if (typeof value === "object") {
      if ("checkbox" in value && typeof value.checkbox === "boolean") {
        return Cell.Types.Checkbox;
      }
      // Hyperlink detection: requires `hyperlink` and either non-empty `text`
      // or a non-empty `richText` array, OR a formula (formula+hyperlink is
      // surfaced as a Hyperlink with the formula's result as display).
      // Checked before RichText/Formula so combined payloads route correctly.
      if ("hyperlink" in value && typeof value.hyperlink === "string" && value.hyperlink) {
        const hasText = "text" in value && typeof value.text === "string" && value.text.length > 0;
        const hasRichText =
          "richText" in value && Array.isArray(value.richText) && value.richText.length > 0;
        const hasFormula =
          "formula" in value && typeof (value as { formula?: unknown }).formula === "string";
        if (hasText || hasRichText || hasFormula) {
          return Cell.Types.Hyperlink;
        }
      }
      if (
        ("formula" in value && value.formula) ||
        ("sharedFormula" in value && value.sharedFormula)
      ) {
        return Cell.Types.Formula;
      }
      // RichText only when the runs array is non-empty. An empty `richText: []`
      // carries no content and falls through to JSON rather than producing a
      // RichText cell with no runs.
      if ("richText" in value && Array.isArray(value.richText) && value.richText.length > 0) {
        return Cell.Types.RichText;
      }
      if ("sharedString" in value && value.sharedString) {
        return Cell.Types.SharedString;
      }
      if ("error" in value && value.error) {
        return Cell.Types.Error;
      }
    }
    return Cell.Types.JSON;
  },

  // map valueType to constructor
  types: [
    { t: Cell.Types.Null, f: NullValue },
    { t: Cell.Types.Number, f: NumberValue },
    { t: Cell.Types.String, f: StringValue },
    { t: Cell.Types.Date, f: DateValue },
    { t: Cell.Types.Hyperlink, f: HyperlinkValue },
    { t: Cell.Types.Formula, f: FormulaValue },
    { t: Cell.Types.Merge, f: MergeValue },
    { t: Cell.Types.JSON, f: JSONValue },
    { t: Cell.Types.SharedString, f: SharedStringValue },
    { t: Cell.Types.RichText, f: RichTextValue },
    { t: Cell.Types.Boolean, f: BooleanValue },
    { t: Cell.Types.Error, f: ErrorValue },
    { t: Cell.Types.Checkbox, f: CheckboxValue }
  ].reduce(
    (
      p: (
        | typeof NullValue
        | typeof NumberValue
        | typeof StringValue
        | typeof DateValue
        | typeof HyperlinkValue
        | typeof FormulaValue
        | typeof MergeValue
        | typeof JSONValue
        | typeof SharedStringValue
        | typeof RichTextValue
        | typeof BooleanValue
        | typeof ErrorValue
        | typeof CheckboxValue
      )[],
      t
    ) => {
      p[t.t] = t.f;
      return p;
    },
    []
  ),

  create(type: number, cell: Cell, value?: CellValueInputType | Cell): ICellValue {
    const T = this.types[type];
    if (!T) {
      throw new InvalidValueTypeError(String(type), "Could not create Value");
    }
    return new T(cell, value as never) as unknown as ICellValue;
  }
};

export { Cell };
