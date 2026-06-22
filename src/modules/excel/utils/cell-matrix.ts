/* oxlint-disable typescript/no-redundant-type-constituents -- Cell is intentionally `CellAddress & any` */
import { colCache } from "@excel/utils/col-cache";

// Safe deep clone that filters out prototype pollution keys
function safeDeepClone<T>(obj: T): T {
  if (obj === null || typeof obj !== "object") {
    return obj;
  }
  if (Array.isArray(obj)) {
    return obj.map(item => safeDeepClone(item)) as T;
  }
  const result: Record<string, unknown> = {};
  for (const key of Object.keys(obj)) {
    if (key !== "__proto__" && key !== "constructor" && key !== "prototype") {
      result[key] = safeDeepClone((obj as Record<string, unknown>)[key]);
    }
  }
  return result as T;
}

interface CellAddress {
  sheetName?: string;
  address: string;
  row: number;
  col: number;
  top?: number;
  bottom?: number;
  left?: number;
  right?: number;
}

type Cell = CellAddress & any;
type Row = Cell[];
type Sheet = Row[];

class CellMatrix {
  template: any;
  sheets: Map<string, Sheet>;

  constructor(template?: any) {
    this.template = template;
    this.sheets = new Map();
  }

  addCell(addressStr: string): void {
    this.addCellEx(colCache.decodeEx(addressStr) as CellAddress);
  }

  getCell(addressStr: string): Cell {
    return this.findCellEx(colCache.decodeEx(addressStr) as CellAddress, true);
  }

  findCell(addressStr: string): Cell | undefined {
    return this.findCellEx(colCache.decodeEx(addressStr) as CellAddress, false);
  }

  findCellAt(sheetName: string, rowNumber: number, colNumber: number): Cell | undefined {
    const sheet = this.sheets.get(sheetName);
    const row = sheet && sheet[rowNumber];
    return row && row[colNumber];
  }

  addCellEx(address: CellAddress): void {
    if (address.top !== undefined) {
      for (let row = address.top; row <= address.bottom!; row++) {
        for (let col = address.left!; col <= address.right!; col++) {
          this.getCellAt(address.sheetName!, row, col);
        }
      }
    } else {
      this.findCellEx(address, true);
    }
  }

  getCellEx(address: CellAddress): Cell {
    return this.findCellEx(address, true);
  }

  findCellEx(address: CellAddress, create: boolean): Cell | undefined {
    const sheet = this.findSheet(address, create);
    const row = this.findSheetRow(sheet, address, create);
    return this.findRowCell(row, address, create);
  }

  getCellAt(sheetName: string, rowNumber: number, colNumber: number): Cell {
    let sheet = this.sheets.get(sheetName);
    if (!sheet) {
      sheet = [];
      this.sheets.set(sheetName, sheet);
    }
    // >>> 0 coerces to uint32, preventing "__proto__" string injection
    const safeRow = rowNumber >>> 0;
    const safeCol = colNumber >>> 0;
    const row = sheet[safeRow] || (sheet[safeRow] = []);
    const cell =
      row[safeCol] ||
      (row[safeCol] = {
        sheetName,
        address: colCache.n2l(colNumber) + rowNumber,
        row: rowNumber,
        col: colNumber
      });
    return cell;
  }

  removeCellEx(address: CellAddress): void {
    const sheet = this.findSheet(address, false);
    if (!sheet) {
      return;
    }
    const row = this.findSheetRow(sheet, address, false);
    if (!row) {
      return;
    }
    row[address.col >>> 0] = undefined!;
  }

  forEachInSheet(
    sheetName: string,
    callback: (cell: Cell, rowNumber: number, colNumber: number) => void
  ): void {
    const sheet = this.sheets.get(sheetName);
    if (sheet) {
      sheet.forEach((row, rowNumber) => {
        if (row) {
          row.forEach((cell, colNumber) => {
            if (cell) {
              callback(cell, rowNumber, colNumber);
            }
          });
        }
      });
    }
  }

  forEach(callback: (cell: Cell) => void): void {
    for (const sheetName of this.sheets.keys()) {
      this.forEachInSheet(sheetName, callback);
    }
  }

  map<T>(callback: (cell: Cell) => T): T[] {
    const results: T[] = [];
    this.forEach(cell => {
      results.push(callback(cell));
    });
    return results;
  }

  findSheet(address: CellAddress, create: boolean): Sheet | undefined {
    const name = address.sheetName!;
    if (this.sheets.has(name)) {
      return this.sheets.get(name);
    }
    if (create) {
      const sheet: Sheet = [];
      this.sheets.set(name, sheet);
      return sheet;
    }
    return undefined;
  }

  findSheetRow(sheet: Sheet | undefined, address: CellAddress, create: boolean): Row | undefined {
    const safeRow = address.row >>> 0;
    if (sheet && sheet[safeRow]) {
      return sheet[safeRow];
    }
    if (create) {
      return (sheet![safeRow] = []);
    }
    return undefined;
  }

  findRowCell(row: Row | undefined, address: CellAddress, create: boolean): Cell | undefined {
    const safeCol = address.col >>> 0;
    if (row && row[safeCol]) {
      return row[safeCol];
    }
    if (create) {
      return (row![safeCol] = this.template
        ? { ...address, ...safeDeepClone(this.template) }
        : address);
    }
    return undefined;
  }

  spliceRows(sheetName: string, start: number, numDelete: number, numInsert: number): void {
    const sheet = this.sheets.get(sheetName);
    if (sheet) {
      const inserts: Row[] = [];
      for (let i = 0; i < numInsert; i++) {
        inserts.push([]);
      }
      sheet.splice(start, numDelete, ...inserts);
    }
  }

  spliceColumns(sheetName: string, start: number, numDelete: number, numInsert: number): void {
    const sheet = this.sheets.get(sheetName);
    if (sheet) {
      const inserts: (Cell | null)[] = [];
      for (let i = 0; i < numInsert; i++) {
        inserts.push(null);
      }
      sheet.forEach((row: Row) => {
        row.splice(start, numDelete, ...inserts);
      });
    }
  }
}

export { CellMatrix };
