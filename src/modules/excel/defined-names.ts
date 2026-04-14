import { Range } from "@excel/range";
import type { Address } from "@excel/types";
import { CellMatrix } from "@excel/utils/cell-matrix";
import { colCache, type DecodedRange } from "@excel/utils/col-cache";

const rangeRegexp = /[$](\w+)[$](\d+)(:[$](\w+)[$](\d+))?/;

// Cell type for defined names - extends Address with mark for exploration algorithm
interface DefinedNameCell {
  sheetName?: string;
  address: string;
  row: number;
  col: number;
  mark?: boolean;
}

// Location can be a single cell address or a range
type CellLocation = Address | DecodedRange;

interface DefinedNameModel {
  name: string;
  ranges: string[];
}

class DefinedNames {
  matrixMap: Record<string, CellMatrix>;

  constructor() {
    this.matrixMap = {};
  }

  getMatrix(name: string): CellMatrix {
    const matrix = this.matrixMap[name] || (this.matrixMap[name] = new CellMatrix());
    return matrix;
  }

  // add a name to a cell. locStr in the form SheetName!$col$row or SheetName!$c1$r1:$c2:$r2
  add(locStr: string, name: string): void {
    const location = colCache.decodeEx(locStr);
    if ("error" in location) {
      return; // Invalid reference, skip
    }
    this.addEx(location, name);
  }

  addEx(location: CellLocation, name: string): void {
    const matrix = this.getMatrix(name);
    if ("top" in location) {
      // It's a range (DecodedRange has top/left/bottom/right from Location)
      for (let col = location.left; col <= location.right; col++) {
        for (let row = location.top; row <= location.bottom; row++) {
          const address = {
            sheetName: location.sheetName,
            address: colCache.n2l(col) + row,
            row,
            col
          };

          matrix.addCellEx(address);
        }
      }
    } else {
      // It's a single cell address
      matrix.addCellEx(location);
    }
  }

  remove(locStr: string, name: string): void {
    const location = colCache.decodeEx(locStr);
    if ("error" in location) {
      return; // Invalid reference, skip
    }
    this.removeEx(location, name);
  }

  removeEx(location: CellLocation, name: string): void {
    const matrix = this.getMatrix(name);
    if ("top" in location) {
      // Range - remove each cell
      for (let col = location.left; col <= location.right; col++) {
        for (let row = location.top; row <= location.bottom; row++) {
          matrix.removeCellEx({
            sheetName: location.sheetName,
            address: colCache.n2l(col) + row,
            row,
            col
          });
        }
      }
    } else {
      matrix.removeCellEx(location);
    }
  }

  removeAllNames(location: CellLocation): void {
    Object.values(this.matrixMap).forEach((matrix: CellMatrix) => {
      if ("top" in location) {
        // Range - remove each cell
        for (let col = location.left; col <= location.right; col++) {
          for (let row = location.top; row <= location.bottom; row++) {
            matrix.removeCellEx({
              sheetName: location.sheetName,
              address: colCache.n2l(col) + row,
              row,
              col
            });
          }
        }
      } else {
        matrix.removeCellEx(location);
      }
    });
  }

  forEach(callback: (name: string, cell: DefinedNameCell) => void): void {
    Object.entries(this.matrixMap).forEach(([name, matrix]) => {
      matrix.forEach((cell: DefinedNameCell) => {
        callback(name, cell);
      });
    });
  }

  // get all the names of a cell
  getNames(addressStr: string): string[] {
    const location = colCache.decodeEx(addressStr);
    if ("error" in location || "top" in location) {
      return []; // Invalid reference or range not supported
    }
    return this.getNamesEx(location);
  }

  getNamesEx(address: Address): string[] {
    return Object.entries(this.matrixMap)
      .map(([name, matrix]) => matrix.findCellEx(address, false) && name)
      .filter((name): name is string => Boolean(name));
  }

  _explore(matrix: CellMatrix, cell: DefinedNameCell): Range {
    cell.mark = false;
    const sheetName = cell.sheetName!; // Always set for cells in defined names

    const range = new Range(cell.row, cell.col, cell.row, cell.col, sheetName);
    let x: number;
    let y: number;

    // Helper to get cell with proper type
    const getCell = (row: number, col: number): DefinedNameCell | undefined => {
      return matrix.findCellAt(sheetName, row, col) as DefinedNameCell | undefined;
    };

    // grow vertical - only one col to worry about
    function vGrow(yy: number, edge: "top" | "bottom"): boolean {
      const c = getCell(yy, cell.col);
      if (!c || !c.mark) {
        return false;
      }
      range[edge] = yy;
      c.mark = false;
      return true;
    }
    for (y = cell.row - 1; vGrow(y, "top"); y--) {
      /* advance */
    }
    for (y = cell.row + 1; vGrow(y, "bottom"); y++) {
      /* advance */
    }

    // grow horizontal - ensure all rows can grow
    function hGrow(xx: number, edge: "left" | "right"): boolean {
      const cells: DefinedNameCell[] = [];
      for (y = range.top; y <= range.bottom; y++) {
        const c = getCell(y, xx);
        if (c && c.mark) {
          cells.push(c);
        } else {
          return false;
        }
      }
      range[edge] = xx;
      for (let i = 0; i < cells.length; i++) {
        cells[i].mark = false;
      }
      return true;
    }
    for (x = cell.col - 1; hGrow(x, "left"); x--) {
      /* advance */
    }
    for (x = cell.col + 1; hGrow(x, "right"); x++) {
      /* advance */
    }

    return range;
  }

  getRanges(name: string, matrix?: CellMatrix): DefinedNameModel {
    matrix = matrix || this.matrixMap[name];

    if (!matrix) {
      return { name, ranges: [] };
    }

    // mark and sweep!
    matrix.forEach((cell: DefinedNameCell) => {
      cell.mark = true;
    });
    const ranges = matrix
      .map((cell: DefinedNameCell) => cell.mark && this._explore(matrix!, cell))
      .filter((range): range is Range => Boolean(range))
      .map((range: Range) => range.$shortRange);

    return {
      name,
      ranges
    };
  }

  normaliseMatrix(matrix: CellMatrix, sheetName: string): void {
    // some of the cells might have shifted on specified sheet
    // need to reassign rows, cols
    matrix.forEachInSheet(
      sheetName,
      (cell: DefinedNameCell | undefined, row: number, col: number) => {
        if (cell) {
          if (cell.row !== row || cell.col !== col) {
            cell.row = row;
            cell.col = col;
            cell.address = colCache.n2l(col) + row;
          }
        }
      }
    );
  }

  spliceRows(sheetName: string, start: number, numDelete: number, numInsert: number): void {
    Object.values(this.matrixMap).forEach((matrix: CellMatrix) => {
      matrix.spliceRows(sheetName, start, numDelete, numInsert);
      this.normaliseMatrix(matrix, sheetName);
    });
  }

  spliceColumns(sheetName: string, start: number, numDelete: number, numInsert: number): void {
    Object.values(this.matrixMap).forEach((matrix: CellMatrix) => {
      matrix.spliceColumns(sheetName, start, numDelete, numInsert);
      this.normaliseMatrix(matrix, sheetName);
    });
  }

  get model(): DefinedNameModel[] {
    // To get names per cell - just iterate over all names finding cells if they exist
    return Object.entries(this.matrixMap)
      .map(([name, matrix]) => this.getRanges(name, matrix))
      .filter((definedName: DefinedNameModel) => definedName.ranges.length);
  }

  set model(value: DefinedNameModel[]) {
    // value is [ { name, ranges }, ... ]
    const matrixMap = (this.matrixMap = {});
    value.forEach(definedName => {
      const matrix = (matrixMap[definedName.name] = new CellMatrix());
      definedName.ranges.forEach(rangeStr => {
        if (rangeRegexp.test(rangeStr.split("!").pop() ?? "")) {
          matrix.addCell(rangeStr);
        }
      });
    });
  }
}

export { DefinedNames, type DefinedNameModel };
