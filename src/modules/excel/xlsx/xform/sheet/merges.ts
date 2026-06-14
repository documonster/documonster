import { Enums } from "@excel/enums";
import {
  type RangeData,
  rangeCreate,
  rangeExpandToAddress,
  rangeRange,
  rangeTl
} from "@excel/range";
import { colCache } from "@excel/utils/col-cache";

interface MergeData {
  address: string;
  master: string;
}

class Merges {
  declare private merges: { [key: string]: RangeData };
  declare private hash?: { [key: string]: RangeData };

  constructor() {
    // optional mergeCells is array of ranges (like the xml)
    this.merges = {};
  }

  add(merge: MergeData): void {
    // merge is {address, master}
    if (this.merges[merge.master]) {
      rangeExpandToAddress(this.merges[merge.master], merge.address);
    } else {
      const range = `${merge.master}:${merge.address}`;
      this.merges[merge.master] = rangeCreate(range);
    }
  }

  get mergeCells(): string[] {
    return Object.values(this.merges).map((merge: RangeData) => rangeRange(merge));
  }

  reconcile(mergeCells: string[], rows: any[]): void {
    // reconcile merge list with merge cells
    mergeCells.forEach((merge: string) => {
      const dimensions: any = colCache.decode(merge);
      for (let i = dimensions.top; i <= dimensions.bottom; i++) {
        const row = rows[i - 1];
        for (let j = dimensions.left; j <= dimensions.right; j++) {
          const cell = row.cells[j - 1];
          if (!cell) {
            // nulls are not included in document - so if master cell has no value - add a null one here
            row.cells[j] = {
              type: Enums.ValueType.Null,
              address: colCache.encodeAddress(i, j)
            };
          } else if (cell.type === Enums.ValueType.Merge) {
            cell.master = dimensions.tl;
          }
        }
      }
    });
  }

  getMasterAddress(address: string): string | undefined {
    // if address has been merged, return its master's address. Assumes reconcile has been called
    const range = this.hash![address];
    return range && rangeTl(range);
  }
}

export { Merges };
