import { ExcelError } from "@excel/errors";
import type { Address } from "@excel/types";
import { colCache } from "@excel/utils/col-cache";

interface RangeModel {
  top: number;
  left: number;
  bottom: number;
  right: number;
  sheetName?: string;
}

interface RowDimensions {
  min: number;
  max: number;
}

interface RowWithDimensions {
  number: number;
  dimensions?: RowDimensions;
}

// Input types for Range constructor and decode
export type RangeInput = Range | RangeModel | string | number | RangeInput[];

// used by worksheet to calculate sheet dimensions
class Range {
  model: RangeModel = {
    top: 0,
    left: 0,
    bottom: 0,
    right: 0
  };

  // Constructor overloads
  constructor();
  constructor(range: Range);
  constructor(model: RangeModel);
  constructor(rangeString: string);
  constructor(args: RangeInput[]);
  constructor(tl: string, br: string);
  constructor(tl: string, br: string, sheetName: string);
  constructor(top: number, left: number, bottom: number, right: number);
  constructor(top: number, left: number, bottom: number, right: number, sheetName: string);
  constructor(...args: RangeInput[]) {
    this.decode(args);
  }

  // setTLBR overloads
  setTLBR(tl: string, br: string, sheetName?: string): void;
  setTLBR(top: number, left: number, bottom: number, right: number, sheetName?: string): void;
  setTLBR(
    t: number | string,
    l: number | string,
    b?: number | string,
    r?: number,
    s?: string
  ): void {
    if (typeof t === "string" && typeof l === "string") {
      // setTLBR(tl, br, s) - t and l are address strings
      const tl = colCache.decodeAddress(t);
      const br = colCache.decodeAddress(l);
      this.model = {
        top: Math.min(tl.row, br.row),
        left: Math.min(tl.col, br.col),
        bottom: Math.max(tl.row, br.row),
        right: Math.max(tl.col, br.col),
        sheetName: typeof b === "string" ? b : undefined
      };
    } else if (
      typeof t === "number" &&
      typeof l === "number" &&
      typeof b === "number" &&
      typeof r === "number"
    ) {
      // setTLBR(t, l, b, r, s) - all numbers
      this.model = {
        top: Math.min(t, b),
        left: Math.min(l, r),
        bottom: Math.max(t, b),
        right: Math.max(l, r),
        sheetName: s
      };
    }
  }

  private decode(argv: RangeInput[]): void {
    switch (argv.length) {
      case 5: // [t,l,b,r,s]
        if (
          typeof argv[0] === "number" &&
          typeof argv[1] === "number" &&
          typeof argv[2] === "number" &&
          typeof argv[3] === "number" &&
          typeof argv[4] === "string"
        ) {
          this.setTLBR(argv[0], argv[1], argv[2], argv[3], argv[4]);
        }
        break;
      case 4: // [t,l,b,r]
        if (
          typeof argv[0] === "number" &&
          typeof argv[1] === "number" &&
          typeof argv[2] === "number" &&
          typeof argv[3] === "number"
        ) {
          this.setTLBR(argv[0], argv[1], argv[2], argv[3]);
        }
        break;

      case 3: // [tl,br,s]
        if (
          typeof argv[0] === "string" &&
          typeof argv[1] === "string" &&
          typeof argv[2] === "string"
        ) {
          this.setTLBR(argv[0], argv[1], argv[2]);
        }
        break;
      case 2: // [tl,br]
        if (typeof argv[0] === "string" && typeof argv[1] === "string") {
          this.setTLBR(argv[0], argv[1]);
        }
        break;

      case 1: {
        const value = argv[0];
        if (value instanceof Range) {
          // copy constructor
          this.model = {
            top: value.model.top,
            left: value.model.left,
            bottom: value.model.bottom,
            right: value.model.right,
            sheetName: value.sheetName
          };
        } else if (Array.isArray(value)) {
          // an arguments array
          this.decode(value);
        } else if (
          typeof value === "object" &&
          "top" in value &&
          "left" in value &&
          "bottom" in value &&
          "right" in value
        ) {
          // a model
          this.model = {
            top: value.top,
            left: value.left,
            bottom: value.bottom,
            right: value.right,
            sheetName: value.sheetName
          };
        } else if (typeof value === "string") {
          // [sheetName!]tl:br
          const decoded = colCache.decodeEx(value);
          if ("top" in decoded) {
            // It's a DecodedRange
            this.model = {
              top: decoded.top,
              left: decoded.left,
              bottom: decoded.bottom,
              right: decoded.right,
              sheetName: decoded.sheetName
            };
          } else if ("row" in decoded) {
            // It's an Address
            this.model = {
              top: decoded.row,
              left: decoded.col,
              bottom: decoded.row,
              right: decoded.col,
              sheetName: decoded.sheetName
            };
          }
        }
        break;
      }

      case 0:
        this.model = {
          top: 0,
          left: 0,
          bottom: 0,
          right: 0
        };
        break;

      default:
        throw new ExcelError(`Invalid number of arguments to _getDimensions() - ${argv.length}`);
    }
  }

  get top(): number {
    return this.model.top || 1;
  }

  set top(value: number) {
    this.model.top = value;
  }

  get left(): number {
    return this.model.left || 1;
  }

  set left(value: number) {
    this.model.left = value;
  }

  get bottom(): number {
    return this.model.bottom || 1;
  }

  set bottom(value: number) {
    this.model.bottom = value;
  }

  get right(): number {
    return this.model.right || 1;
  }

  set right(value: number) {
    this.model.right = value;
  }

  get sheetName(): string | undefined {
    return this.model.sheetName;
  }

  set sheetName(value: string | undefined) {
    this.model.sheetName = value;
  }

  private get _serialisedSheetName(): string {
    const { sheetName } = this.model;
    if (sheetName) {
      // Plain ASCII identifiers (and the empty string) need no quoting.
      if (/^[a-zA-Z0-9]*$/.test(sheetName)) {
        return `${sheetName}!`;
      }
      // Per OOXML / Excel formula syntax, sheet names containing characters
      // outside the bareword set are wrapped in single quotes, and any
      // embedded apostrophes are doubled (e.g. `O'Brien` -> `'O''Brien'`).
      return `'${sheetName.replace(/'/g, "''")}'!`;
    }
    return "";
  }

  expand(top: number, left: number, bottom: number, right: number): void {
    if (!this.model.top || top < this.top) {
      this.top = top;
    }
    if (!this.model.left || left < this.left) {
      this.left = left;
    }
    if (!this.model.bottom || bottom > this.bottom) {
      this.bottom = bottom;
    }
    if (!this.model.right || right > this.right) {
      this.right = right;
    }
  }

  expandRow(row: RowWithDimensions | null | undefined): void {
    if (row) {
      const { dimensions, number } = row;
      if (dimensions) {
        this.expand(number, dimensions.min, number, dimensions.max);
      }
    }
  }

  expandToAddress(addressStr: string): void {
    const address = colCache.decodeEx(addressStr);
    if ("row" in address && "col" in address) {
      this.expand(address.row, address.col, address.row, address.col);
    }
  }

  get tl(): string {
    return colCache.n2l(this.left) + this.top;
  }

  get $t$l(): string {
    return `$${colCache.n2l(this.left)}$${this.top}`;
  }

  get br(): string {
    return colCache.n2l(this.right) + this.bottom;
  }

  get $b$r(): string {
    return `$${colCache.n2l(this.right)}$${this.bottom}`;
  }

  get range(): string {
    return `${this._serialisedSheetName + this.tl}:${this.br}`;
  }

  get $range(): string {
    return `${this._serialisedSheetName + this.$t$l}:${this.$b$r}`;
  }

  get shortRange(): string {
    return this.count > 1 ? this.range : this._serialisedSheetName + this.tl;
  }

  get $shortRange(): string {
    return this.count > 1 ? this.$range : this._serialisedSheetName + this.$t$l;
  }

  get count(): number {
    return (1 + this.bottom - this.top) * (1 + this.right - this.left);
  }

  toString(): string {
    return this.range;
  }

  intersects(other: Range): boolean {
    if (other.sheetName && this.sheetName && other.sheetName !== this.sheetName) {
      return false;
    }
    if (other.bottom < this.top) {
      return false;
    }
    if (other.top > this.bottom) {
      return false;
    }
    if (other.right < this.left) {
      return false;
    }
    if (other.left > this.right) {
      return false;
    }
    return true;
  }

  contains(addressStr: string): boolean {
    const address = colCache.decodeEx(addressStr);
    if ("row" in address && "col" in address) {
      return this.containsEx(address);
    }
    return false;
  }

  containsEx(address: Address): boolean {
    if (address.sheetName && this.sheetName && address.sheetName !== this.sheetName) {
      return false;
    }
    return (
      address.row >= this.top &&
      address.row <= this.bottom &&
      address.col >= this.left &&
      address.col <= this.right
    );
  }

  forEachAddress(cb: (address: string, row: number, col: number) => void): void {
    for (let col = this.left; col <= this.right; col++) {
      for (let row = this.top; row <= this.bottom; row++) {
        cb(colCache.encodeAddress(row, col), row, col);
      }
    }
  }
}

export { Range };
export { Range as Dimensions };
