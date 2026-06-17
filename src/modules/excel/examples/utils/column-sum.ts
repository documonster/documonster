import { Stream } from "@excel/index";

export class ColumnSum {
  private columns: number[];
  private sums: number[];
  public count: number = 0;

  constructor(columns: number[]) {
    this.columns = columns;
    this.sums = [];
    for (const column of this.columns) {
      this.sums[column] = 0;
    }
  }

  add(row: Stream.RowHandle): void {
    for (const column of this.columns) {
      this.sums[column] += Stream.getCellValue(Stream.rowCell(row, column)) as number;
    }
    this.count++;
  }

  toString(): string {
    return this.sums.join(", ");
  }

  toAverages(): string {
    return this.sums.map(value => (value ? value / this.count : value)).join(", ");
  }
}
