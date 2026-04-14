import { colCache } from "@excel/utils/col-cache";
import { BaseXform } from "@excel/xlsx/xform/base-xform";

interface DefinedNameModel {
  name: string;
  ranges: string[];
  localSheetId?: number;
}

class DefinedNamesXform extends BaseXform {
  declare private _parsedName?: string;
  declare private _parsedLocalSheetId?: string;
  declare private _parsedText: string[];

  constructor() {
    super();
    this._parsedText = [];
  }

  render(xmlStream: any, model: DefinedNameModel): void {
    // <definedNames>
    //   <definedName name="name">name.ranges.join(',')</definedName>
    //   <definedName name="_xlnm.Print_Area" localSheetId="0">name.ranges.join(',')</definedName>
    // </definedNames>
    xmlStream.openNode("definedName", {
      name: model.name,
      localSheetId: model.localSheetId
    });
    xmlStream.writeText(model.ranges.join(","));
    xmlStream.closeNode();
  }

  parseOpen(node: any): boolean {
    switch (node.name) {
      case "definedName":
        this._parsedName = node.attributes.name;
        this._parsedLocalSheetId = node.attributes.localSheetId;
        this._parsedText = [];
        return true;
      default:
        return false;
    }
  }

  parseText(text: string): void {
    this._parsedText.push(text);
  }

  parseClose(): boolean {
    this.model = {
      name: this._parsedName!,
      ranges: extractRanges(this._parsedText.join(""))
    };
    if (this._parsedLocalSheetId !== undefined) {
      this.model.localSheetId = parseInt(this._parsedLocalSheetId, 10);
    }
    return false;
  }
}

// Regex to validate cell range format:
// - Cell: $A$1 or A1
// - Range: $A$1:$B$10 or A1:B10
// - Row range: $1:$2 (for print titles)
// - Column range: $A:$B (for print titles)
const cellRangeRegexp = /^[$]?[A-Za-z]{1,3}[$]?\d+(:[$]?[A-Za-z]{1,3}[$]?\d+)?$/;
const rowRangeRegexp = /^[$]?\d+:[$]?\d+$/;
const colRangeRegexp = /^[$]?[A-Za-z]{1,3}:[$]?[A-Za-z]{1,3}$/;

function isValidRange(range: string): boolean {
  // Skip array constants wrapped in {} - these are not valid cell ranges
  // e.g., {"'Sheet1'!$A$1:$B$10"} or {#N/A,#N/A,FALSE,"text"}
  if (range.startsWith("{") || range.endsWith("}")) {
    return false;
  }

  // Extract the cell reference part (after the sheet name if present)
  const cellRef = range.split("!").pop() ?? "";

  // Must match one of the valid patterns
  if (
    !cellRangeRegexp.test(cellRef) &&
    !rowRangeRegexp.test(cellRef) &&
    !colRangeRegexp.test(cellRef)
  ) {
    return false;
  }

  try {
    const decoded = colCache.decodeEx(range);
    // For cell ranges: row/col or top/bottom/left/right should be valid numbers
    // For row ranges ($1:$2): top/bottom are numbers, left/right are null
    // For column ranges ($A:$B): left/right are numbers, top/bottom are null
    if (
      ("row" in decoded && typeof decoded.row === "number") ||
      ("top" in decoded && typeof decoded.top === "number") ||
      ("left" in decoded && typeof decoded.left === "number")
    ) {
      return true;
    }
    return false;
  } catch {
    return false;
  }
}

function extractRanges(parsedText: string): string[] {
  // Skip if the entire text is wrapped in {} (array constant)
  const trimmed = parsedText.trim();
  if (trimmed.startsWith("{") && trimmed.endsWith("}")) {
    return [];
  }

  const ranges: string[] = [];
  let quotesOpened = false;
  let last = "";
  parsedText.split(",").forEach(item => {
    if (!item) {
      return;
    }
    const quotes = (item.match(/'/g) ?? []).length;

    if (!quotes) {
      if (quotesOpened) {
        last += `${item},`;
      } else if (isValidRange(item)) {
        ranges.push(item);
      }
      return;
    }
    const quotesEven = quotes % 2 === 0;

    if (!quotesOpened && quotesEven && isValidRange(item)) {
      ranges.push(item);
    } else if (quotesOpened && !quotesEven) {
      quotesOpened = false;
      if (isValidRange(last + item)) {
        ranges.push(last + item);
      }
      last = "";
    } else {
      quotesOpened = true;
      last += `${item},`;
    }
  });
  return ranges;
}

export { DefinedNamesXform };
