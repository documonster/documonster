/**
 * ExternalLinkXform — read and write `xl/externalLinks/externalLinkN.xml`.
 *
 * This part describes one external workbook referenced by the current file:
 * the list of sheets inside the foreign workbook, plus an optional cache of
 * the primitive values at each referenced address. The `r:id` attribute on
 * `<externalBook>` points (via the neighbouring `_rels/externalLinkN.xml.rels`)
 * at the actual file; that resolution is handled by the writer/reader in
 * `xlsx.browser.ts`, not here.
 *
 * Schema shape:
 *
 *   <externalLink xmlns=".../spreadsheetml/2006/main"
 *                 xmlns:r=".../relationships">
 *     <externalBook r:id="rId1">
 *       <sheetNames>
 *         <sheetName val="Sheet1"/>
 *         <sheetName val="Sheet2"/>
 *       </sheetNames>
 *       <sheetDataSet>
 *         <sheetData sheetId="0">                (0-based index into sheetNames)
 *           <row r="1">
 *             <cell r="A1" t="n"><v>123</v></cell>
 *             <cell r="B1" t="str"><v>hello</v></cell>
 *           </row>
 *           <row r="2">
 *             <cell r="A2" t="b"><v>1</v></cell>
 *           </row>
 *         </sheetData>
 *         <sheetData sheetId="1" refreshError="1"/>   (no cached values)
 *       </sheetDataSet>
 *     </externalBook>
 *   </externalLink>
 *
 * The `t` attribute on `<cell>` mirrors the inline-value types used inside a
 * worksheet's `<c>` element: "n" (number, default), "str" (string), "b"
 * (boolean), "e" (error). We don't emit shared-string indices here because
 * the externalLink part is self-contained — each cached string is serialised
 * inline as plain text inside `<v>`.
 *
 * Parser state machine:
 *   externalLink → externalBook → sheetNames → sheetName  (stores names in order)
 *                            \→ sheetDataSet → sheetData → row → cell → v  (values)
 *
 * The parser tolerates missing cache data (Excel's `refreshError="1"` mode);
 * sheets without a matching `<sheetData>` simply have no entry in
 * `cachedValues`.
 */

import type { ExternalLinkCachedSheet, ExternalLinkModel } from "@excel/core/workbook.browser";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { XmlSink } from "@xml/types";
import { StdDocAttributes } from "@xml/writer";

const NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** Pretty attribute constant used on the root element. */
const ROOT_ATTRIBUTES = { xmlns: NAMESPACE, "xmlns:r": R_NAMESPACE };

/**
 * Parsed shape of a single externalLinkN.xml part. Callers (xlsx.browser.ts)
 * merge `rId` with the matching `workbookRels` entry to recover the model's
 * `target` / `targetMode`, since those live in the sibling `.rels` file.
 */
export interface ParsedExternalLink {
  /** r:id on <externalBook>. Points at the rel inside externalLinkN.xml.rels. */
  externalBookRId: string;
  sheetNames: string[];
  cachedValues: Record<string, ExternalLinkCachedSheet>;
}

class ExternalLinkXform extends BaseXform<ParsedExternalLink> {
  /** Sheet name indexed by sheetId (0-based) — filled during parse. */
  private sheetNamesByIndex: string[] = [];
  /** Current sheet's accumulated cached cells. */
  private currentSheetCells: ExternalLinkCachedSheet | null = null;
  /** Current sheet's name (looked up from sheetId). */
  private currentSheetName: string | null = null;
  /** Current cell being parsed: address + type. */
  private currentCellAddress: string | null = null;
  private currentCellType: string = "n";
  /** Accumulated text content of the current <v> element. */
  private currentCellValue: string | null = null;
  /** Which element we're currently inside. */
  private inV = false;

  get tag(): string {
    return "externalLink";
  }

  reset(): void {
    super.reset();
    this.sheetNamesByIndex = [];
    this.currentSheetCells = null;
    this.currentSheetName = null;
    this.currentCellAddress = null;
    this.currentCellType = "n";
    this.currentCellValue = null;
    this.inV = false;
  }

  // ==========================================================================
  // Rendering
  // ==========================================================================

  render(xmlStream: XmlSink, model: ExternalLinkModel): void {
    xmlStream.openXml(StdDocAttributes);
    xmlStream.openNode("externalLink", ROOT_ATTRIBUTES);

    // We always use "rId1" for the externalBook ref — the externalLinkN.rels
    // file only ever has a single relationship (the externalLinkPath one).
    xmlStream.openNode("externalBook", { "r:id": "rId1" });

    if (model.sheetNames.length > 0) {
      xmlStream.openNode("sheetNames");
      for (const name of model.sheetNames) {
        xmlStream.leafNode("sheetName", { val: name });
      }
      xmlStream.closeNode();
    }

    // Emit cached values when supplied. We must always emit <sheetDataSet>
    // if any sheet has cached values — it's fine to omit when none do.
    const cache = model.cachedValues ?? {};
    const hasAnyCache = Object.keys(cache).some(name => {
      const sheet = cache[name];
      return sheet && Object.keys(sheet).length > 0;
    });

    // Excel requires <sheetDataSet> to be present even when empty —
    // omitting it causes "file cannot be opened" errors. Each declared
    // sheet must have a corresponding <sheetData sheetId="N"/> entry.
    xmlStream.openNode("sheetDataSet");
    for (let i = 0; i < model.sheetNames.length; i++) {
      const sheetName = model.sheetNames[i];
      const cells = hasAnyCache ? cache[sheetName] : undefined;
      if (cells && Object.keys(cells).length > 0) {
        renderSheetData(xmlStream, i, cells);
      } else {
        xmlStream.leafNode("sheetData", { sheetId: String(i) });
      }
    }
    xmlStream.closeNode();

    xmlStream.closeNode(); // </externalBook>
    xmlStream.closeNode(); // </externalLink>
  }

  // ==========================================================================
  // Parsing
  // ==========================================================================

  parseOpen(node: any): boolean {
    const name = node.name as string;

    switch (name) {
      case "externalLink":
        this.model = {
          externalBookRId: "",
          sheetNames: [],
          cachedValues: {}
        };
        this.sheetNamesByIndex = [];
        return true;

      case "externalBook":
        if (this.model) {
          this.model.externalBookRId = node.attributes["r:id"] ?? "";
        }
        return true;

      case "sheetName":
        if (this.model) {
          const val = node.attributes.val ?? "";
          this.model.sheetNames.push(val);
          this.sheetNamesByIndex.push(val);
        }
        return true;

      case "sheetData": {
        const sheetIdRaw = node.attributes.sheetId;
        const sheetId = sheetIdRaw !== undefined ? parseInt(sheetIdRaw, 10) : NaN;
        this.currentSheetName = Number.isFinite(sheetId)
          ? (this.sheetNamesByIndex[sheetId] ?? null)
          : null;
        this.currentSheetCells = this.currentSheetName ? {} : null;
        return true;
      }

      case "cell":
        this.currentCellAddress = (node.attributes.r as string | undefined) ?? null;
        this.currentCellType = (node.attributes.t as string | undefined) ?? "n";
        this.currentCellValue = null;
        return true;

      case "v":
        this.inV = true;
        this.currentCellValue = "";
        return true;

      default:
        return false;
    }
  }

  parseText(text: string): void {
    if (this.inV) {
      this.currentCellValue = (this.currentCellValue ?? "") + text;
    }
  }

  parseClose(name: string): boolean {
    switch (name) {
      case "v":
        this.inV = false;
        return true;

      case "cell":
        if (this.model && this.currentSheetCells && this.currentCellAddress) {
          this.currentSheetCells[this.currentCellAddress.toUpperCase()] = decodePrimitive(
            this.currentCellType,
            this.currentCellValue
          );
        }
        this.currentCellAddress = null;
        this.currentCellType = "n";
        this.currentCellValue = null;
        return true;

      case "sheetData":
        if (
          this.model &&
          this.currentSheetName &&
          this.currentSheetCells &&
          Object.keys(this.currentSheetCells).length > 0
        ) {
          this.model.cachedValues[this.currentSheetName] = this.currentSheetCells;
        }
        this.currentSheetName = null;
        this.currentSheetCells = null;
        return true;

      case "externalLink":
        // Document root closed — stop parsing
        return false;

      default:
        return true;
    }
  }
}

// ==========================================================================
// Serialisation helpers
// ==========================================================================

function renderSheetData(
  xmlStream: XmlSink,
  sheetIdx: number,
  cells: ExternalLinkCachedSheet
): void {
  // Group cells by row so the XML tree is well-formed. Excel tolerates any
  // ordering, but grouping by row matches the schema contract and keeps the
  // output stable for diffing.
  const rows = groupByRow(cells);
  xmlStream.openNode("sheetData", { sheetId: sheetIdx });
  for (const row of rows) {
    xmlStream.openNode("row", { r: row.rowNum });
    for (const { address, value } of row.cells) {
      renderCell(xmlStream, address, value);
    }
    xmlStream.closeNode();
  }
  xmlStream.closeNode();
}

interface GroupedRow {
  rowNum: number;
  cells: Array<{ address: string; value: string | number | boolean | null }>;
}

function groupByRow(cells: ExternalLinkCachedSheet): GroupedRow[] {
  const byRow = new Map<number, GroupedRow>();
  for (const rawAddress of Object.keys(cells)) {
    const address = rawAddress.toUpperCase();
    const rowNum = extractRowNumber(address);
    if (rowNum === undefined) {
      // Addresses that don't match the A1 shape are skipped rather than
      // corrupting the file. An address like "A1" or "$B$12" yields a valid
      // row number; anything else would produce malformed XML.
      continue;
    }
    let bucket = byRow.get(rowNum);
    if (!bucket) {
      bucket = { rowNum, cells: [] };
      byRow.set(rowNum, bucket);
    }
    bucket.cells.push({ address: stripAbsoluteMarkers(address), value: cells[rawAddress] });
  }
  return [...byRow.values()].sort((a, b) => a.rowNum - b.rowNum);
}

function extractRowNumber(address: string): number | undefined {
  // Strip leading $ and column letters, keep trailing digits.
  const match = /^\$?[A-Z]+\$?(\d+)$/.exec(address);
  return match ? parseInt(match[1], 10) : undefined;
}

function stripAbsoluteMarkers(address: string): string {
  return address.replace(/\$/g, "");
}

function renderCell(
  xmlStream: XmlSink,
  address: string,
  value: string | number | boolean | null
): void {
  if (value === null || value === undefined) {
    // An explicit blank — emit the address with no type/value so Excel shows
    // an empty cell when displaying cached values.
    xmlStream.leafNode("cell", { r: address });
    return;
  }

  if (typeof value === "number") {
    xmlStream.openNode("cell", { r: address });
    xmlStream.leafNode("v", undefined, String(value));
    xmlStream.closeNode();
    return;
  }

  if (typeof value === "boolean") {
    xmlStream.openNode("cell", { r: address, t: "b" });
    xmlStream.leafNode("v", undefined, value ? "1" : "0");
    xmlStream.closeNode();
    return;
  }

  // Error values like "#DIV/0!", "#REF!", "#N/A" etc. are stored as t="e".
  // Detect them so the round-trip preserves the type correctly (otherwise
  // Excel would display the error literal as a plain string).
  if (isErrorLiteral(value)) {
    xmlStream.openNode("cell", { r: address, t: "e" });
    xmlStream.leafNode("v", undefined, value);
    xmlStream.closeNode();
    return;
  }

  // String — emit t="str" with the text inside <v>.
  xmlStream.openNode("cell", { r: address, t: "str" });
  xmlStream.leafNode("v", undefined, value);
  xmlStream.closeNode();
}

// ==========================================================================
// Deserialisation helpers
// ==========================================================================

/** Excel error literals that should be written with `t="e"`. */
const ERROR_LITERALS = new Set([
  "#NULL!",
  "#DIV/0!",
  "#VALUE!",
  "#REF!",
  "#NAME?",
  "#NUM!",
  "#N/A",
  "#GETTING_DATA",
  "#SPILL!",
  "#CALC!",
  "#CONNECT!",
  "#BLOCKED!",
  "#UNKNOWN!",
  "#FIELD!",
  "#PYTHON!"
]);

function isErrorLiteral(value: string): boolean {
  return ERROR_LITERALS.has(value);
}

function decodePrimitive(type: string, raw: string | null): string | number | boolean | null {
  if (raw === null || raw === "") {
    return null;
  }
  switch (type) {
    case "b":
      return raw !== "0" && raw.toLowerCase() !== "false";
    case "e":
      // Preserve the error literal as a string — consumers can match on it.
      return raw;
    case "str":
      return raw;
    case "n":
    default: {
      const n = Number(raw);
      return Number.isFinite(n) ? n : raw;
    }
  }
}

export { ExternalLinkXform };
