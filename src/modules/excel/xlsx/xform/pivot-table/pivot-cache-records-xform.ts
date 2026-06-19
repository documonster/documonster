import type { PivotTableSource } from "@excel/core/pivot-table";
import type {
  RecordValue,
  ParsedCacheRecords,
  CacheField,
  SharedItemValue
} from "@excel/core/pivot-table-types";
import { isPivotError } from "@excel/core/pivot-table-types";
import { PivotTableError } from "@excel/errors";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { formatDateForExcel } from "@excel/xlsx/xform/pivot-table/cache-field";
import { parseOoxmlDate } from "@utils/utils";
import { xmlEncode } from "@xml/encode";
import type { XmlSink } from "@xml/types";
import { StdDocAttributes } from "@xml/writer";

/** Attribute keys on <pivotCacheRecords> that are individually parsed (not collected into extraRootAttrs). */
const KNOWN_CACHE_RECORDS_ROOT_KEYS = new Set(["xmlns", "xmlns:r", "count"]);

/**
 * Model for generating pivot cache records (with live source)
 */
interface CacheRecordsModel {
  source: PivotTableSource;
  cacheFields: CacheField[];
}

class PivotCacheRecordsXform extends BaseXform<ParsedCacheRecords | null> {
  // Parser state
  private currentRecord: RecordValue[] | null;

  constructor() {
    super();

    this.model = null;
    this.currentRecord = null;
  }

  get tag(): string {
    // http://www.datypic.com/sc/ooxml/e-ssml_pivotCacheRecords.html
    return "pivotCacheRecords";
  }

  reset(): void {
    this.model = null;
    this.currentRecord = null;
  }

  /**
   * Render pivot cache records XML.
   * Supports both newly created models (with PivotTableSource) and loaded models.
   */
  render(xmlStream: XmlSink, model: CacheRecordsModel | ParsedCacheRecords): void {
    // Check if this is a loaded model
    const isLoaded = ("isLoaded" in model && model.isLoaded) || !("source" in model);

    if (isLoaded) {
      this.renderLoaded(xmlStream, model as ParsedCacheRecords);
    } else {
      this.renderNew(xmlStream, model as CacheRecordsModel);
    }
  }

  /**
   * Render newly created pivot cache records
   */
  private renderNew(xmlStream: XmlSink, model: CacheRecordsModel): void {
    const { source, cacheFields } = model;
    // R8-O2: Use Array.isArray for type safety — getSheetValues() returns a sparse array of row arrays
    const sourceBodyRows = source.getSheetValues().slice(2).filter(Array.isArray);

    xmlStream.openXml(StdDocAttributes);
    xmlStream.openNode(this.tag, {
      ...PivotCacheRecordsXform.PIVOT_CACHE_RECORDS_ATTRIBUTES,
      count: sourceBodyRows.length
    });
    xmlStream.writeRaw(this.renderTableNew(sourceBodyRows, cacheFields));
    xmlStream.closeNode();
  }

  /**
   * Render loaded pivot cache records
   */
  private renderLoaded(xmlStream: XmlSink, model: ParsedCacheRecords): void {
    xmlStream.openXml(StdDocAttributes);
    // R8-B11: Use preserved root attributes instead of hardcoded MS namespaces.
    // The base xmlns/xmlns:r are always needed; extra attributes (xmlns:mc, mc:Ignorable, etc.)
    // come from the parsed original if available, otherwise fall back to the hardcoded defaults.
    const rootAttrs: Record<string, string | number> = {
      xmlns: PivotCacheRecordsXform.PIVOT_CACHE_RECORDS_ATTRIBUTES.xmlns,
      "xmlns:r": PivotCacheRecordsXform.PIVOT_CACHE_RECORDS_ATTRIBUTES["xmlns:r"]
    };
    if (model.extraRootAttrs) {
      for (const [k, v] of Object.entries(model.extraRootAttrs)) {
        rootAttrs[k] = v;
      }
    } else {
      // No preserved attributes — use defaults for new-style rendering
      rootAttrs["xmlns:mc"] = PivotCacheRecordsXform.PIVOT_CACHE_RECORDS_ATTRIBUTES["xmlns:mc"];
      rootAttrs["mc:Ignorable"] =
        PivotCacheRecordsXform.PIVOT_CACHE_RECORDS_ATTRIBUTES["mc:Ignorable"];
      rootAttrs["xmlns:xr"] = PivotCacheRecordsXform.PIVOT_CACHE_RECORDS_ATTRIBUTES["xmlns:xr"];
    }
    rootAttrs.count = model.records.length;
    xmlStream.openNode(this.tag, rootAttrs);

    // Render each record
    for (const record of model.records) {
      xmlStream.writeRaw("\n  <r>");
      for (const value of record) {
        xmlStream.writeRaw("\n    ");
        xmlStream.writeRaw(this.renderRecordValue(value));
      }
      xmlStream.writeRaw("\n  </r>");
    }

    xmlStream.closeNode();
  }

  /**
   * Render a single record value to XML
   */
  private renderRecordValue(value: RecordValue): string {
    switch (value.type) {
      case "x":
        return `<x v="${value.value}" />`;
      case "n":
        // Guard against NaN/Infinity — not valid in OOXML, render as missing
        if (!Number.isFinite(value.value)) {
          return "<m />";
        }
        return `<n v="${value.value}" />`;
      case "s":
        return `<s v="${xmlEncode(value.value)}" />`;
      case "b":
        return `<b v="${value.value ? "1" : "0"}" />`;
      case "m":
        return "<m />";
      case "d":
        return `<d v="${formatDateForExcel(value.value)}" />`;
      case "e":
        return `<e v="${xmlEncode(value.value)}" />`;
      default: {
        const _exhaustive: never = value;
        throw new Error(`Unhandled record value type: ${(_exhaustive as any).type}`);
      }
    }
  }

  // Helper methods for rendering new records
  private renderTableNew(sourceBodyRows: unknown[][], cacheFields: CacheField[]): string {
    const parts: string[] = [];
    for (const row of sourceBodyRows) {
      const realRow = row.slice(1);
      parts.push("\n  <r>");
      const fieldCount = Math.min(realRow.length, cacheFields.length);
      for (let i = 0; i < fieldCount; i++) {
        parts.push("\n    " + this.renderCellNew(realRow[i], cacheFields[i].sharedItems));
      }
      // Pad missing columns with <m /> so every record has exactly one value per cacheField (OOXML requirement)
      for (let i = fieldCount; i < cacheFields.length; i++) {
        parts.push("\n    <m />");
      }
      parts.push("\n  </r>");
    }
    return parts.join("");
  }

  private renderCellNew(value: unknown, sharedItems: SharedItemValue[] | null): string {
    // Handle null/undefined/NaN values first — all treated as missing
    if (
      value === null ||
      value === undefined ||
      (typeof value === "number" && !Number.isFinite(value))
    ) {
      // If no shared items, render as missing value directly
      if (sharedItems === null) {
        return "<m />";
      }
      // With shared items, look up null (undefined is treated as null)
      const idx = sharedItems.indexOf(null);
      if (idx >= 0) {
        return `<x v="${idx}" />`;
      }
      // null not in sharedItems — render as missing
      return "<m />";
    }

    // no shared items — render inline by type
    if (sharedItems === null) {
      if (isPivotError(value)) {
        return `<e v="${xmlEncode(value.code)}" />`;
      }
      if (typeof value === "boolean") {
        return `<b v="${value ? "1" : "0"}" />`;
      }
      if (value instanceof Date) {
        return `<d v="${formatDateForExcel(value)}" />`;
      }
      if (Number.isFinite(value)) {
        return `<n v="${value}" />`;
      }
      return `<s v="${xmlEncode(String(value))}" />`;
    }

    // shared items — look up index (type-aware for Date)
    const sharedItemsIndex = findSharedItemIndex(sharedItems, value);
    if (sharedItemsIndex < 0) {
      throw new PivotTableError(
        `${JSON.stringify(value)} not in sharedItems ${JSON.stringify(sharedItems)}`
      );
    }
    return `<x v="${sharedItemsIndex}" />`;
  }

  parseOpen(node: any): boolean {
    const { name, attributes } = node;

    switch (name) {
      case this.tag: {
        // pivotCacheRecords root element
        this.reset();
        // R8-B11: Collect unknown root attributes for roundtrip preservation
        const extraRootAttrs: Record<string, string> = {};
        for (const [k, v] of Object.entries(attributes)) {
          if (!KNOWN_CACHE_RECORDS_ROOT_KEYS.has(k)) {
            extraRootAttrs[k] = String(v);
          }
        }
        this.model = {
          records: [],
          count: parseInt(attributes.count ?? "0", 10),
          isLoaded: true,
          extraRootAttrs: Object.keys(extraRootAttrs).length > 0 ? extraRootAttrs : undefined
        };
        break;
      }

      case "r":
        // Start of a new record
        this.currentRecord = [];
        break;

      case "x":
        // Shared item index
        if (this.currentRecord) {
          this.currentRecord.push({
            type: "x",
            value: parseInt(attributes.v ?? "0", 10)
          });
        }
        break;

      case "n":
        // Numeric value — missing v → treat as missing to avoid fabricating 0
        if (this.currentRecord) {
          if (attributes.v === undefined || attributes.v === "") {
            this.currentRecord.push({ type: "m" });
          } else {
            this.currentRecord.push({
              type: "n",
              value: parseFloat(attributes.v)
            });
          }
        }
        break;

      case "s":
        // String value
        if (this.currentRecord) {
          this.currentRecord.push({
            type: "s",
            value: attributes.v ?? ""
          });
        }
        break;

      case "b":
        // Boolean value
        if (this.currentRecord) {
          this.currentRecord.push({
            type: "b",
            value: attributes.v === "1"
          });
        }
        break;

      case "m":
        // Missing/null value
        if (this.currentRecord) {
          this.currentRecord.push({ type: "m" });
        }
        break;

      case "d":
        // Date value — force UTC parsing (OOXML dates lack "Z" suffix)
        // Missing/empty v → treat as missing value to avoid Invalid Date
        if (this.currentRecord) {
          if (!attributes.v) {
            this.currentRecord.push({ type: "m" });
          } else {
            // R8-B13: Guard against Invalid Date from malformed date strings
            const date = parseOoxmlDate(attributes.v);
            if (isNaN(date.getTime())) {
              this.currentRecord.push({ type: "m" });
            } else {
              this.currentRecord.push({
                type: "d",
                value: date
              });
            }
          }
        }
        break;

      case "e":
        // Error value
        if (this.currentRecord) {
          this.currentRecord.push({
            type: "e",
            value: attributes.v ?? ""
          });
        }
        break;
    }

    return true;
  }

  parseText(_text: string): void {
    // No text content in cache records elements
  }

  parseClose(name: string): boolean {
    switch (name) {
      case this.tag:
        // End of pivotCacheRecords
        return false;

      case "r":
        // End of record - add to model
        if (this.model && this.currentRecord) {
          this.model.records.push(this.currentRecord);
          this.currentRecord = null;
        }
        break;
    }

    return true;
  }

  static readonly PIVOT_CACHE_RECORDS_ATTRIBUTES = {
    xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "mc:Ignorable": "xr",
    "xmlns:xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
  };
}

export { PivotCacheRecordsXform };

/**
 * Find the index of `value` in `sharedItems`, using type-aware comparison.
 * - Date objects are compared by timestamp (getTime()) since === uses reference equality.
 * - PivotErrorValue objects are compared by their code string.
 * - All other types use strict equality (===) via indexOf.
 */
function findSharedItemIndex(sharedItems: SharedItemValue[], value: unknown): number {
  if (value instanceof Date) {
    const ts = value.getTime();
    for (let i = 0; i < sharedItems.length; i++) {
      const item = sharedItems[i];
      if (item instanceof Date && item.getTime() === ts) {
        return i;
      }
    }
    return -1;
  }
  if (isPivotError(value)) {
    for (let i = 0; i < sharedItems.length; i++) {
      const item = sharedItems[i];
      if (isPivotError(item) && item.code === value.code) {
        return i;
      }
    }
    return -1;
  }
  return sharedItems.indexOf(value as SharedItemValue);
}
