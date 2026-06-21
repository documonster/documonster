import type { CacheField, SharedItemValue } from "@excel/core/pivot-table-types";
import { pivotError } from "@excel/core/pivot-table-types";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { RawXmlCollector } from "@excel/xlsx/xform/pivot-table/raw-xml-collector";
import { parseOoxmlDate } from "@utils/utils";
import type { ParseOpenTag } from "@xml/types";

/** Attribute keys on <cacheField> that are individually parsed (not collected into extraAttrs). */
const KNOWN_CACHE_FIELD_KEYS = new Set(["name", "numFmtId"]);

/**
 * Xform for parsing individual <cacheField> elements within a pivot cache definition.
 *
 * Example XML:
 * ```xml
 * <cacheField name="Category" numFmtId="0">
 *   <sharedItems count="3">
 *     <s v="A" />
 *     <s v="B" />
 *     <s v="C" />
 *   </sharedItems>
 * </cacheField>
 *
 * <cacheField name="Value" numFmtId="0">
 *   <sharedItems containsSemiMixedTypes="0" containsString="0"
 *                containsNumber="1" containsInteger="1" minValue="5" maxValue="45" />
 * </cacheField>
 * ```
 */
class CacheFieldXform extends BaseXform<CacheField | null> {
  private inSharedItems = false;
  private fieldGroupCollector = new RawXmlCollector("fieldGroup");

  constructor() {
    super();
    this.model = null;
  }

  get tag(): string {
    return "cacheField";
  }

  reset(): void {
    this.model = null;
    this.inSharedItems = false;
    this.fieldGroupCollector.reset();
  }

  parseOpen(node: ParseOpenTag): boolean {
    const { name, attributes } = node;

    // Collect fieldGroup XML verbatim for roundtrip preservation
    if (this.fieldGroupCollector.active) {
      this.fieldGroupCollector.feedOpen(name, attributes);
      return true;
    }

    switch (name) {
      case "cacheField": {
        // Initialize the model with field name
        // Collect unknown attributes into extraAttrs bag for roundtrip preservation
        const extraAttrs: Record<string, string> = {};
        for (const [k, v] of Object.entries(attributes)) {
          if (!KNOWN_CACHE_FIELD_KEYS.has(k)) {
            extraAttrs[k] = String(v);
          }
        }
        this.model = {
          name: attributes.name ?? "",
          sharedItems: null,
          numFmtId: attributes.numFmtId,
          isLoaded: true,
          extraAttrs: Object.keys(extraAttrs).length > 0 ? extraAttrs : undefined
        };
        break;
      }

      case "fieldGroup":
        // Start collecting fieldGroup XML for roundtrip preservation
        if (this.model) {
          this.fieldGroupCollector.start(attributes);
        }
        break;

      case "sharedItems":
        this.inSharedItems = true;
        // Store numeric field metadata
        if (this.model) {
          // R8-B4: Store containsNumber/containsInteger as raw strings for roundtrip fidelity.
          // Previously "0" was lost because it was parsed as undefined.
          if (attributes.containsNumber !== undefined) {
            this.model.containsNumber = attributes.containsNumber;
          }
          if (attributes.containsInteger !== undefined) {
            this.model.containsInteger = attributes.containsInteger;
          }
          // R8-B7: Guard against NaN from malformed minValue/maxValue
          if (attributes.minValue !== undefined) {
            const parsed = parseFloat(attributes.minValue);
            if (Number.isFinite(parsed)) {
              this.model.minValue = parsed;
            }
          }
          if (attributes.maxValue !== undefined) {
            const parsed = parseFloat(attributes.maxValue);
            if (Number.isFinite(parsed)) {
              this.model.maxValue = parsed;
            }
          }
          // Preserve original sharedItems type attributes for roundtrip fidelity.
          // These are stored as raw strings so we can re-emit them exactly as loaded.
          if (attributes.containsSemiMixedTypes !== undefined) {
            this.model.containsSemiMixedTypes = attributes.containsSemiMixedTypes;
          }
          if (attributes.containsNonDate !== undefined) {
            this.model.containsNonDate = attributes.containsNonDate;
          }
          if (attributes.containsString !== undefined) {
            this.model.containsString = attributes.containsString;
          }
          if (attributes.containsBlank !== undefined) {
            this.model.containsBlank = attributes.containsBlank;
          }
          if (attributes.containsDate !== undefined) {
            this.model.containsDate = attributes.containsDate;
          }
          if (attributes.containsMixedTypes !== undefined) {
            this.model.containsMixedTypes = attributes.containsMixedTypes;
          }
          // Initialize sharedItems array when sharedItems element is present.
          // count="0" means an explicitly empty shared items list (not the same as absent).
          // No count attribute also gets an empty array since the <sharedItems> element
          // itself is present — child elements (<s>, <n>, etc.) will be pushed into it.
          this.model.sharedItems = [];
        }
        break;

      case "s":
      case "n":
      case "b":
      case "e":
      case "m":
      case "d":
        // Shared item value — push to sharedItems array if we're inside <sharedItems>
        if (this.inSharedItems && this.model && this.model.sharedItems !== null) {
          this.model.sharedItems.push(parseSharedItemValue(name, attributes));
        }
        break;
    }

    return true;
  }

  parseText(text: string): void {
    // Forward text to active collector (B3 fix: text nodes in fieldGroup)
    if (this.fieldGroupCollector.active) {
      this.fieldGroupCollector.feedText(text);
    }
  }

  parseClose(name: string): boolean {
    // Handle fieldGroup collection — close tags
    if (this.fieldGroupCollector.active) {
      if (this.fieldGroupCollector.feedClose(name)) {
        if (this.model) {
          this.model.fieldGroupXml = this.fieldGroupCollector.result;
        }
        this.fieldGroupCollector.reset();
      }
      return true;
    }

    switch (name) {
      case "cacheField":
        // End of this cacheField element
        return false;

      case "sharedItems":
        this.inSharedItems = false;
        break;
    }

    return true;
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Convert a SAX-parsed shared item element into its typed value.
 *
 * Tag mapping:
 * - `s` → string
 * - `n` → number
 * - `b` → boolean
 * - `e` → PivotErrorValue
 * - `m` → null (missing)
 * - `d` → Date (UTC)
 */
function parseSharedItemValue(tag: string, attributes: Record<string, string>): SharedItemValue {
  switch (tag) {
    case "s":
      return attributes.v ?? "";
    case "n":
      // Missing v attribute → treat as null (missing) rather than fabricating 0
      if (attributes.v === undefined || attributes.v === "") {
        return null;
      }
      return parseFloat(attributes.v);
    case "b":
      return attributes.v === "1";
    case "e":
      return pivotError(attributes.v ?? "");
    case "m":
      return null;
    case "d": {
      // Missing/empty v attribute → treat as missing value (null) to avoid Invalid Date
      if (!attributes.v) {
        return null;
      }
      // R8-B13: Guard against Invalid Date from malformed date strings
      const date = parseOoxmlDate(attributes.v);
      return isNaN(date.getTime()) ? null : date;
    }
    default:
      return null;
  }
}

export { CacheFieldXform };
