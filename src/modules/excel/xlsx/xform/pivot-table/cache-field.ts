import type { CacheField as CacheFieldType, SharedItemValue } from "@excel/core/pivot-table-types";
import { isPivotError } from "@excel/core/pivot-table-types";
import { xmlEncode } from "@xml/encode";

/**
 * Format a Date for OOXML pivot cache output.
 * Excel expects `"2024-01-15T00:00:00"` — no milliseconds, no trailing "Z".
 */
function formatDateForExcel(date: Date): string {
  // Guard against Invalid Date — toISOString() throws RangeError for invalid dates
  if (isNaN(date.getTime())) {
    return "";
  }
  return date.toISOString().replace(/\.\d{3}Z$/, "");
}

/**
 * Render a single `<cacheField>` element as an XML string.
 *
 * This is a pure data-in → string-out transformation with no state or lifecycle,
 * so it is a plain function rather than a class.
 */
function renderCacheField(cf: CacheFieldType): string {
  // PivotCache Field: http://www.datypic.com/sc/ooxml/e-ssml_cacheField-1.html
  // Shared Items: http://www.datypic.com/sc/ooxml/e-ssml_sharedItems-1.html

  // Escape XML special characters in name attribute
  const escapedName = xmlEncode(cf.name);
  const fmtId = cf.numFmtId ?? "0";
  // Extra unknown attributes for roundtrip preservation (e.g. caption, formula)
  const extraStr = buildExtraAttrsStr(cf.extraAttrs);
  // fieldGroup raw XML for roundtrip (rendered after sharedItems per OOXML schema)
  const fgXml = cf.fieldGroupXml ? `\n      ${cf.fieldGroupXml}` : "";

  // --- Branch A: no shared-item children ---
  // For loaded fields, only render <sharedItems> if it was present in the original (sharedItems !== null).
  // For new fields, sharedItems === null means a value-only field that still needs <sharedItems />.
  if (cf.sharedItems === null) {
    if (cf.isLoaded) {
      // Loaded field with no <sharedItems> element — don't add one on roundtrip
      if (fgXml) {
        // Has fieldGroup — need open/close tags to wrap it
        return `<cacheField name="${escapedName}" numFmtId="${fmtId}"${extraStr}>${fgXml}\n    </cacheField>`;
      }
      // R8-B6: No children at all — use self-closing tag to avoid whitespace content
      return `<cacheField name="${escapedName}" numFmtId="${fmtId}"${extraStr} />`;
    }
    // New field — render bare empty <sharedItems />
    let attrStr: string;
    if (cf.minValue !== undefined && cf.maxValue !== undefined) {
      // New numeric-only field (no row/column usage) – compute type attributes
      const integerAttr = cf.containsInteger ? ' containsInteger="1"' : "";
      attrStr = ` containsSemiMixedTypes="0" containsString="0" containsNumber="1"${integerAttr} minValue="${cf.minValue}" maxValue="${cf.maxValue}"`;
    } else {
      attrStr = "";
    }
    return `<cacheField name="${escapedName}" numFmtId="${fmtId}"${extraStr}>
      <sharedItems${attrStr} />${fgXml}
    </cacheField>`;
  }

  if (cf.isLoaded && cf.sharedItems.length === 0) {
    // Roundtrip: use preserved attributes exactly as parsed
    const attrStr = buildLoadedSharedItemsAttrs(cf);
    return `<cacheField name="${escapedName}" numFmtId="${fmtId}"${extraStr}>
      <sharedItems${attrStr} />${fgXml}
    </cacheField>`;
  }

  // --- Branch B: shared items present – render each child element ---

  // For loaded fields, skip recomputation – use preserved attributes directly
  if (cf.isLoaded) {
    const itemXmls = renderSharedItemElements(cf.sharedItems);
    const finalAttrStr = buildLoadedSharedItemsAttrs(cf);
    return `<cacheField name="${escapedName}" numFmtId="${fmtId}"${extraStr}>
      <sharedItems${finalAttrStr} count="${cf.sharedItems.length}">
        ${itemXmls}
      </sharedItems>${fgXml}
    </cacheField>`;
  }

  // New field – analyze content types in a single pass (avoids stack overflow from spread on large arrays)
  let hasString = false;
  let hasNumber = false;
  let hasBoolean = false;
  let hasDate = false;
  let hasNull = false;
  let hasError = false;
  let allInteger = true;
  let minValue = Infinity;
  let maxValue = -Infinity;

  for (const item of cf.sharedItems) {
    if (item === null) {
      hasNull = true;
    } else if (isPivotError(item)) {
      hasError = true;
    } else if (typeof item === "string") {
      hasString = true;
    } else if (typeof item === "number") {
      if (Number.isFinite(item)) {
        hasNumber = true;
        if (item < minValue) {
          minValue = item;
        }
        if (item > maxValue) {
          maxValue = item;
        }
        if (!Number.isInteger(item)) {
          allInteger = false;
        }
      } else {
        // NaN, Infinity, -Infinity are not valid OOXML numeric values — treat as missing
        hasNull = true;
      }
    } else if (typeof item === "boolean") {
      hasBoolean = true;
    } else if (item instanceof Date) {
      hasDate = true;
    }
  }

  // Build sharedItems attributes per OOXML spec:
  // - containsSemiMixedTypes: "0" if no strings at all (default "1")
  // - containsString: "0" if no string items (default "1")
  // - containsNumber: "1" if any numeric items (default "0")
  // - containsInteger: "1" if all numbers are integers (only when containsNumber="1")
  // - containsBlank: "1" if any null/missing items
  // - containsMixedTypes: "1" if multiple non-string types present
  // - containsDate: "1" if any date items (default "0")
  const attrs: string[] = [];
  if (!hasString) {
    // No string items at all (errors, numbers, booleans, dates, nulls are all non-string)
    attrs.push('containsSemiMixedTypes="0"');
    attrs.push('containsString="0"');
  }
  // containsMixedTypes: set when multiple distinct non-string type families are present
  // e.g., number+date, string+number, error+number, etc.
  const typeCount =
    (hasString ? 1 : 0) +
    (hasNumber ? 1 : 0) +
    (hasBoolean ? 1 : 0) +
    (hasDate ? 1 : 0) +
    (hasError ? 1 : 0);
  if (typeCount > 1) {
    attrs.push('containsMixedTypes="1"');
  }
  if (hasNumber) {
    attrs.push('containsNumber="1"');
    if (allInteger) {
      attrs.push('containsInteger="1"');
    }
    attrs.push(`minValue="${minValue}"`);
    attrs.push(`maxValue="${maxValue}"`);
  }
  if (hasDate) {
    attrs.push('containsDate="1"');
  }
  if (hasNull) {
    attrs.push('containsBlank="1"');
  }

  const itemXmls = renderSharedItemElements(cf.sharedItems);
  const attrStr = attrs.length > 0 ? ` ${attrs.join(" ")}` : "";
  return `<cacheField name="${escapedName}" numFmtId="${fmtId}"${extraStr}>
      <sharedItems${attrStr} count="${cf.sharedItems.length}">
        ${itemXmls}
      </sharedItems>${fgXml}
    </cacheField>`;
}

/**
 * Render shared item elements as an XML string.
 */
function renderSharedItemElements(sharedItems: SharedItemValue[]): string {
  const itemXmls: string[] = [];
  for (const item of sharedItems) {
    if (item === null) {
      itemXmls.push("<m />");
    } else if (isPivotError(item)) {
      itemXmls.push(`<e v="${xmlEncode(item.code)}" />`);
    } else if (typeof item === "number") {
      if (Number.isFinite(item)) {
        itemXmls.push(`<n v="${item}" />`);
      } else {
        // NaN, Infinity, -Infinity — not valid in OOXML, render as missing
        itemXmls.push("<m />");
      }
    } else if (typeof item === "boolean") {
      itemXmls.push(`<b v="${item ? "1" : "0"}" />`);
    } else if (item instanceof Date) {
      itemXmls.push(`<d v="${formatDateForExcel(item)}" />`);
    } else {
      // string
      itemXmls.push(`<s v="${xmlEncode(item)}" />`);
    }
  }
  return itemXmls.join("");
}

/**
 * Build extra attributes string for cacheField element (roundtrip preservation).
 */
function buildExtraAttrsStr(extraAttrs: Record<string, string> | undefined): string {
  if (!extraAttrs) {
    return "";
  }
  return Object.entries(extraAttrs)
    .map(([k, v]) => ` ${k}="${xmlEncode(v)}"`)
    .join("");
}

/**
 * Build sharedItems attribute string from preserved loaded attributes.
 * Reconstructs the original attribute order for roundtrip fidelity.
 */
function buildLoadedSharedItemsAttrs(cf: CacheFieldType): string {
  const parts: string[] = [];
  if (cf.containsSemiMixedTypes !== undefined) {
    parts.push(`containsSemiMixedTypes="${cf.containsSemiMixedTypes}"`);
  }
  if (cf.containsNonDate !== undefined) {
    parts.push(`containsNonDate="${cf.containsNonDate}"`);
  }
  if (cf.containsString !== undefined) {
    parts.push(`containsString="${cf.containsString}"`);
  }
  if (cf.containsBlank !== undefined) {
    parts.push(`containsBlank="${cf.containsBlank}"`);
  }
  if (cf.containsMixedTypes !== undefined) {
    parts.push(`containsMixedTypes="${cf.containsMixedTypes}"`);
  }
  // R8-B4: containsNumber/containsInteger are now raw strings ("0"/"1") for roundtrip fidelity
  if (cf.containsNumber !== undefined) {
    parts.push(`containsNumber="${cf.containsNumber}"`);
  }
  if (cf.containsInteger !== undefined) {
    parts.push(`containsInteger="${cf.containsInteger}"`);
  }
  // R8-B5: Emit minValue/maxValue independently of containsNumber — the original XML
  // may have these attributes even without containsNumber="1"
  if (cf.minValue !== undefined) {
    parts.push(`minValue="${cf.minValue}"`);
  }
  if (cf.maxValue !== undefined) {
    parts.push(`maxValue="${cf.maxValue}"`);
  }
  if (cf.containsDate !== undefined) {
    parts.push(`containsDate="${cf.containsDate}"`);
  }
  return parts.length > 0 ? ` ${parts.join(" ")}` : "";
}

export { renderCacheField, formatDateForExcel };
