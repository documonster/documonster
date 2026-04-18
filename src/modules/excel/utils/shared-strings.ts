import type { RichText } from "@excel/types";

/** A shared-string entry may be a plain string or a rich-text payload. */
export type SharedStringValue = string | { richText: RichText[] };

/**
 * Canonical JSON serializer with sorted object keys.
 *
 * Used to derive an insertion-order-independent dedupe key for rich-text
 * shared-string entries. Two semantically identical run/font objects must
 * map to the same key regardless of the order their properties were assigned.
 */
function canonicalStringify(value: unknown): string {
  if (value === null || typeof value !== "object") {
    return JSON.stringify(value);
  }
  if (Array.isArray(value)) {
    return `[${value.map(canonicalStringify).join(",")}]`;
  }
  const keys = Object.keys(value as Record<string, unknown>).sort();
  const obj = value as Record<string, unknown>;
  return `{${keys.map(k => `${JSON.stringify(k)}:${canonicalStringify(obj[k])}`).join(",")}}`;
}

class SharedStrings {
  declare private _values: SharedStringValue[];
  declare private _totalRefs: number;
  // String-to-index map. For rich-text entries, the key is a canonical
  // (key-order-independent) JSON serialization so semantically equivalent
  // rich-text payloads dedupe even if their object properties were assigned
  // in different orders.
  declare private _hash: Record<string, number>;

  constructor() {
    this._values = [];
    this._totalRefs = 0;
    this._hash = Object.create(null);
  }

  get count(): number {
    return this._values.length;
  }

  get values(): SharedStringValue[] {
    return this._values;
  }

  get totalRefs(): number {
    return this._totalRefs;
  }

  getString(index: number): SharedStringValue {
    return this._values[index];
  }

  add(value: SharedStringValue): number {
    const key =
      typeof value === "string" ? `s:${value}` : `r:${canonicalStringify(value.richText)}`;
    let index = this._hash[key];
    if (index === undefined) {
      index = this._hash[key] = this._values.length;
      this._values.push(value);
    }
    this._totalRefs++;
    return index;
  }
}

export { SharedStrings };
