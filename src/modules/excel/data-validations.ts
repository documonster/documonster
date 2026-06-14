import type { DataValidation } from "@excel/types";
import { colCache, type DecodedRange } from "@excel/utils/col-cache";

interface ValidationModel {
  [address: string]: DataValidation | undefined;
}

/**
 * Plain-data data-validation registry (de-classed domain model).
 *
 * Keeps the `model` map field (so callers reading `ws.dataValidations.model`
 * are unaffected); add/find/remove become flat helpers.
 */
export interface DataValidationsData {
  model: ValidationModel;
}

/** Create a data-validation registry, optionally seeded from a parsed model. */
export function createDataValidations(model?: ValidationModel): DataValidationsData {
  return { model: model || {} };
}

/** Register a validation at an exact address. */
export function dataValidationAdd(
  dv: DataValidationsData,
  address: string,
  validation: DataValidation
): DataValidation {
  return (dv.model[address] = validation);
}

/**
 * Resolve the validation that applies to `address`: first an exact-address
 * match, then any `range:`-prefixed key whose decoded range contains it.
 */
export function dataValidationFind(
  dv: DataValidationsData,
  address: string
): DataValidation | undefined {
  // First check direct address match
  const direct = dv.model[address];
  if (direct !== undefined) {
    return direct;
  }

  // Check range: prefixed keys in model (from parsing ranges)
  // Only decode address if we see at least one range key.
  let decoded: { row: number; col: number } | undefined;
  for (const key of Object.keys(dv.model)) {
    if (!key.startsWith("range:")) {
      continue;
    }

    decoded ||= colCache.decodeAddress(address);

    const rangeStr = key.slice(6); // Remove "range:" prefix
    const rangeDecoded = colCache.decodeEx(rangeStr);
    if (!("dimensions" in rangeDecoded)) {
      continue;
    }

    const { tl, br } = rangeDecoded as DecodedRange;
    const tlAddr = typeof tl === "string" ? colCache.decodeAddress(tl) : tl;
    const brAddr = typeof br === "string" ? colCache.decodeAddress(br) : br;
    if (
      decoded.row >= tlAddr.row &&
      decoded.row <= brAddr.row &&
      decoded.col >= tlAddr.col &&
      decoded.col <= brAddr.col
    ) {
      return dv.model[key];
    }
  }

  return undefined;
}

/** Clear the validation registered at `address`. */
export function dataValidationRemove(dv: DataValidationsData, address: string): void {
  dv.model[address] = undefined;
}
