import type { DataValidation } from "@excel/types";
import { colCache, type DecodedRange } from "@excel/utils/col-cache";

interface ValidationModel {
  [address: string]: DataValidation | undefined;
}

class DataValidations {
  model: ValidationModel;

  constructor(model?: ValidationModel) {
    this.model = model || {};
  }

  add(address: string, validation: DataValidation): DataValidation {
    return (this.model[address] = validation);
  }

  find(address: string): DataValidation | undefined {
    // First check direct address match
    const direct = this.model[address];
    if (direct !== undefined) {
      return direct;
    }

    // Check range: prefixed keys in model (from parsing ranges)
    // Only decode address if we see at least one range key.
    let decoded: { row: number; col: number } | undefined;
    for (const key of Object.keys(this.model)) {
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
        return this.model[key];
      }
    }

    return undefined;
  }

  remove(address: string): void {
    this.model[address] = undefined;
  }
}

export { DataValidations };
