/**
 * Financial Functions — Native RuntimeValue implementation.
 */

import { excelToDate } from "@utils/utils.base";

import type { RuntimeValue, ErrorValue } from "../runtime/values";
import { RVKind, ERRORS, isError, isArray, toNumberRV, rvNumber } from "../runtime/values";

// ============================================================================
// Type alias for native function signature
// ============================================================================

type NativeFunction = (args: RuntimeValue[]) => RuntimeValue;

// ============================================================================
// Internal helpers
// ============================================================================

/** Extract numbers from a list of RuntimeValue args (for NPV/IRR/XNPV/XIRR/MIRR).
 *  Arrays: only pick Number cells (skip booleans/strings/blanks — Excel behavior).
 *  Scalars: coerce via toNumberRV.
 *  Returns an ErrorValue on the first error encountered, otherwise a number[]. */
function flattenNumbersRV(args: RuntimeValue[]): number[] | ErrorValue {
  const result: number[] = [];
  for (const arg of args) {
    if (isArray(arg)) {
      for (const row of arg.rows) {
        for (const cell of row) {
          if (cell.kind === RVKind.Error) {
            return cell;
          }
          if (cell.kind === RVKind.Number) {
            result.push(cell.value);
          }
        }
      }
    } else {
      const n = toNumberRV(arg);
      if (isError(n)) {
        return n;
      }
      result.push(n.value);
    }
  }
  return result;
}

/** Type guard: is the result of flattenNumbersRV an error? */
function isFlattenError(v: number[] | ErrorValue): v is ErrorValue {
  return !Array.isArray(v);
}

// ============================================================================
// Financial Functions
// ============================================================================

export const fnPMT: NativeFunction = args => {
  const rate = toNumberRV(args[0]);
  if (isError(rate)) {
    return rate;
  }
  const nper = toNumberRV(args[1]);
  if (isError(nper)) {
    return nper;
  }
  const pv = toNumberRV(args[2]);
  if (isError(pv)) {
    return pv;
  }
  const fv = args.length > 3 ? toNumberRV(args[3]) : rvNumber(0);
  if (isError(fv)) {
    return fv;
  }
  const type = args.length > 4 ? toNumberRV(args[4]) : rvNumber(0);
  if (isError(type)) {
    return type;
  }
  if (rate.value === 0) {
    return rvNumber(-(pv.value + fv.value) / nper.value);
  }
  const pvif = Math.pow(1 + rate.value, nper.value);
  return rvNumber(
    -(rate.value * (pv.value * pvif + fv.value)) / (pvif - 1) / (1 + rate.value * type.value)
  );
};

export const fnFV: NativeFunction = args => {
  const rate = toNumberRV(args[0]);
  if (isError(rate)) {
    return rate;
  }
  const nper = toNumberRV(args[1]);
  if (isError(nper)) {
    return nper;
  }
  const pmt = toNumberRV(args[2]);
  if (isError(pmt)) {
    return pmt;
  }
  const pv = args.length > 3 ? toNumberRV(args[3]) : rvNumber(0);
  if (isError(pv)) {
    return pv;
  }
  const type = args.length > 4 ? toNumberRV(args[4]) : rvNumber(0);
  if (isError(type)) {
    return type;
  }
  if (rate.value === 0) {
    return rvNumber(-(pv.value + pmt.value * nper.value));
  }
  const pvif = Math.pow(1 + rate.value, nper.value);
  return rvNumber(
    -(pv.value * pvif + pmt.value * (1 + rate.value * type.value) * ((pvif - 1) / rate.value))
  );
};

export const fnPV: NativeFunction = args => {
  const rate = toNumberRV(args[0]);
  if (isError(rate)) {
    return rate;
  }
  const nper = toNumberRV(args[1]);
  if (isError(nper)) {
    return nper;
  }
  const pmt = toNumberRV(args[2]);
  if (isError(pmt)) {
    return pmt;
  }
  const fv = args.length > 3 ? toNumberRV(args[3]) : rvNumber(0);
  if (isError(fv)) {
    return fv;
  }
  const type = args.length > 4 ? toNumberRV(args[4]) : rvNumber(0);
  if (isError(type)) {
    return type;
  }
  if (rate.value === 0) {
    return rvNumber(-pmt.value * nper.value - fv.value);
  }
  const pvif = Math.pow(1 + rate.value, nper.value);
  return rvNumber(
    -(fv.value + pmt.value * (1 + rate.value * type.value) * ((pvif - 1) / rate.value)) / pvif
  );
};

export const fnNPV: NativeFunction = args => {
  const rate = toNumberRV(args[0]);
  if (isError(rate)) {
    return rate;
  }
  const values: number[] = [];
  for (let i = 1; i < args.length; i++) {
    const a = args[i];
    if (isArray(a)) {
      for (const row of a.rows) {
        for (const cell of row) {
          if (cell.kind === RVKind.Number) {
            values.push(cell.value);
          }
        }
      }
    } else {
      const n = toNumberRV(a);
      if (isError(n)) {
        return n;
      }
      values.push(n.value);
    }
  }
  let npv = 0;
  for (let i = 0; i < values.length; i++) {
    npv += values[i] / Math.pow(1 + rate.value, i + 1);
  }
  return rvNumber(npv);
};

export const fnIRR: NativeFunction = args => {
  if (!isArray(args[0])) {
    return ERRORS.VALUE;
  }
  const values: number[] = [];
  for (const row of args[0].rows) {
    for (const cell of row) {
      if (cell.kind === RVKind.Number) {
        values.push(cell.value);
      }
    }
  }
  if (values.length < 2) {
    return ERRORS.NUM;
  }
  const guessRV = args.length > 1 ? toNumberRV(args[1]) : rvNumber(0.1);
  if (isError(guessRV)) {
    return guessRV;
  }
  let g = guessRV.value;
  // Newton-Raphson
  for (let iter = 0; iter < 100; iter++) {
    let npv = 0;
    let dnpv = 0;
    for (let i = 0; i < values.length; i++) {
      npv += values[i] / Math.pow(1 + g, i);
      dnpv -= (i * values[i]) / Math.pow(1 + g, i + 1);
    }
    if (Math.abs(dnpv) < 1e-15) {
      break;
    }
    const newGuess = g - npv / dnpv;
    if (Math.abs(newGuess - g) < 1e-10) {
      return rvNumber(newGuess);
    }
    g = newGuess;
  }
  // Did not converge after 100 iterations
  return ERRORS.NUM;
};

export const fnNPER: NativeFunction = args => {
  const rate = toNumberRV(args[0]);
  if (isError(rate)) {
    return rate;
  }
  const pmt = toNumberRV(args[1]);
  if (isError(pmt)) {
    return pmt;
  }
  const pv = toNumberRV(args[2]);
  if (isError(pv)) {
    return pv;
  }
  const fv = args.length > 3 ? toNumberRV(args[3]) : rvNumber(0);
  if (isError(fv)) {
    return fv;
  }
  const type = args.length > 4 ? toNumberRV(args[4]) : rvNumber(0);
  if (isError(type)) {
    return type;
  }
  if (rate.value === 0) {
    return rvNumber(-(pv.value + fv.value) / pmt.value);
  }
  const num = pmt.value * (1 + rate.value * type.value) - fv.value * rate.value;
  const den = pv.value * rate.value + pmt.value * (1 + rate.value * type.value);
  if (num / den <= 0) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.log(num / den) / Math.log(1 + rate.value));
};

export const fnRATE: NativeFunction = args => {
  const nper = toNumberRV(args[0]);
  if (isError(nper)) {
    return nper;
  }
  const pmt = toNumberRV(args[1]);
  if (isError(pmt)) {
    return pmt;
  }
  const pv = toNumberRV(args[2]);
  if (isError(pv)) {
    return pv;
  }
  const fv = args.length > 3 ? toNumberRV(args[3]) : rvNumber(0);
  if (isError(fv)) {
    return fv;
  }
  const type = args.length > 4 ? toNumberRV(args[4]) : rvNumber(0);
  if (isError(type)) {
    return type;
  }
  const guess = args.length > 5 ? toNumberRV(args[5]) : rvNumber(0.1);
  if (isError(guess)) {
    return guess;
  }
  let g = guess.value;
  const nperV = nper.value;
  const pmtV = pmt.value;
  const pvV = pv.value;
  const fvV = fv.value;
  const typeV = type.value;
  // Newton-Raphson
  for (let iter = 0; iter < 100; iter++) {
    if (g <= -1) {
      g = -0.99;
    }
    const pvif = Math.pow(1 + g, nperV);
    const fvifa = (pvif - 1) / g;
    const f = pvV * pvif + pmtV * (1 + g * typeV) * fvifa + fvV;
    const df =
      nperV * pvV * Math.pow(1 + g, nperV - 1) +
      (pmtV * (1 + g * typeV) * (nperV * Math.pow(1 + g, nperV - 1) * g - pvif + 1)) / (g * g) +
      (typeV ? pmtV * fvifa : 0);
    if (Math.abs(df) < 1e-15) {
      break;
    }
    const newGuess = g - f / df;
    if (Math.abs(newGuess - g) < 1e-10) {
      return rvNumber(newGuess);
    }
    g = newGuess;
  }
  // Did not converge after 100 iterations
  return ERRORS.NUM;
};

export const fnSLN: NativeFunction = args => {
  const cost = toNumberRV(args[0]);
  if (isError(cost)) {
    return cost;
  }
  const salvage = toNumberRV(args[1]);
  if (isError(salvage)) {
    return salvage;
  }
  const life = toNumberRV(args[2]);
  if (isError(life)) {
    return life;
  }
  if (life.value === 0) {
    return ERRORS.DIV0;
  }
  return rvNumber((cost.value - salvage.value) / life.value);
};

export const fnDB: NativeFunction = args => {
  const cost = toNumberRV(args[0]);
  if (isError(cost)) {
    return cost;
  }
  const salvage = toNumberRV(args[1]);
  if (isError(salvage)) {
    return salvage;
  }
  const life = toNumberRV(args[2]);
  if (isError(life)) {
    return life;
  }
  const period = toNumberRV(args[3]);
  if (isError(period)) {
    return period;
  }
  const month = args.length > 4 ? toNumberRV(args[4]) : rvNumber(12);
  if (isError(month)) {
    return month;
  }
  if (life.value === 0 || cost.value === 0) {
    return rvNumber(0);
  }
  const rate = Math.round((1 - Math.pow(salvage.value / cost.value, 1 / life.value)) * 1000) / 1000;
  let totalDepreciation = 0;
  let depn: number = 0;
  for (let p = 1; p <= period.value; p++) {
    if (p === 1) {
      depn = (cost.value * rate * month.value) / 12;
    } else if (p === Math.floor(life.value) + 1) {
      depn = ((cost.value - totalDepreciation) * rate * (12 - month.value)) / 12;
    } else {
      depn = (cost.value - totalDepreciation) * rate;
    }
    totalDepreciation += depn;
  }
  return rvNumber(depn);
};

export const fnDDB: NativeFunction = args => {
  const cost = toNumberRV(args[0]);
  if (isError(cost)) {
    return cost;
  }
  const salvage = toNumberRV(args[1]);
  if (isError(salvage)) {
    return salvage;
  }
  const life = toNumberRV(args[2]);
  if (isError(life)) {
    return life;
  }
  const period = toNumberRV(args[3]);
  if (isError(period)) {
    return period;
  }
  const factor = args.length > 4 ? toNumberRV(args[4]) : rvNumber(2);
  if (isError(factor)) {
    return factor;
  }
  let bookValue = cost.value;
  let depn = 0;
  for (let p = 1; p <= period.value; p++) {
    depn = Math.min(bookValue * (factor.value / life.value), bookValue - salvage.value);
    if (depn < 0) {
      depn = 0;
    }
    bookValue -= depn;
  }
  return rvNumber(depn);
};

/** Internal PMT computation that returns a raw number (for IPMT/PPMT/CUMPRINC/CUMIPMT). */
function pmtRaw(rate: number, nper: number, pv: number, fv: number, type: number): number {
  if (rate === 0) {
    return -(pv + fv) / nper;
  }
  const pvif = Math.pow(1 + rate, nper);
  return -(rate * (pv * pvif + fv)) / (pvif - 1) / (1 + rate * type);
}

/** Internal IPMT computation that returns a raw number. */
function ipmtRaw(
  rate: number,
  per: number,
  nper: number,
  pv: number,
  fv: number,
  type: number
): number {
  const pmt = pmtRaw(rate, nper, pv, fv, type);
  if (rate === 0) {
    return 0;
  }
  // Compute FV of original loan at period (per-1)
  let bal = pv;
  for (let i = 1; i < per; i++) {
    bal = bal * (1 + rate) + pmt * (1 + rate * type);
  }
  const ipmt = type === 1 && per === 1 ? 0 : bal * rate;
  return type === 1 ? ipmt / (1 + rate) : ipmt;
}

export const fnIPMT: NativeFunction = args => {
  const rate = toNumberRV(args[0]);
  if (isError(rate)) {
    return rate;
  }
  const per = toNumberRV(args[1]);
  if (isError(per)) {
    return per;
  }
  const nper = toNumberRV(args[2]);
  if (isError(nper)) {
    return nper;
  }
  const pv = toNumberRV(args[3]);
  if (isError(pv)) {
    return pv;
  }
  const fv = args.length > 4 ? toNumberRV(args[4]) : rvNumber(0);
  if (isError(fv)) {
    return fv;
  }
  const type = args.length > 5 ? toNumberRV(args[5]) : rvNumber(0);
  if (isError(type)) {
    return type;
  }
  return rvNumber(ipmtRaw(rate.value, per.value, nper.value, pv.value, fv.value, type.value));
};

export const fnPPMT: NativeFunction = args => {
  const rate = toNumberRV(args[0]);
  if (isError(rate)) {
    return rate;
  }
  const per = toNumberRV(args[1]);
  if (isError(per)) {
    return per;
  }
  const nper = toNumberRV(args[2]);
  if (isError(nper)) {
    return nper;
  }
  const pv = toNumberRV(args[3]);
  if (isError(pv)) {
    return pv;
  }
  const fv = args.length > 4 ? toNumberRV(args[4]) : rvNumber(0);
  if (isError(fv)) {
    return fv;
  }
  const type = args.length > 5 ? toNumberRV(args[5]) : rvNumber(0);
  if (isError(type)) {
    return type;
  }
  const pmtVal = pmtRaw(rate.value, nper.value, pv.value, fv.value, type.value);
  const ipmtVal = ipmtRaw(rate.value, per.value, nper.value, pv.value, fv.value, type.value);
  return rvNumber(pmtVal - ipmtVal);
};

export const fnEFFECT: NativeFunction = args => {
  const nomRate = toNumberRV(args[0]);
  if (isError(nomRate)) {
    return nomRate;
  }
  const npery = toNumberRV(args[1]);
  if (isError(npery)) {
    return npery;
  }
  if (nomRate.value <= 0 || npery.value < 1) {
    return ERRORS.NUM;
  }
  return rvNumber(
    Math.pow(1 + nomRate.value / Math.floor(npery.value), Math.floor(npery.value)) - 1
  );
};

export const fnNOMINAL: NativeFunction = args => {
  const effRate = toNumberRV(args[0]);
  if (isError(effRate)) {
    return effRate;
  }
  const npery = toNumberRV(args[1]);
  if (isError(npery)) {
    return npery;
  }
  if (effRate.value <= 0 || npery.value < 1) {
    return ERRORS.NUM;
  }
  const np = Math.floor(npery.value);
  return rvNumber(np * (Math.pow(effRate.value + 1, 1 / np) - 1));
};

export const fnXNPV: NativeFunction = args => {
  const rate = toNumberRV(args[0]);
  if (isError(rate)) {
    return rate;
  }
  if (!isArray(args[1]) || !isArray(args[2])) {
    return ERRORS.VALUE;
  }
  const valuesResult = flattenNumbersRV([args[1]]);
  if (isFlattenError(valuesResult)) {
    return valuesResult;
  }
  const datesResult = flattenNumbersRV([args[2]]);
  if (isFlattenError(datesResult)) {
    return datesResult;
  }
  const values = valuesResult;
  const dates = datesResult;
  if (values.length === 0 || values.length !== dates.length) {
    return ERRORS.NUM;
  }
  const d0 = dates[0];
  let npv = 0;
  for (let i = 0; i < values.length; i++) {
    npv += values[i] / Math.pow(1 + rate.value, (dates[i] - d0) / 365);
  }
  return rvNumber(npv);
};

export const fnXIRR: NativeFunction = args => {
  if (!isArray(args[0]) || !isArray(args[1])) {
    return ERRORS.VALUE;
  }
  const valuesResult = flattenNumbersRV([args[0]]);
  if (isFlattenError(valuesResult)) {
    return valuesResult;
  }
  const datesResult = flattenNumbersRV([args[1]]);
  if (isFlattenError(datesResult)) {
    return datesResult;
  }
  const values = valuesResult;
  const dates = datesResult;
  if (values.length < 2 || values.length !== dates.length) {
    return ERRORS.NUM;
  }
  const guessRV = args.length > 2 ? toNumberRV(args[2]) : rvNumber(0.1);
  if (isError(guessRV)) {
    return guessRV;
  }
  const d0 = dates[0];
  let g = guessRV.value;
  for (let iter = 0; iter < 100; iter++) {
    let npv = 0;
    let dnpv = 0;
    for (let i = 0; i < values.length; i++) {
      const t = (dates[i] - d0) / 365;
      npv += values[i] / Math.pow(1 + g, t);
      dnpv -= (t * values[i]) / Math.pow(1 + g, t + 1);
    }
    if (Math.abs(dnpv) < 1e-15) {
      break;
    }
    const newG = g - npv / dnpv;
    if (Math.abs(newG - g) < 1e-10) {
      return rvNumber(newG);
    }
    g = newG;
  }
  return ERRORS.NUM;
};

export const fnMIRR: NativeFunction = args => {
  if (!isArray(args[0])) {
    return ERRORS.VALUE;
  }
  const valuesResult = flattenNumbersRV([args[0]]);
  if (isFlattenError(valuesResult)) {
    return valuesResult;
  }
  const values = valuesResult;
  const financeRate = toNumberRV(args[1]);
  if (isError(financeRate)) {
    return financeRate;
  }
  const reinvestRate = toNumberRV(args[2]);
  if (isError(reinvestRate)) {
    return reinvestRate;
  }
  const n = values.length;
  if (n < 2) {
    return ERRORS.NUM;
  }
  let npvPos = 0;
  let npvNeg = 0;
  for (let i = 0; i < n; i++) {
    if (values[i] >= 0) {
      npvPos += values[i] * Math.pow(1 + reinvestRate.value, n - 1 - i);
    } else {
      npvNeg += values[i] / Math.pow(1 + financeRate.value, i);
    }
  }
  if (npvNeg === 0) {
    return ERRORS.DIV0;
  }
  return rvNumber(Math.pow(-npvPos / npvNeg, 1 / (n - 1)) - 1);
};

export const fnISPMT: NativeFunction = args => {
  const rate = toNumberRV(args[0]);
  if (isError(rate)) {
    return rate;
  }
  const per = toNumberRV(args[1]);
  if (isError(per)) {
    return per;
  }
  const nper = toNumberRV(args[2]);
  if (isError(nper)) {
    return nper;
  }
  const pv = toNumberRV(args[3]);
  if (isError(pv)) {
    return pv;
  }
  return rvNumber(pv.value * rate.value * (per.value / nper.value - 1));
};

export const fnCUMPRINC: NativeFunction = args => {
  const rate = toNumberRV(args[0]);
  if (isError(rate)) {
    return rate;
  }
  const nper = toNumberRV(args[1]);
  if (isError(nper)) {
    return nper;
  }
  const pv = toNumberRV(args[2]);
  if (isError(pv)) {
    return pv;
  }
  const startPeriod = toNumberRV(args[3]);
  if (isError(startPeriod)) {
    return startPeriod;
  }
  const endPeriod = toNumberRV(args[4]);
  if (isError(endPeriod)) {
    return endPeriod;
  }
  const type = toNumberRV(args[5]);
  if (isError(type)) {
    return type;
  }
  if (rate.value <= 0 || nper.value <= 0 || pv.value <= 0) {
    return ERRORS.NUM;
  }
  let cumPrinc = 0;
  for (let p = Math.floor(startPeriod.value); p <= Math.floor(endPeriod.value); p++) {
    const pmtVal = pmtRaw(rate.value, nper.value, pv.value, 0, type.value);
    const ipmtVal = ipmtRaw(rate.value, p, nper.value, pv.value, 0, type.value);
    cumPrinc += pmtVal - ipmtVal;
  }
  return rvNumber(cumPrinc);
};

export const fnCUMIPMT: NativeFunction = args => {
  const rate = toNumberRV(args[0]);
  if (isError(rate)) {
    return rate;
  }
  const nper = toNumberRV(args[1]);
  if (isError(nper)) {
    return nper;
  }
  const pv = toNumberRV(args[2]);
  if (isError(pv)) {
    return pv;
  }
  const startPeriod = toNumberRV(args[3]);
  if (isError(startPeriod)) {
    return startPeriod;
  }
  const endPeriod = toNumberRV(args[4]);
  if (isError(endPeriod)) {
    return endPeriod;
  }
  const type = toNumberRV(args[5]);
  if (isError(type)) {
    return type;
  }
  if (rate.value <= 0 || nper.value <= 0 || pv.value <= 0) {
    return ERRORS.NUM;
  }
  let cumIpmt = 0;
  for (let p = Math.floor(startPeriod.value); p <= Math.floor(endPeriod.value); p++) {
    cumIpmt += ipmtRaw(rate.value, p, nper.value, pv.value, 0, type.value);
  }
  return rvNumber(cumIpmt);
};

export const fnDOLLARDE: NativeFunction = args => {
  const fractionalDollar = toNumberRV(args[0]);
  if (isError(fractionalDollar)) {
    return fractionalDollar;
  }
  const fraction = toNumberRV(args[1]);
  if (isError(fraction)) {
    return fraction;
  }
  if (fraction.value < 1) {
    return ERRORS.NUM;
  }
  const f = Math.floor(fraction.value);
  const intPart = Math.trunc(fractionalDollar.value);
  const fracPart = fractionalDollar.value - intPart;
  return rvNumber(intPart + (fracPart / f) * Math.pow(10, Math.ceil(Math.log10(f))));
};

export const fnDOLLARFR: NativeFunction = args => {
  const decimalDollar = toNumberRV(args[0]);
  if (isError(decimalDollar)) {
    return decimalDollar;
  }
  const fraction = toNumberRV(args[1]);
  if (isError(fraction)) {
    return fraction;
  }
  if (fraction.value < 1) {
    return ERRORS.NUM;
  }
  const f = Math.floor(fraction.value);
  const intPart = Math.trunc(decimalDollar.value);
  const fracPart = decimalDollar.value - intPart;
  return rvNumber(intPart + (fracPart * f) / Math.pow(10, Math.ceil(Math.log10(f))));
};

export const fnDISC: NativeFunction = args => {
  const settlement = toNumberRV(args[0]);
  if (isError(settlement)) {
    return settlement;
  }
  const maturity = toNumberRV(args[1]);
  if (isError(maturity)) {
    return maturity;
  }
  const pr = toNumberRV(args[2]);
  if (isError(pr)) {
    return pr;
  }
  const redemption = toNumberRV(args[3]);
  if (isError(redemption)) {
    return redemption;
  }
  const basis = args.length > 4 ? toNumberRV(args[4]) : rvNumber(0);
  if (isError(basis)) {
    return basis;
  }
  const days = Math.floor(maturity.value) - Math.floor(settlement.value);
  if (days <= 0 || redemption.value <= 0) {
    return ERRORS.NUM;
  }
  const yearDays = basis.value === 1 ? 365.25 : basis.value === 3 ? 365 : 360;
  return rvNumber(((redemption.value - pr.value) / redemption.value) * (yearDays / days));
};

export const fnPRICEDISC: NativeFunction = args => {
  const settlement = toNumberRV(args[0]);
  if (isError(settlement)) {
    return settlement;
  }
  const maturity = toNumberRV(args[1]);
  if (isError(maturity)) {
    return maturity;
  }
  const disc = toNumberRV(args[2]);
  if (isError(disc)) {
    return disc;
  }
  const redemption = toNumberRV(args[3]);
  if (isError(redemption)) {
    return redemption;
  }
  const basis = args.length > 4 ? toNumberRV(args[4]) : rvNumber(0);
  if (isError(basis)) {
    return basis;
  }
  const days = Math.floor(maturity.value) - Math.floor(settlement.value);
  const yearDays = basis.value === 1 ? 365.25 : basis.value === 3 ? 365 : 360;
  return rvNumber(redemption.value - disc.value * redemption.value * (days / yearDays));
};

export const fnYIELDDISC: NativeFunction = args => {
  const settlement = toNumberRV(args[0]);
  if (isError(settlement)) {
    return settlement;
  }
  const maturity = toNumberRV(args[1]);
  if (isError(maturity)) {
    return maturity;
  }
  const pr = toNumberRV(args[2]);
  if (isError(pr)) {
    return pr;
  }
  const redemption = toNumberRV(args[3]);
  if (isError(redemption)) {
    return redemption;
  }
  const basis = args.length > 4 ? toNumberRV(args[4]) : rvNumber(0);
  if (isError(basis)) {
    return basis;
  }
  const days = Math.floor(maturity.value) - Math.floor(settlement.value);
  if (days <= 0 || pr.value <= 0) {
    return ERRORS.NUM;
  }
  const yearDays = basis.value === 1 ? 365.25 : basis.value === 3 ? 365 : 360;
  return rvNumber(((redemption.value - pr.value) / pr.value) * (yearDays / days));
};

export const fnRECEIVED: NativeFunction = args => {
  const settlement = toNumberRV(args[0]);
  if (isError(settlement)) {
    return settlement;
  }
  const maturity = toNumberRV(args[1]);
  if (isError(maturity)) {
    return maturity;
  }
  const investment = toNumberRV(args[2]);
  if (isError(investment)) {
    return investment;
  }
  const disc = toNumberRV(args[3]);
  if (isError(disc)) {
    return disc;
  }
  const basis = args.length > 4 ? toNumberRV(args[4]) : rvNumber(0);
  if (isError(basis)) {
    return basis;
  }
  const days = Math.floor(maturity.value) - Math.floor(settlement.value);
  const yearDays = basis.value === 1 ? 365.25 : basis.value === 3 ? 365 : 360;
  const denom = 1 - disc.value * (days / yearDays);
  if (denom === 0) {
    return ERRORS.NUM;
  }
  return rvNumber(investment.value / denom);
};

export const fnINTRATE: NativeFunction = args => {
  const settlement = toNumberRV(args[0]);
  if (isError(settlement)) {
    return settlement;
  }
  const maturity = toNumberRV(args[1]);
  if (isError(maturity)) {
    return maturity;
  }
  const investment = toNumberRV(args[2]);
  if (isError(investment)) {
    return investment;
  }
  const redemption = toNumberRV(args[3]);
  if (isError(redemption)) {
    return redemption;
  }
  const basis = args.length > 4 ? toNumberRV(args[4]) : rvNumber(0);
  if (isError(basis)) {
    return basis;
  }
  const days = Math.floor(maturity.value) - Math.floor(settlement.value);
  if (days <= 0 || investment.value <= 0) {
    return ERRORS.NUM;
  }
  const yearDays = basis.value === 1 ? 365.25 : basis.value === 3 ? 365 : 360;
  return rvNumber(((redemption.value - investment.value) / investment.value) * (yearDays / days));
};

// ============================================================================
// Bond Math — day-count conventions and coupon helpers
// ============================================================================

/**
 * Day-count fraction between two Excel date serials under a given basis.
 *
 * Implements:
 *   basis 0 — US (NASD) 30/360 with end-of-month adjustments.
 *   basis 1 — Actual/Actual (year length from the spanning year).
 *   basis 2 — Actual/360.
 *   basis 3 — Actual/365.
 *   basis 4 — European 30/360.
 *
 * Unknown basis values fall back to basis 0 for safety.
 */
function dayCountFraction(startSerial: number, endSerial: number, basis: number): number {
  const startD = excelToDate(startSerial);
  const endD = excelToDate(endSerial);
  const diffDays = Math.floor(endSerial) - Math.floor(startSerial);

  switch (basis) {
    case 1: {
      const y1 = startD.getFullYear();
      const y2 = endD.getFullYear();
      if (y1 === y2) {
        const yearDays =
          (new Date(y1 + 1, 0, 1).getTime() - new Date(y1, 0, 1).getTime()) / 86400000;
        return diffDays / yearDays;
      }
      const totalYearDays =
        (new Date(y2 + 1, 0, 1).getTime() - new Date(y1, 0, 1).getTime()) / 86400000;
      const avgYear = totalYearDays / (y2 - y1 + 1);
      return diffDays / avgYear;
    }
    case 2:
      return diffDays / 360;
    case 3:
      return diffDays / 365;
    case 4: {
      const d1 = Math.min(startD.getDate(), 30);
      const d2 = Math.min(endD.getDate(), 30);
      const m1 = startD.getMonth() + 1;
      const m2 = endD.getMonth() + 1;
      const y1 = startD.getFullYear();
      const y2 = endD.getFullYear();
      return ((y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)) / 360;
    }
    case 0:
    default: {
      let d1 = startD.getDate();
      const m1 = startD.getMonth() + 1;
      const y1 = startD.getFullYear();
      let d2 = endD.getDate();
      const m2 = endD.getMonth() + 1;
      const y2 = endD.getFullYear();
      if (d1 === 31) {
        d1 = 30;
      }
      if (d2 === 31 && d1 >= 30) {
        d2 = 30;
      }
      return ((y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)) / 360;
    }
  }
}

/**
 * Subtract `months` from an Excel date serial, clamping to month-end as needed.
 */
function addMonthsToSerial(serial: number, months: number): number {
  const d = excelToDate(serial);
  const y = d.getFullYear();
  const m = d.getMonth();
  const day = d.getDate();
  const target = new Date(y, m + months, 1);
  const lastDay = new Date(target.getFullYear(), target.getMonth() + 1, 0).getDate();
  target.setDate(Math.min(day, lastDay));
  // Date at local midnight → Excel serial (1900 epoch with leap bug handled in excelToDate).
  // We round-trip via the difference in ms from a known reference (Excel serial 0 = 1899-12-30).
  const epoch = new Date(1899, 11, 30).getTime();
  return Math.round((target.getTime() - epoch) / 86400000);
}

/**
 * Find the next coupon date on or after settlement, given a maturity date
 * and coupon frequency (1=annual, 2=semi-annual, 4=quarterly).
 *
 * Computed by stepping backward from maturity by (12/frequency) months
 * until the resulting date is ≤ settlement, then stepping one coupon forward.
 */
function nextCouponAfter(settlement: number, maturity: number, frequency: number): number {
  const stepMonths = Math.round(12 / frequency);
  let cur = maturity;
  // Step backward until just before or at settlement.
  while (cur > settlement) {
    const prev = addMonthsToSerial(cur, -stepMonths);
    if (prev <= settlement) {
      return cur;
    }
    cur = prev;
  }
  return cur;
}

/**
 * Find the previous coupon date (on or before settlement).
 */
function prevCouponOnOrBefore(settlement: number, maturity: number, frequency: number): number {
  const stepMonths = Math.round(12 / frequency);
  let cur = maturity;
  while (cur > settlement) {
    cur = addMonthsToSerial(cur, -stepMonths);
  }
  return cur;
}

/**
 * Count coupon periods between settlement and maturity (rounded up).
 */
function couponsBetween(settlement: number, maturity: number, frequency: number): number {
  const stepMonths = Math.round(12 / frequency);
  let count = 0;
  let cur = maturity;
  while (cur > settlement) {
    cur = addMonthsToSerial(cur, -stepMonths);
    count++;
  }
  return count;
}

/**
 * Validate the basic bond arguments (frequency and basis). Returns null if
 * valid, or the appropriate #NUM! error.
 */
function validateBondBasis(frequency: number, basis: number): ErrorValue | null {
  if (frequency !== 1 && frequency !== 2 && frequency !== 4) {
    return ERRORS.NUM;
  }
  if (basis < 0 || basis > 4) {
    return ERRORS.NUM;
  }
  return null;
}

/**
 * PRICE(settlement, maturity, rate, yield, redemption, frequency, [basis])
 *
 * Price per $100 face value of a security that pays periodic interest.
 * Excel formula (standard case, more than one coupon period remaining):
 *
 *   P = [redemption / (1 + y/f)^(N - 1 + DSC/E)]
 *     + Σ_{k=1..N} [100 * r / f / (1 + y/f)^(k - 1 + DSC/E)]
 *     - 100 * r / f * A/E
 *
 * Where:
 *   f  = frequency
 *   N  = number of coupons from settlement to maturity
 *   DSC = days from settlement to next coupon
 *   E   = days in coupon period containing settlement
 *   A   = days from beginning of coupon period to settlement
 */
export const fnPRICE: NativeFunction = args => {
  const settlementRV = toNumberRV(args[0]);
  if (isError(settlementRV)) {
    return settlementRV;
  }
  const maturityRV = toNumberRV(args[1]);
  if (isError(maturityRV)) {
    return maturityRV;
  }
  const rateRV = toNumberRV(args[2]);
  if (isError(rateRV)) {
    return rateRV;
  }
  const yieldRV = toNumberRV(args[3]);
  if (isError(yieldRV)) {
    return yieldRV;
  }
  const redemptionRV = toNumberRV(args[4]);
  if (isError(redemptionRV)) {
    return redemptionRV;
  }
  const frequencyRV = toNumberRV(args[5]);
  if (isError(frequencyRV)) {
    return frequencyRV;
  }
  const basisRV = args.length > 6 ? toNumberRV(args[6]) : rvNumber(0);
  if (isError(basisRV)) {
    return basisRV;
  }

  const settlement = Math.floor(settlementRV.value);
  const maturity = Math.floor(maturityRV.value);
  const rate = rateRV.value;
  const yld = yieldRV.value;
  const redemption = redemptionRV.value;
  const frequency = Math.floor(frequencyRV.value);
  const basis = Math.floor(basisRV.value);

  if (settlement >= maturity || rate < 0 || yld < 0 || redemption <= 0) {
    return ERRORS.NUM;
  }
  const basisErr = validateBondBasis(frequency, basis);
  if (basisErr) {
    return basisErr;
  }

  const nextCoupon = nextCouponAfter(settlement, maturity, frequency);
  const prevCoupon = prevCouponOnOrBefore(settlement, maturity, frequency);
  const dsc = dayCountFraction(settlement, nextCoupon, basis) * (basis === 1 ? 1 : 360 / frequency);
  // For basis 0/2/3/4 the "E" (period length) and "A" (accrued days) are
  // measured in actual days; for basis 1 we use actual day counts.
  const periodDays = Math.floor(nextCoupon) - Math.floor(prevCoupon);
  const accruedDays = settlement - Math.floor(prevCoupon);
  const e = periodDays;
  const dscDays = Math.floor(nextCoupon) - settlement;
  const a = accruedDays;
  // Fractional position in period (dsc/E).
  const dscE = e === 0 ? 0 : dscDays / e;
  void dsc;

  const N = couponsBetween(settlement, maturity, frequency);
  const couponAmt = (100 * rate) / frequency;
  const discountBase = 1 + yld / frequency;

  let price = redemption / Math.pow(discountBase, N - 1 + dscE);
  for (let k = 1; k <= N; k++) {
    price += couponAmt / Math.pow(discountBase, k - 1 + dscE);
  }
  price -= couponAmt * (a / e);
  return rvNumber(price);
};

/**
 * YIELD(settlement, maturity, rate, pr, redemption, frequency, [basis])
 *
 * Inverse of PRICE: solve numerically for yield such that PRICE(...y) = pr.
 * Uses bracketed bisection in [0, 1] (100% yield upper bound covers all
 * realistic bond scenarios) followed by a light Newton polish.
 */
export const fnYIELD: NativeFunction = args => {
  const settlementRV = toNumberRV(args[0]);
  if (isError(settlementRV)) {
    return settlementRV;
  }
  const maturityRV = toNumberRV(args[1]);
  if (isError(maturityRV)) {
    return maturityRV;
  }
  const rateRV = toNumberRV(args[2]);
  if (isError(rateRV)) {
    return rateRV;
  }
  const prRV = toNumberRV(args[3]);
  if (isError(prRV)) {
    return prRV;
  }
  const redemptionRV = toNumberRV(args[4]);
  if (isError(redemptionRV)) {
    return redemptionRV;
  }
  const frequencyRV = toNumberRV(args[5]);
  if (isError(frequencyRV)) {
    return frequencyRV;
  }
  const basisRV = args.length > 6 ? toNumberRV(args[6]) : rvNumber(0);
  if (isError(basisRV)) {
    return basisRV;
  }

  if (prRV.value <= 0 || redemptionRV.value <= 0) {
    return ERRORS.NUM;
  }

  const priceAt = (y: number): number => {
    const result = fnPRICE([
      settlementRV,
      maturityRV,
      rateRV,
      rvNumber(y),
      redemptionRV,
      frequencyRV,
      basisRV
    ]);
    if (result.kind !== RVKind.Number) {
      return NaN;
    }
    return result.value;
  };

  // Bisection in [0, 1].
  let lo = 0;
  let hi = 1;
  const fLo = priceAt(lo) - prRV.value;
  const fHi = priceAt(hi) - prRV.value;
  if (isNaN(fLo) || isNaN(fHi)) {
    return ERRORS.NUM;
  }
  if (fLo * fHi > 0) {
    // Try extending upper bound.
    hi = 10;
    if ((priceAt(hi) - prRV.value) * fLo > 0) {
      return ERRORS.NUM;
    }
  }
  for (let i = 0; i < 200; i++) {
    const mid = (lo + hi) / 2;
    const f = priceAt(mid) - prRV.value;
    if (isNaN(f)) {
      return ERRORS.NUM;
    }
    // PRICE decreases as yield increases, so f = priceAtY - pr decreases.
    if (f > 0) {
      lo = mid;
    } else {
      hi = mid;
    }
    if (hi - lo < 1e-12) {
      break;
    }
  }
  return rvNumber((lo + hi) / 2);
};

/**
 * DURATION(settlement, maturity, coupon, yield, frequency, [basis])
 *
 * Macaulay duration of a bond: the weighted average time to cash flows,
 * weighted by present value. Expressed in years.
 */
export const fnDURATION: NativeFunction = args => {
  const settlementRV = toNumberRV(args[0]);
  if (isError(settlementRV)) {
    return settlementRV;
  }
  const maturityRV = toNumberRV(args[1]);
  if (isError(maturityRV)) {
    return maturityRV;
  }
  const couponRV = toNumberRV(args[2]);
  if (isError(couponRV)) {
    return couponRV;
  }
  const yieldRV = toNumberRV(args[3]);
  if (isError(yieldRV)) {
    return yieldRV;
  }
  const frequencyRV = toNumberRV(args[4]);
  if (isError(frequencyRV)) {
    return frequencyRV;
  }
  const basisRV = args.length > 5 ? toNumberRV(args[5]) : rvNumber(0);
  if (isError(basisRV)) {
    return basisRV;
  }

  const settlement = Math.floor(settlementRV.value);
  const maturity = Math.floor(maturityRV.value);
  const coupon = couponRV.value;
  const yld = yieldRV.value;
  const frequency = Math.floor(frequencyRV.value);
  const basis = Math.floor(basisRV.value);

  if (settlement >= maturity || coupon < 0 || yld < 0) {
    return ERRORS.NUM;
  }
  const basisErr = validateBondBasis(frequency, basis);
  if (basisErr) {
    return basisErr;
  }

  const nextCoupon = nextCouponAfter(settlement, maturity, frequency);
  const prevCoupon = prevCouponOnOrBefore(settlement, maturity, frequency);
  const periodDays = Math.floor(nextCoupon) - Math.floor(prevCoupon);
  const dscDays = Math.floor(nextCoupon) - settlement;
  const dscE = periodDays === 0 ? 0 : dscDays / periodDays;
  const N = couponsBetween(settlement, maturity, frequency);
  const couponPerPeriod = (100 * coupon) / frequency;
  const discountBase = 1 + yld / frequency;

  let pv = 0;
  let weighted = 0;
  for (let k = 1; k <= N; k++) {
    const t = (k - 1 + dscE) / frequency; // time in years
    const cf = k === N ? couponPerPeriod + 100 : couponPerPeriod;
    const df = Math.pow(discountBase, k - 1 + dscE);
    pv += cf / df;
    weighted += (t * cf) / df;
  }
  if (pv === 0) {
    return ERRORS.NUM;
  }
  return rvNumber(weighted / pv);
};

/**
 * MDURATION — modified duration = DURATION / (1 + yield/frequency).
 */
export const fnMDURATION: NativeFunction = args => {
  const dur = fnDURATION(args);
  if (dur.kind !== RVKind.Number) {
    return dur;
  }
  const yieldRV = toNumberRV(args[3]);
  if (isError(yieldRV)) {
    return yieldRV;
  }
  const frequencyRV = toNumberRV(args[4]);
  if (isError(frequencyRV)) {
    return frequencyRV;
  }
  return rvNumber(dur.value / (1 + yieldRV.value / frequencyRV.value));
};

/**
 * ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method])
 *
 * Accrued interest for a security that pays periodic interest.
 * The simplified implementation (calc_method TRUE, the default) treats
 * accrued interest from issue to settlement as par * rate * dcf(issue, settlement, basis).
 */
export const fnACCRINT: NativeFunction = args => {
  const issueRV = toNumberRV(args[0]);
  if (isError(issueRV)) {
    return issueRV;
  }
  // first_interest (args[1]) is unused in the simplified implementation.
  const settlementRV = toNumberRV(args[2]);
  if (isError(settlementRV)) {
    return settlementRV;
  }
  const rateRV = toNumberRV(args[3]);
  if (isError(rateRV)) {
    return rateRV;
  }
  const parRV = toNumberRV(args[4]);
  if (isError(parRV)) {
    return parRV;
  }
  const frequencyRV = toNumberRV(args[5]);
  if (isError(frequencyRV)) {
    return frequencyRV;
  }
  const basisRV = args.length > 6 ? toNumberRV(args[6]) : rvNumber(0);
  if (isError(basisRV)) {
    return basisRV;
  }
  // calc_method (args[7]) — we ignore the distinction between TRUE/FALSE
  // because the simplified semantics always accrue from issue date.

  const issue = Math.floor(issueRV.value);
  const settlement = Math.floor(settlementRV.value);
  const frequency = Math.floor(frequencyRV.value);
  const basis = Math.floor(basisRV.value);

  if (issue >= settlement || rateRV.value <= 0 || parRV.value <= 0) {
    return ERRORS.NUM;
  }
  const basisErr = validateBondBasis(frequency, basis);
  if (basisErr) {
    return basisErr;
  }

  const dcf = dayCountFraction(issue, settlement, basis);
  return rvNumber(parRV.value * rateRV.value * dcf);
};
