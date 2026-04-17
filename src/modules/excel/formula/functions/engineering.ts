/**
 * Engineering Functions — Native RuntimeValue implementation.
 */

import type { RuntimeValue, ErrorValue } from "../runtime/values";
import {
  RVKind,
  ERRORS,
  isError,
  toNumberRV,
  toStringRV,
  topLeft,
  rvNumber,
  rvString
} from "../runtime/values";

// ============================================================================
// Type alias for native function signature
// ============================================================================

type NativeFunction = (args: RuntimeValue[]) => RuntimeValue;

// ============================================================================
// Local utility
// ============================================================================

function checkError(v: RuntimeValue): ErrorValue | null {
  const s = topLeft(v);
  return s.kind === RVKind.Error ? s : null;
}

// ============================================================================
// Base Conversion Functions
// ============================================================================

export const fnBIN2DEC: NativeFunction = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const s = toStringRV(args[0]);
  if (!/^[01]{1,10}$/.test(s)) {
    return ERRORS.NUM;
  }
  // 10-bit two's complement
  if (s.length === 10 && s[0] === "1") {
    return rvNumber(parseInt(s.slice(1), 2) - 512);
  }
  return rvNumber(parseInt(s, 2));
};

export const fnDEC2BIN: NativeFunction = args => {
  const nRV = toNumberRV(args[0]);
  if (isError(nRV)) {
    return nRV;
  }
  const n = Math.floor(nRV.value);
  if (n < -512 || n > 511) {
    return ERRORS.NUM;
  }
  const placesRV = args.length > 1 ? toNumberRV(args[1]) : rvNumber(0);
  if (isError(placesRV)) {
    return placesRV;
  }
  const places = placesRV.value;
  if (n < 0) {
    return rvString((n + 1024).toString(2));
  }
  const result = n.toString(2);
  return rvString(places > 0 ? result.padStart(places, "0") : result);
};

export const fnHEX2DEC: NativeFunction = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const s = toStringRV(args[0]);
  if (!/^[0-9A-Fa-f]{1,10}$/.test(s)) {
    return ERRORS.NUM;
  }
  const num = parseInt(s, 16);
  // 10-digit hex: 40-bit two's complement
  if (s.length === 10 && parseInt(s[0], 16) >= 8) {
    return rvNumber(num - Math.pow(16, 10));
  }
  return rvNumber(num);
};

export const fnDEC2HEX: NativeFunction = args => {
  const nRV = toNumberRV(args[0]);
  if (isError(nRV)) {
    return nRV;
  }
  const n = Math.floor(nRV.value);
  const placesRV = args.length > 1 ? toNumberRV(args[1]) : rvNumber(0);
  if (isError(placesRV)) {
    return placesRV;
  }
  const places = placesRV.value;
  if (n < 0) {
    return rvString((n + Math.pow(16, 10)).toString(16).toUpperCase());
  }
  const result = n.toString(16).toUpperCase();
  return rvString(places > 0 ? result.padStart(places, "0") : result);
};

export const fnOCT2DEC: NativeFunction = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const s = toStringRV(args[0]);
  if (!/^[0-7]{1,10}$/.test(s)) {
    return ERRORS.NUM;
  }
  const num = parseInt(s, 8);
  if (s.length === 10 && parseInt(s[0]) >= 4) {
    return rvNumber(num - Math.pow(8, 10));
  }
  return rvNumber(num);
};

export const fnDEC2OCT: NativeFunction = args => {
  const nRV = toNumberRV(args[0]);
  if (isError(nRV)) {
    return nRV;
  }
  const n = Math.floor(nRV.value);
  const placesRV = args.length > 1 ? toNumberRV(args[1]) : rvNumber(0);
  if (isError(placesRV)) {
    return placesRV;
  }
  const places = placesRV.value;
  if (n < 0) {
    return rvString((n + Math.pow(8, 10)).toString(8));
  }
  const result = n.toString(8);
  return rvString(places > 0 ? result.padStart(places, "0") : result);
};

export const fnDELTA: NativeFunction = args => {
  const n1 = toNumberRV(args[0]);
  if (isError(n1)) {
    return n1;
  }
  const n2 = args.length > 1 ? toNumberRV(args[1]) : rvNumber(0);
  if (isError(n2)) {
    return n2;
  }
  return rvNumber(n1.value === n2.value ? 1 : 0);
};

export const fnGESTEP: NativeFunction = args => {
  const n = toNumberRV(args[0]);
  if (isError(n)) {
    return n;
  }
  const step = args.length > 1 ? toNumberRV(args[1]) : rvNumber(0);
  if (isError(step)) {
    return step;
  }
  return rvNumber(n.value >= step.value ? 1 : 0);
};

// ============================================================================
// Complex Numbers, Bit Operations
// ============================================================================

function parseComplex(s: string): [number, number] | null {
  const text = s.trim();
  if (text === "") {
    return null;
  }
  const pureReal = Number(text);
  if (!isNaN(pureReal) && !text.endsWith("i") && !text.endsWith("j")) {
    return [pureReal, 0];
  }
  if (text === "i" || text === "j") {
    return [0, 1];
  }
  if (text === "-i" || text === "-j") {
    return [0, -1];
  }
  if ((text.endsWith("i") || text.endsWith("j")) && !isNaN(Number(text.slice(0, -1)))) {
    return [0, Number(text.slice(0, -1))];
  }
  const re = /^([+-]?[\d.eE+-]*?)([+-][\d.eE]*)?[ij]$/;
  const m = re.exec(text);
  if (!m) {
    return null;
  }
  const realPart = m[1] === "" || m[1] === "+" ? 0 : m[1] === "-" ? 0 : Number(m[1]);
  let imagPart = m[2] === undefined ? 0 : m[2] === "+" ? 1 : m[2] === "-" ? -1 : Number(m[2]);
  if (m[2] === undefined && m[1] !== "") {
    imagPart = m[1] === "" || m[1] === "+" ? 1 : m[1] === "-" ? -1 : Number(m[1]);
    return [0, imagPart];
  }
  return [isNaN(realPart) ? 0 : realPart, isNaN(imagPart) ? 0 : imagPart];
}

function formatComplex(re: number, im: number, suffix: string = "i"): string {
  if (im === 0) {
    return String(re);
  }
  if (re === 0) {
    if (im === 1) {
      return suffix;
    }
    if (im === -1) {
      return "-" + suffix;
    }
    return im + suffix;
  }
  const imStr =
    im === 1 ? "+" + suffix : im === -1 ? "-" + suffix : (im > 0 ? "+" : "") + im + suffix;
  return re + imStr;
}

export const fnCOMPLEX: NativeFunction = args => {
  const re = toNumberRV(args[0]);
  if (isError(re)) {
    return re;
  }
  const im = toNumberRV(args[1]);
  if (isError(im)) {
    return im;
  }
  let suffix = "i";
  if (args.length > 2) {
    const e2 = checkError(args[2]);
    if (e2) {
      return e2;
    }
    suffix = toStringRV(args[2]);
  }
  if (suffix !== "i" && suffix !== "j") {
    return ERRORS.VALUE;
  }
  return rvString(formatComplex(re.value, im.value, suffix));
};

export const fnIMREAL: NativeFunction = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(args[0]));
  return c ? rvNumber(c[0]) : ERRORS.NUM;
};

export const fnIMAGINARY: NativeFunction = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(args[0]));
  return c ? rvNumber(c[1]) : ERRORS.NUM;
};

export const fnIMABS: NativeFunction = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(args[0]));
  if (!c) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.sqrt(c[0] * c[0] + c[1] * c[1]));
};

export const fnIMARGUMENT: NativeFunction = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(args[0]));
  if (!c) {
    return ERRORS.NUM;
  }
  if (c[0] === 0 && c[1] === 0) {
    return ERRORS.DIV0;
  }
  return rvNumber(Math.atan2(c[1], c[0]));
};

export const fnIMCONJUGATE: NativeFunction = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(args[0]));
  if (!c) {
    return ERRORS.NUM;
  }
  return rvString(formatComplex(c[0], -c[1]));
};

export const fnIMSUM: NativeFunction = args => {
  let re = 0,
    im = 0;
  for (const a of args) {
    const err = checkError(a);
    if (err) {
      return err;
    }
    const s =
      a.kind === RVKind.Array && a.height > 0 && a.width > 0
        ? toStringRV(a.rows[0][0])
        : toStringRV(a);
    const c = parseComplex(s);
    if (!c) {
      return ERRORS.NUM;
    }
    re += c[0];
    im += c[1];
  }
  return rvString(formatComplex(re, im));
};

export const fnIMSUB: NativeFunction = args => {
  const e0 = checkError(args[0]);
  if (e0) {
    return e0;
  }
  const e1 = checkError(args[1]);
  if (e1) {
    return e1;
  }
  const c1 = parseComplex(toStringRV(args[0]));
  const c2 = parseComplex(toStringRV(args[1]));
  if (!c1 || !c2) {
    return ERRORS.NUM;
  }
  return rvString(formatComplex(c1[0] - c2[0], c1[1] - c2[1]));
};

export const fnIMPRODUCT: NativeFunction = args => {
  let re = 1,
    im = 0;
  for (const a of args) {
    const err = checkError(a);
    if (err) {
      return err;
    }
    const s =
      a.kind === RVKind.Array && a.height > 0 && a.width > 0
        ? toStringRV(a.rows[0][0])
        : toStringRV(a);
    const c = parseComplex(s);
    if (!c) {
      return ERRORS.NUM;
    }
    const nRe = re * c[0] - im * c[1];
    const nIm = re * c[1] + im * c[0];
    re = nRe;
    im = nIm;
  }
  return rvString(formatComplex(re, im));
};

export const fnIMDIV: NativeFunction = args => {
  const e0 = checkError(args[0]);
  if (e0) {
    return e0;
  }
  const e1 = checkError(args[1]);
  if (e1) {
    return e1;
  }
  const c1 = parseComplex(toStringRV(args[0]));
  const c2 = parseComplex(toStringRV(args[1]));
  if (!c1 || !c2) {
    return ERRORS.NUM;
  }
  const d = c2[0] * c2[0] + c2[1] * c2[1];
  if (d === 0) {
    return ERRORS.NUM;
  }
  return rvString(
    formatComplex((c1[0] * c2[0] + c1[1] * c2[1]) / d, (c1[1] * c2[0] - c1[0] * c2[1]) / d)
  );
};

export const fnIMPOWER: NativeFunction = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(args[0]));
  if (!c) {
    return ERRORS.NUM;
  }
  const n = toNumberRV(args[1]);
  if (isError(n)) {
    return n;
  }
  const r = Math.sqrt(c[0] * c[0] + c[1] * c[1]);
  const theta = Math.atan2(c[1], c[0]);
  const rn = Math.pow(r, n.value);
  return rvString(formatComplex(rn * Math.cos(n.value * theta), rn * Math.sin(n.value * theta)));
};

export const fnIMSQRT: NativeFunction = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(args[0]));
  if (!c) {
    return ERRORS.NUM;
  }
  const r = Math.sqrt(c[0] * c[0] + c[1] * c[1]);
  const theta = Math.atan2(c[1], c[0]);
  const sr = Math.sqrt(r);
  return rvString(formatComplex(sr * Math.cos(theta / 2), sr * Math.sin(theta / 2)));
};

export const fnIMLN: NativeFunction = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(args[0]));
  if (!c) {
    return ERRORS.NUM;
  }
  const r = Math.sqrt(c[0] * c[0] + c[1] * c[1]);
  if (r === 0) {
    return ERRORS.NUM;
  }
  return rvString(formatComplex(Math.log(r), Math.atan2(c[1], c[0])));
};

export const fnIMLOG2: NativeFunction = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(args[0]));
  if (!c) {
    return ERRORS.NUM;
  }
  const r = Math.sqrt(c[0] * c[0] + c[1] * c[1]);
  if (r === 0) {
    return ERRORS.NUM;
  }
  const ln2 = Math.log(2);
  return rvString(formatComplex(Math.log(r) / ln2, Math.atan2(c[1], c[0]) / ln2));
};

export const fnIMLOG10: NativeFunction = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(args[0]));
  if (!c) {
    return ERRORS.NUM;
  }
  const r = Math.sqrt(c[0] * c[0] + c[1] * c[1]);
  if (r === 0) {
    return ERRORS.NUM;
  }
  const ln10 = Math.log(10);
  return rvString(formatComplex(Math.log(r) / ln10, Math.atan2(c[1], c[0]) / ln10));
};

export const fnIMEXP: NativeFunction = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(args[0]));
  if (!c) {
    return ERRORS.NUM;
  }
  const er = Math.exp(c[0]);
  return rvString(formatComplex(er * Math.cos(c[1]), er * Math.sin(c[1])));
};

export const fnIMSIN: NativeFunction = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(args[0]));
  if (!c) {
    return ERRORS.NUM;
  }
  return rvString(
    formatComplex(Math.sin(c[0]) * Math.cosh(c[1]), Math.cos(c[0]) * Math.sinh(c[1]))
  );
};

export const fnIMCOS: NativeFunction = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(args[0]));
  if (!c) {
    return ERRORS.NUM;
  }
  return rvString(
    formatComplex(Math.cos(c[0]) * Math.cosh(c[1]), -Math.sin(c[0]) * Math.sinh(c[1]))
  );
};

export const fnBITAND: NativeFunction = args => {
  const a = toNumberRV(args[0]);
  if (isError(a)) {
    return a;
  }
  const b = toNumberRV(args[1]);
  if (isError(b)) {
    return b;
  }
  if (a.value < 0 || b.value < 0) {
    return ERRORS.NUM;
  }
  return rvNumber((Math.floor(a.value) & Math.floor(b.value)) >>> 0);
};

export const fnBITOR: NativeFunction = args => {
  const a = toNumberRV(args[0]);
  if (isError(a)) {
    return a;
  }
  const b = toNumberRV(args[1]);
  if (isError(b)) {
    return b;
  }
  if (a.value < 0 || b.value < 0) {
    return ERRORS.NUM;
  }
  return rvNumber((Math.floor(a.value) | Math.floor(b.value)) >>> 0);
};

export const fnBITXOR: NativeFunction = args => {
  const a = toNumberRV(args[0]);
  if (isError(a)) {
    return a;
  }
  const b = toNumberRV(args[1]);
  if (isError(b)) {
    return b;
  }
  if (a.value < 0 || b.value < 0) {
    return ERRORS.NUM;
  }
  return rvNumber((Math.floor(a.value) ^ Math.floor(b.value)) >>> 0);
};

export const fnBITLSHIFT: NativeFunction = args => {
  const num = toNumberRV(args[0]);
  if (isError(num)) {
    return num;
  }
  const shift = toNumberRV(args[1]);
  if (isError(shift)) {
    return shift;
  }
  if (num.value < 0) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.floor(num.value) * Math.pow(2, Math.floor(shift.value)));
};

export const fnBITRSHIFT: NativeFunction = args => {
  const num = toNumberRV(args[0]);
  if (isError(num)) {
    return num;
  }
  const shift = toNumberRV(args[1]);
  if (isError(shift)) {
    return shift;
  }
  if (num.value < 0) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.floor(Math.floor(num.value) / Math.pow(2, Math.floor(shift.value))));
};
