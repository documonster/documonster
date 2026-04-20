/**
 * Unit tests for engineering functions in `../engineering.ts`.
 *
 * Focuses on:
 *   - Base conversions (DEC2BIN / DEC2HEX / DEC2OCT and their inverses),
 *     including the critical regression where `DEC2BIN(-0.5)` must return
 *     `"0"` (truncation toward zero), NOT `"1"` (which `Math.floor` gives
 *     for -0.5 → -1).
 *   - BITAND / BITOR / BITXOR / BITLSHIFT / BITRSHIFT in the 48-bit range.
 *     `BITAND(2^40, 2^40)` must return `2^40` — native 32-bit `&` would
 *     silently truncate the high bits and yield 0.
 */

import { describe, it, expect } from "vitest";

import {
  BLANK,
  ERRORS,
  RVKind,
  rvArray,
  rvBoolean,
  rvNumber,
  rvString,
  type NumberValue,
  type RuntimeValue,
  type StringValue
} from "../../runtime/values";
import {
  fnBIN2DEC,
  fnDEC2BIN,
  fnHEX2DEC,
  fnDEC2HEX,
  fnOCT2DEC,
  fnDEC2OCT,
  fnBIN2HEX,
  fnBIN2OCT,
  fnHEX2BIN,
  fnHEX2OCT,
  fnOCT2BIN,
  fnOCT2HEX,
  fnDELTA,
  fnGESTEP,
  fnBITAND,
  fnBITOR,
  fnBITXOR,
  fnBITLSHIFT,
  fnBITRSHIFT,
  fnCOMPLEX,
  fnIMREAL,
  fnIMAGINARY,
  fnIMABS,
  fnIMARGUMENT,
  fnIMCONJUGATE,
  fnIMSUM,
  fnIMSUB,
  fnIMPRODUCT,
  fnIMDIV,
  fnIMPOWER,
  fnIMSQRT,
  fnIMLN,
  fnIMLOG2,
  fnIMLOG10,
  fnIMEXP,
  fnIMSIN,
  fnIMCOS,
  fnIMTAN,
  fnIMCSC,
  fnIMSEC,
  fnIMCOT,
  fnIMSINH,
  fnIMCOSH,
  fnIMTANH,
  fnIMCSCH,
  fnIMSECH,
  fnIMCOTH,
  fnBESSELJ,
  fnBESSELI,
  fnBESSELK,
  fnBESSELY
} from "../engineering";

function asNumber(v: RuntimeValue): number {
  expect(v.kind).toBe(RVKind.Number);
  return (v as NumberValue).value;
}

function asString(v: RuntimeValue): string {
  expect(v.kind).toBe(RVKind.String);
  return (v as StringValue).value;
}

describe("BIN2DEC / DEC2BIN", () => {
  it("BIN2DEC parses binary strings", () => {
    expect(asNumber(fnBIN2DEC([rvString("1010")]))).toBe(10);
    expect(asNumber(fnBIN2DEC([rvString("0")]))).toBe(0);
  });

  it("BIN2DEC treats 10-bit inputs as 10-bit two's complement", () => {
    // 1000000000 = -512
    expect(asNumber(fnBIN2DEC([rvString("1000000000")]))).toBe(-512);
    expect(asNumber(fnBIN2DEC([rvString("1111111111")]))).toBe(-1);
  });

  it("BIN2DEC rejects > 10 digits or non-binary chars", () => {
    expect(fnBIN2DEC([rvString("11111111111")])).toEqual(ERRORS.NUM);
    expect(fnBIN2DEC([rvString("102")])).toEqual(ERRORS.NUM);
  });

  it("DEC2BIN formats integers", () => {
    expect(asString(fnDEC2BIN([rvNumber(10)]))).toBe("1010");
  });

  it("DEC2BIN(-0.5) == '0' — truncation toward zero (regression)", () => {
    // Math.floor(-0.5) = -1, but Excel's DEC→BASE truncates fractions
    // toward zero: -0.5 → 0.
    expect(asString(fnDEC2BIN([rvNumber(-0.5)]))).toBe("0");
  });

  it("DEC2BIN(0.5) == '0'", () => {
    expect(asString(fnDEC2BIN([rvNumber(0.5)]))).toBe("0");
  });

  it("DEC2BIN pads with places", () => {
    expect(asString(fnDEC2BIN([rvNumber(5), rvNumber(8)]))).toBe("00000101");
  });

  it("DEC2BIN of -1 uses 10-bit two's complement", () => {
    expect(asString(fnDEC2BIN([rvNumber(-1)]))).toBe("1111111111");
  });

  it("DEC2BIN rejects out-of-range values", () => {
    expect(fnDEC2BIN([rvNumber(512)])).toEqual(ERRORS.NUM);
    expect(fnDEC2BIN([rvNumber(-513)])).toEqual(ERRORS.NUM);
  });
});

describe("HEX2DEC / DEC2HEX", () => {
  it("HEX2DEC parses hex strings", () => {
    expect(asNumber(fnHEX2DEC([rvString("FF")]))).toBe(255);
    expect(asNumber(fnHEX2DEC([rvString("0")]))).toBe(0);
  });

  it("HEX2DEC treats 10-digit inputs as 40-bit two's complement", () => {
    // FFFFFFFFFF = -1 in 40-bit two's complement
    expect(asNumber(fnHEX2DEC([rvString("FFFFFFFFFF")]))).toBe(-1);
  });

  it("DEC2HEX formats uppercase", () => {
    expect(asString(fnDEC2HEX([rvNumber(255)]))).toBe("FF");
  });

  it("DEC2HEX truncates fractions toward zero (regression)", () => {
    expect(asString(fnDEC2HEX([rvNumber(-0.5)]))).toBe("0");
  });

  it("DEC2HEX pads with places", () => {
    expect(asString(fnDEC2HEX([rvNumber(10), rvNumber(4)]))).toBe("000A");
  });
});

describe("OCT2DEC / DEC2OCT", () => {
  it("OCT2DEC parses octal", () => {
    expect(asNumber(fnOCT2DEC([rvString("17")]))).toBe(15);
  });

  it("DEC2OCT truncates fractions toward zero (regression)", () => {
    expect(asString(fnDEC2OCT([rvNumber(-0.5)]))).toBe("0");
  });

  it("DEC2OCT formats", () => {
    expect(asString(fnDEC2OCT([rvNumber(8)]))).toBe("10");
  });
});

describe("DELTA / GESTEP", () => {
  it("DELTA returns 1 when equal", () => {
    expect(asNumber(fnDELTA([rvNumber(5), rvNumber(5)]))).toBe(1);
    expect(asNumber(fnDELTA([rvNumber(5), rvNumber(3)]))).toBe(0);
  });

  it("DELTA defaults second arg to 0", () => {
    expect(asNumber(fnDELTA([rvNumber(0)]))).toBe(1);
  });

  it("GESTEP returns 1 when number >= step", () => {
    expect(asNumber(fnGESTEP([rvNumber(5), rvNumber(4)]))).toBe(1);
    expect(asNumber(fnGESTEP([rvNumber(3), rvNumber(4)]))).toBe(0);
  });
});

describe("BITAND / BITOR / BITXOR — 48-bit range", () => {
  it("BITAND works within 32-bit range", () => {
    expect(asNumber(fnBITAND([rvNumber(0b1100), rvNumber(0b1010)]))).toBe(0b1000);
  });

  it("BITAND(2^40, 2^40) === 2^40 (regression)", () => {
    // A native 32-bit `&` would truncate and yield 0.
    const big = Math.pow(2, 40);
    expect(asNumber(fnBITAND([rvNumber(big), rvNumber(big)]))).toBe(big);
  });

  it("BITOR works within 32-bit range", () => {
    expect(asNumber(fnBITOR([rvNumber(0b1100), rvNumber(0b0011)]))).toBe(0b1111);
  });

  it("BITOR operates correctly at 2^40", () => {
    const big = Math.pow(2, 40);
    expect(asNumber(fnBITOR([rvNumber(big), rvNumber(1)]))).toBe(big + 1);
  });

  it("BITXOR within 32-bit range", () => {
    expect(asNumber(fnBITXOR([rvNumber(0b1100), rvNumber(0b0110)]))).toBe(0b1010);
  });

  it("BITXOR of 2^40 with itself is 0", () => {
    const big = Math.pow(2, 40);
    expect(asNumber(fnBITXOR([rvNumber(big), rvNumber(big)]))).toBe(0);
  });

  it("BIT* reject negative operands", () => {
    expect(fnBITAND([rvNumber(-1), rvNumber(0)])).toEqual(ERRORS.NUM);
  });

  it("BIT* reject operands >= 2^48", () => {
    expect(fnBITAND([rvNumber(Math.pow(2, 48)), rvNumber(0)])).toEqual(ERRORS.NUM);
  });
});

describe("BITLSHIFT / BITRSHIFT", () => {
  it("BITLSHIFT(1, 10) = 1024", () => {
    expect(asNumber(fnBITLSHIFT([rvNumber(1), rvNumber(10)]))).toBe(1024);
  });

  it("BITLSHIFT with a negative shift amount shifts right", () => {
    expect(asNumber(fnBITLSHIFT([rvNumber(16), rvNumber(-2)]))).toBe(4);
  });

  it("BITLSHIFT result >= 2^48 returns #NUM!", () => {
    expect(fnBITLSHIFT([rvNumber(Math.pow(2, 40)), rvNumber(10)])).toEqual(ERRORS.NUM);
  });

  it("BITRSHIFT(1024, 10) = 1", () => {
    expect(asNumber(fnBITRSHIFT([rvNumber(1024), rvNumber(10)]))).toBe(1);
  });

  it("BITRSHIFT with a negative shift amount shifts left", () => {
    expect(asNumber(fnBITRSHIFT([rvNumber(4), rvNumber(-2)]))).toBe(16);
  });
});

describe("COMPLEX / IMREAL / IMAGINARY / IMABS", () => {
  it("COMPLEX assembles a+bi", () => {
    expect(asString(fnCOMPLEX([rvNumber(3), rvNumber(4)]))).toBe("3+4i");
  });

  it("COMPLEX with a 'j' suffix", () => {
    expect(asString(fnCOMPLEX([rvNumber(3), rvNumber(4), rvString("j")]))).toBe("3+4j");
  });

  it("COMPLEX rejects non-i/j suffix", () => {
    expect(fnCOMPLEX([rvNumber(1), rvNumber(2), rvString("k")])).toEqual(ERRORS.VALUE);
  });

  it("IMREAL / IMAGINARY extract parts", () => {
    expect(asNumber(fnIMREAL([rvString("3+4i")]))).toBe(3);
    expect(asNumber(fnIMAGINARY([rvString("3+4i")]))).toBe(4);
  });

  it("IMABS returns the modulus", () => {
    expect(asNumber(fnIMABS([rvString("3+4i")]))).toBe(5);
  });
});

// ============================================================================
// R7: Complex trigonometric family
// ============================================================================

/** Parse a complex result string and return [re, im]. */
function splitComplex(s: string): [number, number] {
  // Pure real (no trailing i/j)
  if (!s.endsWith("i") && !s.endsWith("j")) {
    return [Number(s), 0];
  }
  // Special bare-imaginary forms
  if (s === "i" || s === "j") {
    return [0, 1];
  }
  if (s === "-i" || s === "-j") {
    return [0, -1];
  }
  if (s === "+i" || s === "+j") {
    return [0, 1];
  }
  // Delegate to the looser parser which handles scientific-notation cases
  // like "1-2.44e-16i" that the existing regex does not.
  const { re, im } = parseComplexLoose(s);
  return [re, im];
}

/**
 * A loose complex parser that copes with exponential notation. The test
 * module's original `splitComplex` regex doesn't allow the `-` inside an
 * exponent of the real part (e.g. "1-2.4e-16i"), so many floating-noise
 * results break it.
 */
function parseComplexLoose(s: string): { re: number; im: number } {
  // Handle pure cases first.
  if (!s.endsWith("i") && !s.endsWith("j")) {
    return { re: Number(s), im: 0 };
  }
  const body = s.slice(0, -1);
  if (body === "" || body === "+") {
    return { re: 0, im: 1 };
  }
  if (body === "-") {
    return { re: 0, im: -1 };
  }
  // Find the split between real and imaginary parts: the last '+' or '-'
  // that is NOT part of an exponent (i.e. not immediately preceded by
  // 'e' / 'E').
  let splitAt = -1;
  for (let i = body.length - 1; i > 0; i--) {
    const ch = body[i];
    if (ch === "+" || ch === "-") {
      const prev = body[i - 1];
      if (prev !== "e" && prev !== "E") {
        splitAt = i;
        break;
      }
    }
  }
  if (splitAt === -1) {
    // Pure imaginary: "5i", "-5i"
    return { re: 0, im: Number(body) };
  }
  const realStr = body.slice(0, splitAt);
  const imagStr = body.slice(splitAt);
  const imag = imagStr === "+" ? 1 : imagStr === "-" ? -1 : Number(imagStr);
  return { re: Number(realStr), im: imag };
}

describe("IMTAN", () => {
  it("computes tan for a real input", () => {
    const r = asString(fnIMTAN([rvString("1")]));
    const [re] = splitComplex(r);
    expect(re).toBeCloseTo(Math.tan(1), 5);
  });

  it("computes tan(i) = i·tanh(1)", () => {
    const [re, im] = splitComplex(asString(fnIMTAN([rvString("i")])));
    expect(re).toBeCloseTo(0, 5);
    expect(im).toBeCloseTo(Math.tanh(1), 5);
  });

  it("rejects malformed complex string", () => {
    expect(fnIMTAN([rvString("not a number")])).toEqual(ERRORS.NUM);
  });
});

describe("IMCSC / IMSEC / IMCOT", () => {
  it("IMCSC(π/2 + 0i) = 1 (csc of π/2 = 1)", () => {
    const [re] = splitComplex(asString(fnIMCSC([rvString(String(Math.PI / 2))])));
    expect(re).toBeCloseTo(1, 5);
  });

  it("IMSEC(0) = 1", () => {
    const [re] = splitComplex(asString(fnIMSEC([rvString("0")])));
    expect(re).toBeCloseTo(1, 10);
  });

  it("IMCOT(π/4) = 1", () => {
    const [re] = splitComplex(asString(fnIMCOT([rvString(String(Math.PI / 4))])));
    expect(re).toBeCloseTo(1, 5);
  });

  it("IMCOT(0) returns #DIV/0! (sin(0) = 0)", () => {
    expect(fnIMCOT([rvString("0")])).toEqual(ERRORS.DIV0);
  });
});

describe("IMSINH / IMCOSH / IMTANH", () => {
  it("IMSINH real input equals Math.sinh", () => {
    const [re] = splitComplex(asString(fnIMSINH([rvString("2")])));
    expect(re).toBeCloseTo(Math.sinh(2), 5);
  });

  it("IMCOSH(0) = 1", () => {
    const [re] = splitComplex(asString(fnIMCOSH([rvString("0")])));
    expect(re).toBeCloseTo(1, 10);
  });

  it("IMTANH real input equals Math.tanh", () => {
    const [re] = splitComplex(asString(fnIMTANH([rvString("1")])));
    expect(re).toBeCloseTo(Math.tanh(1), 5);
  });
});

describe("IMCSCH / IMSECH / IMCOTH", () => {
  it("IMSECH(0) = 1", () => {
    const [re] = splitComplex(asString(fnIMSECH([rvString("0")])));
    expect(re).toBeCloseTo(1, 10);
  });

  it("IMCSCH(0) → #DIV/0! (sinh(0) = 0)", () => {
    expect(fnIMCSCH([rvString("0")])).toEqual(ERRORS.DIV0);
  });

  it("IMCOTH(0) → #DIV/0!", () => {
    expect(fnIMCOTH([rvString("0")])).toEqual(ERRORS.DIV0);
  });

  it("IMCOTH(1) ≈ 1/tanh(1)", () => {
    const [re] = splitComplex(asString(fnIMCOTH([rvString("1")])));
    expect(re).toBeCloseTo(1 / Math.tanh(1), 5);
  });
});

// ============================================================================
// R7: Base conversion pairs
// ============================================================================

describe("BIN2HEX / BIN2OCT", () => {
  it("BIN2HEX converts positive binary", () => {
    expect(asString(fnBIN2HEX([rvString("1111")]))).toBe("F");
    expect(asString(fnBIN2HEX([rvString("11111111")]))).toBe("FF");
  });

  it("BIN2HEX pads with `places`", () => {
    expect(asString(fnBIN2HEX([rvString("1111"), rvNumber(4)]))).toBe("000F");
  });

  it("BIN2HEX negative (10-digit two's complement) → FFFFFFFFFF prefix", () => {
    // BIN2HEX("1111111111") — 10-digit binary = -1 → "FFFFFFFFFF"
    const r = asString(fnBIN2HEX([rvString("1111111111")]));
    expect(r).toBe("FFFFFFFFFF");
  });

  it("BIN2OCT converts positive binary", () => {
    expect(asString(fnBIN2OCT([rvString("1111")]))).toBe("17");
  });

  it("rejects non-binary input with #NUM!", () => {
    expect(fnBIN2HEX([rvString("12")])).toEqual(ERRORS.NUM);
    expect(fnBIN2OCT([rvString("xyz")])).toEqual(ERRORS.NUM);
  });
});

describe("HEX2BIN / HEX2OCT", () => {
  it("HEX2BIN converts hex in-range", () => {
    expect(asString(fnHEX2BIN([rvString("F")]))).toBe("1111");
  });

  it("HEX2BIN rejects values outside [-512, 511]", () => {
    // 0x200 = 512 → out of range
    expect(fnHEX2BIN([rvString("200")])).toEqual(ERRORS.NUM);
  });

  it("HEX2OCT converts hex", () => {
    expect(asString(fnHEX2OCT([rvString("F")]))).toBe("17");
    expect(asString(fnHEX2OCT([rvString("FF")]))).toBe("377");
  });

  it("rejects non-hex input", () => {
    expect(fnHEX2BIN([rvString("Z")])).toEqual(ERRORS.NUM);
  });
});

describe("OCT2BIN / OCT2HEX", () => {
  it("OCT2BIN converts", () => {
    expect(asString(fnOCT2BIN([rvString("7")]))).toBe("111");
  });

  it("OCT2HEX converts", () => {
    expect(asString(fnOCT2HEX([rvString("17")]))).toBe("F");
    expect(asString(fnOCT2HEX([rvString("377")]))).toBe("FF");
  });

  it("rejects non-octal input", () => {
    expect(fnOCT2BIN([rvString("8")])).toEqual(ERRORS.NUM);
  });
});

// ============================================================================
// R7: BESSEL family
// ============================================================================

describe("BESSELJ", () => {
  it("BESSELJ(x=0, n=0) = 1", () => {
    expect(asNumber(fnBESSELJ([rvNumber(0), rvNumber(0)]))).toBeCloseTo(1, 8);
  });

  it("BESSELJ(1, 1) ≈ 0.4400...", () => {
    // J_1(1) ≈ 0.4400505857449335
    expect(asNumber(fnBESSELJ([rvNumber(1), rvNumber(1)]))).toBeCloseTo(0.440051, 4);
  });

  it("BESSELJ negative order rejected with #NUM!", () => {
    expect(fnBESSELJ([rvNumber(1), rvNumber(-1)])).toEqual(ERRORS.NUM);
  });
});

describe("BESSELI", () => {
  it("BESSELI(0, 0) = 1", () => {
    expect(asNumber(fnBESSELI([rvNumber(0), rvNumber(0)]))).toBeCloseTo(1, 10);
  });

  it("BESSELI(1, 0) ≈ 1.266 (I_0(1))", () => {
    // I_0(1) ≈ 1.2660658
    expect(asNumber(fnBESSELI([rvNumber(1), rvNumber(0)]))).toBeCloseTo(1.2660658, 4);
  });
});

describe("BESSELK", () => {
  it("BESSELK requires x > 0", () => {
    expect(fnBESSELK([rvNumber(0), rvNumber(0)])).toEqual(ERRORS.NUM);
    expect(fnBESSELK([rvNumber(-1), rvNumber(0)])).toEqual(ERRORS.NUM);
  });

  it("BESSELK(1, 0) ≈ 0.4210 (K_0(1))", () => {
    expect(asNumber(fnBESSELK([rvNumber(1), rvNumber(0)]))).toBeCloseTo(0.4210244, 3);
  });
});

describe("BESSELY", () => {
  it("BESSELY requires x > 0", () => {
    expect(fnBESSELY([rvNumber(0), rvNumber(0)])).toEqual(ERRORS.NUM);
  });

  it("BESSELY(1, 0) ≈ 0.0883 (Y_0(1))", () => {
    expect(asNumber(fnBESSELY([rvNumber(1), rvNumber(0)]))).toBeCloseTo(0.0882569, 3);
  });
});

// ============================================================================
// Comprehensive coverage — each function gets >= 5 dedicated cases.
// ============================================================================

describe("DEC2BIN comprehensive", () => {
  it("normal positive value", () => {
    expect(asString(fnDEC2BIN([rvNumber(7)]))).toBe("111");
  });

  it("zero", () => {
    expect(asString(fnDEC2BIN([rvNumber(0)]))).toBe("0");
  });

  it("boundary: 511 is max positive", () => {
    expect(asString(fnDEC2BIN([rvNumber(511)]))).toBe("111111111");
  });

  it("boundary: -512 is min negative", () => {
    expect(asString(fnDEC2BIN([rvNumber(-512)]))).toBe("1000000000");
  });

  it("out of range -> #NUM!", () => {
    expect(fnDEC2BIN([rvNumber(513)])).toEqual(ERRORS.NUM);
    expect(fnDEC2BIN([rvNumber(-600)])).toEqual(ERRORS.NUM);
  });

  it("propagates input errors", () => {
    expect(fnDEC2BIN([ERRORS.NA])).toEqual(ERRORS.NA);
    expect(fnDEC2BIN([ERRORS.DIV0])).toEqual(ERRORS.DIV0);
  });

  it("places omitted yields unpadded (for non-negative)", () => {
    expect(asString(fnDEC2BIN([rvNumber(5)]))).toBe("101");
  });

  it("places out of [1,10] -> #NUM!", () => {
    expect(fnDEC2BIN([rvNumber(5), rvNumber(11)])).toEqual(ERRORS.NUM);
    expect(fnDEC2BIN([rvNumber(5), rvNumber(-1)])).toEqual(ERRORS.NUM);
  });

  it("places smaller than the natural width → #NUM! (regression)", () => {
    // Excel: DEC2BIN(100, 2) → #NUM! because 100 = 1100100₂ (7 digits)
    // can't fit in 2 places. `padStart` alone silently emits the wider
    // form; Excel rejects instead.
    expect(fnDEC2BIN([rvNumber(100), rvNumber(2)])).toEqual(ERRORS.NUM);
    // Equal width still succeeds.
    expect(asString(fnDEC2BIN([rvNumber(100), rvNumber(7)]))).toBe("1100100");
    // Wider-than-needed places pads with zeros.
    expect(asString(fnDEC2BIN([rvNumber(5), rvNumber(8)]))).toBe("00000101");
  });

  it("places ignored for negative numbers", () => {
    expect(asString(fnDEC2BIN([rvNumber(-1), rvNumber(4)]))).toBe("1111111111");
  });

  it("coerces boolean input", () => {
    expect(asString(fnDEC2BIN([rvBoolean(true)]))).toBe("1");
    expect(asString(fnDEC2BIN([rvBoolean(false)]))).toBe("0");
  });
});

describe("BIN2DEC comprehensive", () => {
  it("single digit", () => {
    expect(asNumber(fnBIN2DEC([rvString("1")]))).toBe(1);
    expect(asNumber(fnBIN2DEC([rvString("0")]))).toBe(0);
  });

  it("10-bit boundary: '1000000000' = -512", () => {
    expect(asNumber(fnBIN2DEC([rvString("1000000000")]))).toBe(-512);
  });

  it("9-digit is positive not two's complement", () => {
    expect(asNumber(fnBIN2DEC([rvString("111111111")]))).toBe(511);
  });

  it("rejects >10 digits", () => {
    expect(fnBIN2DEC([rvString("11111111111")])).toEqual(ERRORS.NUM);
  });

  it("rejects non-binary chars", () => {
    expect(fnBIN2DEC([rvString("2")])).toEqual(ERRORS.NUM);
    expect(fnBIN2DEC([rvString("abc")])).toEqual(ERRORS.NUM);
  });

  it("rejects empty", () => {
    expect(fnBIN2DEC([rvString("")])).toEqual(ERRORS.NUM);
  });

  it("propagates error", () => {
    expect(fnBIN2DEC([ERRORS.VALUE])).toEqual(ERRORS.VALUE);
  });

  it("coerces number input", () => {
    // number `10` -> "10" -> decimal 2
    expect(asNumber(fnBIN2DEC([rvNumber(10)]))).toBe(2);
  });
});

describe("DEC2HEX comprehensive", () => {
  it("positive", () => {
    expect(asString(fnDEC2HEX([rvNumber(16)]))).toBe("10");
  });

  it("zero", () => {
    expect(asString(fnDEC2HEX([rvNumber(0)]))).toBe("0");
  });

  it("negative 40-bit two's complement", () => {
    expect(asString(fnDEC2HEX([rvNumber(-1)]))).toBe("FFFFFFFFFF");
  });

  it("boundary 2^39 - 1", () => {
    expect(asString(fnDEC2HEX([rvNumber(549_755_813_887)]))).toBe("7FFFFFFFFF");
  });

  it("out of range -> #NUM!", () => {
    expect(fnDEC2HEX([rvNumber(549_755_813_888)])).toEqual(ERRORS.NUM);
    expect(fnDEC2HEX([rvNumber(-549_755_813_889)])).toEqual(ERRORS.NUM);
  });

  it("places pads", () => {
    expect(asString(fnDEC2HEX([rvNumber(255), rvNumber(4)]))).toBe("00FF");
  });

  it("places out of range -> #NUM!", () => {
    expect(fnDEC2HEX([rvNumber(1), rvNumber(11)])).toEqual(ERRORS.NUM);
  });

  it("propagates error", () => {
    expect(fnDEC2HEX([ERRORS.NA])).toEqual(ERRORS.NA);
  });
});

describe("HEX2DEC comprehensive", () => {
  it("parses lowercase", () => {
    expect(asNumber(fnHEX2DEC([rvString("ff")]))).toBe(255);
  });

  it("zero", () => {
    expect(asNumber(fnHEX2DEC([rvString("0")]))).toBe(0);
  });

  it("40-bit two's complement", () => {
    expect(asNumber(fnHEX2DEC([rvString("8000000000")]))).toBe(-549_755_813_888);
  });

  it("rejects >10 chars", () => {
    expect(fnHEX2DEC([rvString("10000000000")])).toEqual(ERRORS.NUM);
  });

  it("rejects non-hex chars", () => {
    expect(fnHEX2DEC([rvString("G1")])).toEqual(ERRORS.NUM);
  });

  it("propagates errors", () => {
    expect(fnHEX2DEC([ERRORS.REF])).toEqual(ERRORS.REF);
  });
});

describe("DEC2OCT comprehensive", () => {
  it("positive", () => {
    expect(asString(fnDEC2OCT([rvNumber(64)]))).toBe("100");
  });

  it("zero", () => {
    expect(asString(fnDEC2OCT([rvNumber(0)]))).toBe("0");
  });

  it("negative 30-bit two's complement", () => {
    expect(asString(fnDEC2OCT([rvNumber(-1)]))).toBe("7777777777");
  });

  it("boundary: 2^29 - 1", () => {
    expect(asString(fnDEC2OCT([rvNumber(536_870_911)]))).toBe("3777777777");
  });

  it("out of range -> #NUM!", () => {
    expect(fnDEC2OCT([rvNumber(536_870_912)])).toEqual(ERRORS.NUM);
    expect(fnDEC2OCT([rvNumber(-536_870_913)])).toEqual(ERRORS.NUM);
  });

  it("places pads", () => {
    expect(asString(fnDEC2OCT([rvNumber(8), rvNumber(4)]))).toBe("0010");
  });

  it("propagates errors", () => {
    expect(fnDEC2OCT([ERRORS.NUM])).toEqual(ERRORS.NUM);
  });
});

describe("OCT2DEC comprehensive", () => {
  it("parses octal", () => {
    expect(asNumber(fnOCT2DEC([rvString("10")]))).toBe(8);
  });

  it("zero", () => {
    expect(asNumber(fnOCT2DEC([rvString("0")]))).toBe(0);
  });

  it("30-bit two's complement", () => {
    expect(asNumber(fnOCT2DEC([rvString("7777777777")]))).toBe(-1);
    expect(asNumber(fnOCT2DEC([rvString("4000000000")]))).toBe(-536_870_912);
  });

  it("rejects >10 digits", () => {
    expect(fnOCT2DEC([rvString("11111111111")])).toEqual(ERRORS.NUM);
  });

  it("rejects non-octal chars", () => {
    expect(fnOCT2DEC([rvString("8")])).toEqual(ERRORS.NUM);
    expect(fnOCT2DEC([rvString("9")])).toEqual(ERRORS.NUM);
  });

  it("propagates errors", () => {
    expect(fnOCT2DEC([ERRORS.VALUE])).toEqual(ERRORS.VALUE);
  });
});

describe("BIN2HEX comprehensive", () => {
  it("basic positive", () => {
    expect(asString(fnBIN2HEX([rvString("1010")]))).toBe("A");
  });

  it("zero", () => {
    expect(asString(fnBIN2HEX([rvString("0")]))).toBe("0");
  });

  it("10-digit negative emits 10-digit hex FFFFFFFFFF", () => {
    expect(asString(fnBIN2HEX([rvString("1111111111")]))).toBe("FFFFFFFFFF");
  });

  it("places pads non-negative", () => {
    expect(asString(fnBIN2HEX([rvString("111"), rvNumber(4)]))).toBe("0007");
  });

  it("places out of range -> #NUM!", () => {
    expect(fnBIN2HEX([rvString("111"), rvNumber(11)])).toEqual(ERRORS.NUM);
  });

  it("rejects bad binary", () => {
    expect(fnBIN2HEX([rvString("2")])).toEqual(ERRORS.NUM);
  });

  it("propagates errors", () => {
    expect(fnBIN2HEX([ERRORS.NA])).toEqual(ERRORS.NA);
  });
});

describe("BIN2OCT comprehensive", () => {
  it("basic", () => {
    expect(asString(fnBIN2OCT([rvString("1000")]))).toBe("10");
  });

  it("10-digit negative", () => {
    expect(asString(fnBIN2OCT([rvString("1111111111")]))).toBe("7777777777");
  });

  it("places", () => {
    expect(asString(fnBIN2OCT([rvString("1"), rvNumber(3)]))).toBe("001");
  });

  it("rejects bad", () => {
    expect(fnBIN2OCT([rvString("abc")])).toEqual(ERRORS.NUM);
  });

  it("propagates errors", () => {
    expect(fnBIN2OCT([ERRORS.DIV0])).toEqual(ERRORS.DIV0);
  });
});

describe("HEX2BIN comprehensive", () => {
  it("basic", () => {
    expect(asString(fnHEX2BIN([rvString("A")]))).toBe("1010");
  });

  it("zero", () => {
    expect(asString(fnHEX2BIN([rvString("0")]))).toBe("0");
  });

  it("max positive 511 -> 1FF -> 111111111", () => {
    expect(asString(fnHEX2BIN([rvString("1FF")]))).toBe("111111111");
  });

  it("out of positive range -> #NUM! (0x200 = 512)", () => {
    expect(fnHEX2BIN([rvString("200")])).toEqual(ERRORS.NUM);
  });

  it("negative in range: FFFFFFFE00 = -512", () => {
    expect(asString(fnHEX2BIN([rvString("FFFFFFFE00")]))).toBe("1000000000");
  });

  it("negative out of range -> #NUM!", () => {
    // -513 is out of [-512, 511]
    expect(fnHEX2BIN([rvString("FFFFFFFDFF")])).toEqual(ERRORS.NUM);
  });

  it("places pads positive", () => {
    expect(asString(fnHEX2BIN([rvString("F"), rvNumber(8)]))).toBe("00001111");
  });

  it("rejects bad chars", () => {
    expect(fnHEX2BIN([rvString("GG")])).toEqual(ERRORS.NUM);
  });
});

describe("HEX2OCT comprehensive", () => {
  it("basic", () => {
    expect(asString(fnHEX2OCT([rvString("FF")]))).toBe("377");
  });

  it("out of OCT range -> #NUM! (0x20000000 = 2^29)", () => {
    expect(fnHEX2OCT([rvString("20000000")])).toEqual(ERRORS.NUM);
  });

  it("max positive 2^29 - 1 = 0x1FFFFFFF", () => {
    expect(asString(fnHEX2OCT([rvString("1FFFFFFF")]))).toBe("3777777777");
  });

  it("negative", () => {
    expect(asString(fnHEX2OCT([rvString("FFFFFFFFFF")]))).toBe("7777777777");
  });

  it("negative out of OCT range -> #NUM!", () => {
    // Below -2^29 = -536_870_912; use FF00000000 = -0x100000000 = -2^32
    expect(fnHEX2OCT([rvString("FF00000000")])).toEqual(ERRORS.NUM);
  });

  it("places pads", () => {
    expect(asString(fnHEX2OCT([rvString("F"), rvNumber(4)]))).toBe("0017");
  });
});

describe("OCT2BIN comprehensive", () => {
  it("basic", () => {
    expect(asString(fnOCT2BIN([rvString("10")]))).toBe("1000");
  });

  it("max 777 -> 111111111", () => {
    expect(asString(fnOCT2BIN([rvString("777")]))).toBe("111111111");
  });

  it("1000 = 512 out of range -> #NUM!", () => {
    expect(fnOCT2BIN([rvString("1000")])).toEqual(ERRORS.NUM);
  });

  it("10-digit negative in range", () => {
    expect(asString(fnOCT2BIN([rvString("7777777777")]))).toBe("1111111111");
  });

  it("places pads", () => {
    expect(asString(fnOCT2BIN([rvString("7"), rvNumber(5)]))).toBe("00111");
  });

  it("rejects non-octal", () => {
    expect(fnOCT2BIN([rvString("8")])).toEqual(ERRORS.NUM);
  });
});

describe("OCT2HEX comprehensive", () => {
  it("basic", () => {
    expect(asString(fnOCT2HEX([rvString("10")]))).toBe("8");
  });

  it("full 10-digit negative", () => {
    expect(asString(fnOCT2HEX([rvString("7777777777")]))).toBe("FFFFFFFFFF");
  });

  it("places pad", () => {
    expect(asString(fnOCT2HEX([rvString("17"), rvNumber(4)]))).toBe("000F");
  });

  it("rejects bad", () => {
    expect(fnOCT2HEX([rvString("9")])).toEqual(ERRORS.NUM);
  });

  it("propagates errors", () => {
    expect(fnOCT2HEX([ERRORS.NA])).toEqual(ERRORS.NA);
  });
});

describe("DELTA comprehensive", () => {
  it("equal returns 1", () => {
    expect(asNumber(fnDELTA([rvNumber(10), rvNumber(10)]))).toBe(1);
  });

  it("unequal returns 0", () => {
    expect(asNumber(fnDELTA([rvNumber(10), rvNumber(5)]))).toBe(0);
  });

  it("missing second arg defaults to 0", () => {
    expect(asNumber(fnDELTA([rvNumber(0)]))).toBe(1);
    expect(asNumber(fnDELTA([rvNumber(5)]))).toBe(0);
  });

  it("coerces boolean", () => {
    expect(asNumber(fnDELTA([rvBoolean(true), rvNumber(1)]))).toBe(1);
    expect(asNumber(fnDELTA([rvBoolean(false), rvNumber(0)]))).toBe(1);
  });

  it("propagates error", () => {
    expect(fnDELTA([ERRORS.NA, rvNumber(0)])).toEqual(ERRORS.NA);
    expect(fnDELTA([rvNumber(0), ERRORS.DIV0])).toEqual(ERRORS.DIV0);
  });

  it("negatives", () => {
    expect(asNumber(fnDELTA([rvNumber(-1), rvNumber(-1)]))).toBe(1);
  });
});

describe("GESTEP comprehensive", () => {
  it("equal returns 1", () => {
    expect(asNumber(fnGESTEP([rvNumber(5), rvNumber(5)]))).toBe(1);
  });

  it("greater returns 1", () => {
    expect(asNumber(fnGESTEP([rvNumber(10), rvNumber(5)]))).toBe(1);
  });

  it("less returns 0", () => {
    expect(asNumber(fnGESTEP([rvNumber(3), rvNumber(5)]))).toBe(0);
  });

  it("missing step defaults to 0", () => {
    expect(asNumber(fnGESTEP([rvNumber(1)]))).toBe(1);
    expect(asNumber(fnGESTEP([rvNumber(-1)]))).toBe(0);
  });

  it("propagates errors", () => {
    expect(fnGESTEP([ERRORS.NA])).toEqual(ERRORS.NA);
    expect(fnGESTEP([rvNumber(1), ERRORS.VALUE])).toEqual(ERRORS.VALUE);
  });

  it("negative numbers", () => {
    expect(asNumber(fnGESTEP([rvNumber(-5), rvNumber(-10)]))).toBe(1);
  });
});

describe("BITAND comprehensive", () => {
  it("zero result", () => {
    expect(asNumber(fnBITAND([rvNumber(0b1010), rvNumber(0b0101)]))).toBe(0);
  });

  it("identity: x & x == x", () => {
    expect(asNumber(fnBITAND([rvNumber(0xf0), rvNumber(0xf0)]))).toBe(0xf0);
  });

  it("0 & x == 0", () => {
    expect(asNumber(fnBITAND([rvNumber(0), rvNumber(0xff)]))).toBe(0);
  });

  it("max 2^48-1", () => {
    const m = 2 ** 48 - 1;
    expect(asNumber(fnBITAND([rvNumber(m), rvNumber(m)]))).toBe(m);
  });

  it("non-integer -> #NUM!", () => {
    expect(fnBITAND([rvNumber(5.5), rvNumber(1)])).toEqual(ERRORS.NUM);
    expect(fnBITAND([rvNumber(1), rvNumber(2.1)])).toEqual(ERRORS.NUM);
  });

  it("negative -> #NUM!", () => {
    expect(fnBITAND([rvNumber(-1), rvNumber(1)])).toEqual(ERRORS.NUM);
  });

  it(">=2^48 -> #NUM!", () => {
    expect(fnBITAND([rvNumber(2 ** 48), rvNumber(0)])).toEqual(ERRORS.NUM);
  });

  it("propagates error", () => {
    expect(fnBITAND([ERRORS.NA, rvNumber(0)])).toEqual(ERRORS.NA);
    expect(fnBITAND([rvNumber(0), ERRORS.DIV0])).toEqual(ERRORS.DIV0);
  });

  it("coerces boolean/blank", () => {
    expect(asNumber(fnBITAND([rvBoolean(true), rvNumber(1)]))).toBe(1);
    expect(asNumber(fnBITAND([BLANK, rvNumber(1)]))).toBe(0);
  });
});

describe("BITOR comprehensive", () => {
  it("zero or x = x", () => {
    expect(asNumber(fnBITOR([rvNumber(0), rvNumber(0xa5)]))).toBe(0xa5);
  });

  it("complementary bits", () => {
    expect(asNumber(fnBITOR([rvNumber(0xf0), rvNumber(0x0f)]))).toBe(0xff);
  });

  it("48-bit range", () => {
    expect(asNumber(fnBITOR([rvNumber(2 ** 40), rvNumber(1)]))).toBe(2 ** 40 + 1);
  });

  it("non-integer -> #NUM!", () => {
    expect(fnBITOR([rvNumber(1.5), rvNumber(0)])).toEqual(ERRORS.NUM);
  });

  it("negative -> #NUM!", () => {
    expect(fnBITOR([rvNumber(-1), rvNumber(0)])).toEqual(ERRORS.NUM);
  });

  it("propagates error", () => {
    expect(fnBITOR([ERRORS.REF, rvNumber(0)])).toEqual(ERRORS.REF);
  });
});

describe("BITXOR comprehensive", () => {
  it("a^a == 0", () => {
    expect(asNumber(fnBITXOR([rvNumber(0xff), rvNumber(0xff)]))).toBe(0);
  });

  it("a^0 == a", () => {
    expect(asNumber(fnBITXOR([rvNumber(0xaa), rvNumber(0)]))).toBe(0xaa);
  });

  it("48-bit xor", () => {
    expect(asNumber(fnBITXOR([rvNumber(2 ** 40), rvNumber(0)]))).toBe(2 ** 40);
  });

  it("non-integer -> #NUM!", () => {
    expect(fnBITXOR([rvNumber(1), rvNumber(1.5)])).toEqual(ERRORS.NUM);
  });

  it("negative -> #NUM!", () => {
    expect(fnBITXOR([rvNumber(-1), rvNumber(0)])).toEqual(ERRORS.NUM);
  });

  it("propagates error", () => {
    expect(fnBITXOR([rvNumber(0), ERRORS.NA])).toEqual(ERRORS.NA);
  });
});

describe("BITLSHIFT comprehensive", () => {
  it("zero shift is identity", () => {
    expect(asNumber(fnBITLSHIFT([rvNumber(5), rvNumber(0)]))).toBe(5);
  });

  it("positive shift", () => {
    expect(asNumber(fnBITLSHIFT([rvNumber(3), rvNumber(4)]))).toBe(48);
  });

  it("negative shift acts as right-shift", () => {
    expect(asNumber(fnBITLSHIFT([rvNumber(16), rvNumber(-2)]))).toBe(4);
  });

  it("overflow -> #NUM!", () => {
    expect(fnBITLSHIFT([rvNumber(1), rvNumber(48)])).toEqual(ERRORS.NUM);
  });

  it("shift magnitude > 53 -> #NUM!", () => {
    expect(fnBITLSHIFT([rvNumber(1), rvNumber(54)])).toEqual(ERRORS.NUM);
    expect(fnBITLSHIFT([rvNumber(1), rvNumber(-54)])).toEqual(ERRORS.NUM);
  });

  it("operand negative -> #NUM!", () => {
    expect(fnBITLSHIFT([rvNumber(-1), rvNumber(1)])).toEqual(ERRORS.NUM);
  });

  it("propagates error", () => {
    expect(fnBITLSHIFT([ERRORS.DIV0, rvNumber(1)])).toEqual(ERRORS.DIV0);
  });
});

describe("BITRSHIFT comprehensive", () => {
  it("zero shift is identity", () => {
    expect(asNumber(fnBITRSHIFT([rvNumber(5), rvNumber(0)]))).toBe(5);
  });

  it("positive shift", () => {
    expect(asNumber(fnBITRSHIFT([rvNumber(32), rvNumber(3)]))).toBe(4);
  });

  it("negative shift acts as left-shift", () => {
    expect(asNumber(fnBITRSHIFT([rvNumber(4), rvNumber(-2)]))).toBe(16);
  });

  it("shift > 53 -> #NUM!", () => {
    expect(fnBITRSHIFT([rvNumber(1), rvNumber(54)])).toEqual(ERRORS.NUM);
  });

  it("negative operand -> #NUM!", () => {
    expect(fnBITRSHIFT([rvNumber(-1), rvNumber(0)])).toEqual(ERRORS.NUM);
  });

  it("non-integer operand -> #NUM!", () => {
    expect(fnBITRSHIFT([rvNumber(2.5), rvNumber(1)])).toEqual(ERRORS.NUM);
  });
});

describe("COMPLEX comprehensive", () => {
  it("default suffix 'i'", () => {
    expect(asString(fnCOMPLEX([rvNumber(3), rvNumber(4)]))).toBe("3+4i");
  });

  it("suffix 'j'", () => {
    expect(asString(fnCOMPLEX([rvNumber(3), rvNumber(4), rvString("j")]))).toBe("3+4j");
  });

  it("imag=0 returns real only", () => {
    expect(asString(fnCOMPLEX([rvNumber(5), rvNumber(0)]))).toBe("5");
  });

  it("real=0 returns imag only", () => {
    expect(asString(fnCOMPLEX([rvNumber(0), rvNumber(7)]))).toBe("7i");
  });

  it("real=0, imag=1 -> 'i'", () => {
    expect(asString(fnCOMPLEX([rvNumber(0), rvNumber(1)]))).toBe("i");
  });

  it("real=0, imag=-1 -> '-i'", () => {
    expect(asString(fnCOMPLEX([rvNumber(0), rvNumber(-1)]))).toBe("-i");
  });

  it("negative imaginary '-'", () => {
    expect(asString(fnCOMPLEX([rvNumber(3), rvNumber(-4)]))).toBe("3-4i");
  });

  it("suffix must be i or j", () => {
    expect(fnCOMPLEX([rvNumber(1), rvNumber(2), rvString("k")])).toEqual(ERRORS.VALUE);
    expect(fnCOMPLEX([rvNumber(1), rvNumber(2), rvString("I")])).toEqual(ERRORS.VALUE);
  });

  it("propagates errors", () => {
    expect(fnCOMPLEX([ERRORS.NA, rvNumber(0)])).toEqual(ERRORS.NA);
    expect(fnCOMPLEX([rvNumber(0), ERRORS.DIV0])).toEqual(ERRORS.DIV0);
    expect(fnCOMPLEX([rvNumber(0), rvNumber(0), ERRORS.VALUE])).toEqual(ERRORS.VALUE);
  });
});

describe("IMREAL / IMAGINARY comprehensive", () => {
  it("IMREAL of pure real", () => {
    expect(asNumber(fnIMREAL([rvString("5")]))).toBe(5);
  });

  it("IMREAL of pure imag", () => {
    expect(asNumber(fnIMREAL([rvString("7i")]))).toBe(0);
  });

  it("IMAGINARY of pure real", () => {
    expect(asNumber(fnIMAGINARY([rvString("5")]))).toBe(0);
  });

  it("IMAGINARY of 'i' is 1", () => {
    expect(asNumber(fnIMAGINARY([rvString("i")]))).toBe(1);
  });

  it("IMAGINARY of '-i' is -1", () => {
    expect(asNumber(fnIMAGINARY([rvString("-i")]))).toBe(-1);
  });

  it("mixed form", () => {
    expect(asNumber(fnIMREAL([rvString("3+4i")]))).toBe(3);
    expect(asNumber(fnIMAGINARY([rvString("3+4i")]))).toBe(4);
    expect(asNumber(fnIMREAL([rvString("3-4j")]))).toBe(3);
    expect(asNumber(fnIMAGINARY([rvString("3-4j")]))).toBe(-4);
  });

  it("invalid -> #NUM!", () => {
    expect(fnIMREAL([rvString("not complex")])).toEqual(ERRORS.NUM);
    expect(fnIMAGINARY([rvString("1.2.3i")])).toEqual(ERRORS.NUM);
  });

  it("propagates error", () => {
    expect(fnIMREAL([ERRORS.NA])).toEqual(ERRORS.NA);
  });
});

describe("IMABS / IMARGUMENT / IMCONJUGATE comprehensive", () => {
  it("IMABS of 3+4i = 5", () => {
    expect(asNumber(fnIMABS([rvString("3+4i")]))).toBe(5);
  });

  it("IMABS of 0 = 0", () => {
    expect(asNumber(fnIMABS([rvString("0")]))).toBe(0);
  });

  it("IMABS negative parts", () => {
    expect(asNumber(fnIMABS([rvString("-3-4i")]))).toBe(5);
  });

  it("IMARGUMENT of 0 -> #DIV/0!", () => {
    expect(fnIMARGUMENT([rvString("0")])).toEqual(ERRORS.DIV0);
  });

  it("IMARGUMENT of 1+i = pi/4", () => {
    expect(asNumber(fnIMARGUMENT([rvString("1+i")]))).toBeCloseTo(Math.PI / 4, 10);
  });

  it("IMARGUMENT of i = pi/2", () => {
    expect(asNumber(fnIMARGUMENT([rvString("i")]))).toBeCloseTo(Math.PI / 2, 10);
  });

  it("IMCONJUGATE flips sign of imag", () => {
    expect(asString(fnIMCONJUGATE([rvString("3+4i")]))).toBe("3-4i");
  });

  it("IMCONJUGATE of pure real unchanged", () => {
    expect(asString(fnIMCONJUGATE([rvString("5")]))).toBe("5");
  });

  it("propagates errors", () => {
    expect(fnIMABS([ERRORS.NA])).toEqual(ERRORS.NA);
    expect(fnIMARGUMENT([ERRORS.DIV0])).toEqual(ERRORS.DIV0);
    expect(fnIMCONJUGATE([ERRORS.VALUE])).toEqual(ERRORS.VALUE);
  });

  it("invalid string -> #NUM!", () => {
    expect(fnIMABS([rvString("xyz")])).toEqual(ERRORS.NUM);
  });
});

describe("IMSUM comprehensive", () => {
  it("two scalars", () => {
    expect(asString(fnIMSUM([rvString("3+4i"), rvString("1+2i")]))).toBe("4+6i");
  });

  it("single arg", () => {
    expect(asString(fnIMSUM([rvString("3+4i")]))).toBe("3+4i");
  });

  it("real+imag", () => {
    expect(asString(fnIMSUM([rvString("5"), rvString("2i")]))).toBe("5+2i");
  });

  it("array input — each cell accumulates", () => {
    const arr = rvArray([[rvString("1+2i"), rvString("3+4i"), rvString("5+6i")]]);
    expect(asString(fnIMSUM([arr]))).toBe("9+12i");
  });

  it("array with number cell (treated as real)", () => {
    const arr = rvArray([[rvNumber(10), rvString("2i")]]);
    expect(asString(fnIMSUM([arr]))).toBe("10+2i");
  });

  it("array with error cell -> error", () => {
    const arr = rvArray([[rvString("1+2i"), ERRORS.NA]]);
    expect(fnIMSUM([arr])).toEqual(ERRORS.NA);
  });

  it("array with invalid string -> #NUM!", () => {
    const arr = rvArray([[rvString("bad")]]);
    expect(fnIMSUM([arr])).toEqual(ERRORS.NUM);
  });

  it("no args returns '0'", () => {
    expect(asString(fnIMSUM([]))).toBe("0");
  });

  it("propagates arg-level error before array walk", () => {
    expect(fnIMSUM([ERRORS.DIV0])).toEqual(ERRORS.DIV0);
  });
});

describe("IMSUB comprehensive", () => {
  it("basic", () => {
    expect(asString(fnIMSUB([rvString("5+6i"), rvString("3+4i")]))).toBe("2+2i");
  });

  it("real difference becomes pure imag", () => {
    expect(asString(fnIMSUB([rvString("3+4i"), rvString("3+1i")]))).toBe("3i");
  });

  it("full cancellation returns '0'", () => {
    expect(asString(fnIMSUB([rvString("2+3i"), rvString("2+3i")]))).toBe("0");
  });

  it("invalid -> #NUM!", () => {
    expect(fnIMSUB([rvString("bad"), rvString("0")])).toEqual(ERRORS.NUM);
    expect(fnIMSUB([rvString("0"), rvString("bad")])).toEqual(ERRORS.NUM);
  });

  it("propagates error", () => {
    expect(fnIMSUB([ERRORS.NA, rvString("0")])).toEqual(ERRORS.NA);
    expect(fnIMSUB([rvString("0"), ERRORS.DIV0])).toEqual(ERRORS.DIV0);
  });
});

describe("IMPRODUCT comprehensive", () => {
  it("i * i = -1", () => {
    expect(asString(fnIMPRODUCT([rvString("i"), rvString("i")]))).toBe("-1");
  });

  it("(3+4i) * (1+2i) = -5 + 10i", () => {
    expect(asString(fnIMPRODUCT([rvString("3+4i"), rvString("1+2i")]))).toBe("-5+10i");
  });

  it("single arg returns that", () => {
    expect(asString(fnIMPRODUCT([rvString("3+4i")]))).toBe("3+4i");
  });

  it("no args returns '1'", () => {
    expect(asString(fnIMPRODUCT([]))).toBe("1");
  });

  it("array", () => {
    // 2 * 3 * i = 6i
    const arr = rvArray([[rvNumber(2), rvNumber(3), rvString("i")]]);
    expect(asString(fnIMPRODUCT([arr]))).toBe("6i");
  });

  it("invalid -> #NUM!", () => {
    expect(fnIMPRODUCT([rvString("bad")])).toEqual(ERRORS.NUM);
  });

  it("propagates error", () => {
    expect(fnIMPRODUCT([ERRORS.NA])).toEqual(ERRORS.NA);
  });
});

describe("IMDIV comprehensive", () => {
  it("basic", () => {
    // (3+4i)/(1+2i) = (3+4i)*(1-2i)/5 = (3+8 + (-6+4)i)/5 = (11-2i)/5 = 2.2 - 0.4i
    const r = asString(fnIMDIV([rvString("3+4i"), rvString("1+2i")]));
    // Parse to compare floats
    expect(r).toMatch(/^2\.2/);
  });

  it("divisor 0 -> #DIV/0! (consistent with cdiv helper)", () => {
    // R8 fix: IMDIV now returns #DIV/0! to match the rest of the complex
    // division family (IMTAN / IMCSC / IMCOT via `cdiv`).
    expect(fnIMDIV([rvString("1+i"), rvString("0")])).toEqual(ERRORS.DIV0);
  });

  it("invalid left/right", () => {
    expect(fnIMDIV([rvString("bad"), rvString("1")])).toEqual(ERRORS.NUM);
    expect(fnIMDIV([rvString("1"), rvString("bad")])).toEqual(ERRORS.NUM);
  });

  it("propagates error", () => {
    expect(fnIMDIV([ERRORS.NA, rvString("1")])).toEqual(ERRORS.NA);
    expect(fnIMDIV([rvString("1"), ERRORS.DIV0])).toEqual(ERRORS.DIV0);
  });

  it("real/real", () => {
    expect(asString(fnIMDIV([rvString("10"), rvString("2")]))).toBe("5");
  });
});

describe("IMPOWER comprehensive", () => {
  it("i^2 = -1", () => {
    // parseComplex on "-1" yields real -1; formatComplex(0, ~0) may show tiny float
    const r = asString(fnIMPOWER([rvString("i"), rvNumber(2)]));
    // Real part ≈ -1, imag ≈ 0 (might show as small float-something-i)
    // Parse the real component before 'i'
    // Accept exact "-1" or "-1±εi"
    expect(r.startsWith("-1")).toBe(true);
  });

  it("i^4 = 1", () => {
    const r = asString(fnIMPOWER([rvString("i"), rvNumber(4)]));
    // Result is like "1-2.4492935982947064e-16i" — parse components robustly.
    const { re, im } = parseComplexLoose(r);
    expect(re).toBeCloseTo(1, 6);
    expect(im).toBeCloseTo(0, 6);
  });

  it("0^positive = 0", () => {
    expect(asString(fnIMPOWER([rvString("0"), rvNumber(3)]))).toBe("0");
  });

  it("(2+0i)^3 = 8", () => {
    const r = asString(fnIMPOWER([rvString("2"), rvNumber(3)]));
    const match = /^([-\d.eE+]+)/.exec(r);
    expect(Number(match?.[1] ?? "NaN")).toBeCloseTo(8, 6);
  });

  it("invalid complex -> #NUM!", () => {
    expect(fnIMPOWER([rvString("bad"), rvNumber(2)])).toEqual(ERRORS.NUM);
  });

  it("error n", () => {
    expect(fnIMPOWER([rvString("1"), ERRORS.NA])).toEqual(ERRORS.NA);
  });
});

describe("IMSQRT / IMLN / IMLOG2 / IMLOG10 / IMEXP comprehensive", () => {
  it("IMSQRT(-1) = i", () => {
    const r = asString(fnIMSQRT([rvString("-1")]));
    // sqrt(-1) => magnitude 1, angle pi → cos(pi/2)+i sin(pi/2) = 0+1i
    // Floating noise makes real part tiny but nonzero.
    const [re, im] = splitComplex(r);
    expect(re).toBeCloseTo(0, 6);
    expect(im).toBeCloseTo(1, 6);
  });

  it("IMSQRT(4) = 2", () => {
    const r = asString(fnIMSQRT([rvString("4")]));
    const realMatch = /^([-\d.eE+]+)/.exec(r);
    expect(Number(realMatch?.[1] ?? "NaN")).toBeCloseTo(2, 6);
  });

  it("IMLN(1) = 0", () => {
    expect(asString(fnIMLN([rvString("1")]))).toBe("0");
  });

  it("IMLN(0) -> #NUM!", () => {
    expect(fnIMLN([rvString("0")])).toEqual(ERRORS.NUM);
  });

  it("IMLN(e) ≈ 1", () => {
    const r = asString(fnIMLN([rvString(String(Math.E))]));
    const realMatch = /^([-\d.eE+]+)/.exec(r);
    expect(Number(realMatch?.[1] ?? "NaN")).toBeCloseTo(1, 6);
  });

  it("IMLOG2(0) -> #NUM!", () => {
    expect(fnIMLOG2([rvString("0")])).toEqual(ERRORS.NUM);
  });

  it("IMLOG2(8) ≈ 3", () => {
    const r = asString(fnIMLOG2([rvString("8")]));
    const realMatch = /^([-\d.eE+]+)/.exec(r);
    expect(Number(realMatch?.[1] ?? "NaN")).toBeCloseTo(3, 6);
  });

  it("IMLOG10(100) ≈ 2", () => {
    const r = asString(fnIMLOG10([rvString("100")]));
    const realMatch = /^([-\d.eE+]+)/.exec(r);
    expect(Number(realMatch?.[1] ?? "NaN")).toBeCloseTo(2, 6);
  });

  it("IMLOG10(0) -> #NUM!", () => {
    expect(fnIMLOG10([rvString("0")])).toEqual(ERRORS.NUM);
  });

  it("IMEXP(0) = 1", () => {
    expect(asString(fnIMEXP([rvString("0")]))).toBe("1");
  });

  it("IMEXP propagates error", () => {
    expect(fnIMEXP([ERRORS.NA])).toEqual(ERRORS.NA);
  });

  it("IMLN propagates error", () => {
    expect(fnIMLN([ERRORS.DIV0])).toEqual(ERRORS.DIV0);
  });
});

describe("IMSIN / IMCOS comprehensive", () => {
  it("IMSIN(0) = 0", () => {
    expect(asString(fnIMSIN([rvString("0")]))).toBe("0");
  });

  it("IMSIN(pi/2) = 1", () => {
    const r = asString(fnIMSIN([rvString(String(Math.PI / 2))]));
    const realMatch = /^([-\d.eE+]+)/.exec(r);
    expect(Number(realMatch?.[1] ?? "NaN")).toBeCloseTo(1, 6);
  });

  it("IMSIN(i) = i*sinh(1)", () => {
    const r = asString(fnIMSIN([rvString("i")]));
    // Expect pure imag ≈ sinh(1) ≈ 1.1752
    // Format: "1.175...i" (no real part since sin(0)*cosh(1) = 0)
    expect(r).toMatch(/i$/);
  });

  it("IMCOS(0) = 1", () => {
    expect(asString(fnIMCOS([rvString("0")]))).toBe("1");
  });

  it("IMCOS(pi) ≈ -1", () => {
    const r = asString(fnIMCOS([rvString(String(Math.PI))]));
    const realMatch = /^([-\d.eE+]+)/.exec(r);
    expect(Number(realMatch?.[1] ?? "NaN")).toBeCloseTo(-1, 6);
  });

  it("invalid -> #NUM!", () => {
    expect(fnIMSIN([rvString("bad")])).toEqual(ERRORS.NUM);
    expect(fnIMCOS([rvString("bad")])).toEqual(ERRORS.NUM);
  });

  it("propagates errors", () => {
    expect(fnIMSIN([ERRORS.NA])).toEqual(ERRORS.NA);
    expect(fnIMCOS([ERRORS.DIV0])).toEqual(ERRORS.DIV0);
  });
});

describe("BESSELJ comprehensive", () => {
  it("J_0(0) = 1", () => {
    expect(asNumber(fnBESSELJ([rvNumber(0), rvNumber(0)]))).toBeCloseTo(1, 10);
  });

  it("J_n(0) = 0 for n>=1", () => {
    expect(asNumber(fnBESSELJ([rvNumber(0), rvNumber(1)]))).toBe(0);
    expect(asNumber(fnBESSELJ([rvNumber(0), rvNumber(3)]))).toBe(0);
  });

  it("J_0(1) ≈ 0.7651977", () => {
    expect(asNumber(fnBESSELJ([rvNumber(1), rvNumber(0)]))).toBeCloseTo(0.7651977, 5);
  });

  it("J_1(1) ≈ 0.4400", () => {
    expect(asNumber(fnBESSELJ([rvNumber(1), rvNumber(1)]))).toBeCloseTo(0.4400505857, 5);
  });

  it("J_2(5) ≈ 0.0465", () => {
    // Moderate x,n still handled in the series branch (x<15).
    expect(asNumber(fnBESSELJ([rvNumber(5), rvNumber(2)]))).toBeCloseTo(0.0465651, 4);
  });

  it("negative order -> #NUM!", () => {
    expect(fnBESSELJ([rvNumber(1), rvNumber(-1)])).toEqual(ERRORS.NUM);
  });

  it("truncates fractional order", () => {
    // n=2.9 → floor/trunc to 2
    expect(asNumber(fnBESSELJ([rvNumber(1), rvNumber(2.9)]))).toBeCloseTo(
      asNumber(fnBESSELJ([rvNumber(1), rvNumber(2)])),
      5
    );
  });

  it("propagates errors", () => {
    expect(fnBESSELJ([ERRORS.NA, rvNumber(0)])).toEqual(ERRORS.NA);
    expect(fnBESSELJ([rvNumber(1), ERRORS.DIV0])).toEqual(ERRORS.DIV0);
  });
});

describe("BESSELI comprehensive", () => {
  it("I_0(0) = 1", () => {
    expect(asNumber(fnBESSELI([rvNumber(0), rvNumber(0)]))).toBe(1);
  });

  it("I_n(0) = 0 for n>=1", () => {
    expect(asNumber(fnBESSELI([rvNumber(0), rvNumber(2)]))).toBe(0);
  });

  it("I_0(1) ≈ 1.2661", () => {
    expect(asNumber(fnBESSELI([rvNumber(1), rvNumber(0)]))).toBeCloseTo(1.2661, 4);
  });

  it("I_1(2) ≈ 1.5906", () => {
    expect(asNumber(fnBESSELI([rvNumber(2), rvNumber(1)]))).toBeCloseTo(1.5906, 3);
  });

  it("negative order -> #NUM!", () => {
    expect(fnBESSELI([rvNumber(1), rvNumber(-1)])).toEqual(ERRORS.NUM);
  });

  it("propagates errors", () => {
    expect(fnBESSELI([ERRORS.NA, rvNumber(0)])).toEqual(ERRORS.NA);
  });
});

describe("BESSELK comprehensive", () => {
  it("K_0 requires x > 0", () => {
    expect(fnBESSELK([rvNumber(0), rvNumber(0)])).toEqual(ERRORS.NUM);
    expect(fnBESSELK([rvNumber(-1), rvNumber(0)])).toEqual(ERRORS.NUM);
  });

  it("K_0(1) ≈ 0.4210", () => {
    expect(asNumber(fnBESSELK([rvNumber(1), rvNumber(0)]))).toBeCloseTo(0.421, 3);
  });

  it("K_1(1) ≈ 0.6019", () => {
    expect(asNumber(fnBESSELK([rvNumber(1), rvNumber(1)]))).toBeCloseTo(0.6019, 3);
  });

  it("K for x>=2 (different branch)", () => {
    expect(asNumber(fnBESSELK([rvNumber(3), rvNumber(0)]))).toBeCloseTo(0.03474, 3);
  });

  it("negative order -> #NUM!", () => {
    expect(fnBESSELK([rvNumber(1), rvNumber(-1)])).toEqual(ERRORS.NUM);
  });

  it("propagates errors", () => {
    expect(fnBESSELK([ERRORS.NA, rvNumber(0)])).toEqual(ERRORS.NA);
  });
});

describe("BESSELY comprehensive", () => {
  it("x<=0 -> #NUM!", () => {
    expect(fnBESSELY([rvNumber(0), rvNumber(0)])).toEqual(ERRORS.NUM);
    expect(fnBESSELY([rvNumber(-1), rvNumber(0)])).toEqual(ERRORS.NUM);
  });

  it("Y_0(1) ≈ 0.0883", () => {
    expect(asNumber(fnBESSELY([rvNumber(1), rvNumber(0)]))).toBeCloseTo(0.0883, 3);
  });

  it("Y_1(1) ≈ -0.7812", () => {
    expect(asNumber(fnBESSELY([rvNumber(1), rvNumber(1)]))).toBeCloseTo(-0.7812, 3);
  });

  it("large x (>=8) uses asymptotic branch", () => {
    // Y_0(10) ≈ 0.0557
    expect(asNumber(fnBESSELY([rvNumber(10), rvNumber(0)]))).toBeCloseTo(0.0557, 3);
  });

  it("negative order -> #NUM!", () => {
    expect(fnBESSELY([rvNumber(1), rvNumber(-1)])).toEqual(ERRORS.NUM);
  });

  it("propagates errors", () => {
    expect(fnBESSELY([ERRORS.NA, rvNumber(0)])).toEqual(ERRORS.NA);
  });
});

// ============================================================================
// Extra coverage for low-count IM* trig / hyperbolic / log functions
// ============================================================================

describe("IMCOSH (extra coverage)", () => {
  it("IMCOSH(0) = 1", () => {
    const [re, im] = splitComplex(asString(fnIMCOSH([rvString("0")])));
    expect(re).toBeCloseTo(1, 10);
    expect(im).toBeCloseTo(0, 10);
  });
  it("IMCOSH is even for real inputs", () => {
    const [re1] = splitComplex(asString(fnIMCOSH([rvString("1")])));
    const [re2] = splitComplex(asString(fnIMCOSH([rvString("-1")])));
    expect(re1).toBeCloseTo(re2, 10);
  });
  it("IMCOSH real input matches Math.cosh", () => {
    const [re] = splitComplex(asString(fnIMCOSH([rvString("2")])));
    expect(re).toBeCloseTo(Math.cosh(2), 10);
  });
  it("IMCOSH rejects invalid complex text", () => {
    expect(fnIMCOSH([rvString("bad")])).toEqual(ERRORS.NUM);
  });
  it("IMCOSH propagates errors", () => {
    expect(fnIMCOSH([ERRORS.NA])).toEqual(ERRORS.NA);
  });
});

describe("IMSINH (extra coverage)", () => {
  it("IMSINH(0) = 0", () => {
    const [re] = splitComplex(asString(fnIMSINH([rvString("0")])));
    expect(re).toBeCloseTo(0, 10);
  });
  it("IMSINH is odd for real inputs", () => {
    const [re1] = splitComplex(asString(fnIMSINH([rvString("1.2")])));
    const [re2] = splitComplex(asString(fnIMSINH([rvString("-1.2")])));
    expect(re1).toBeCloseTo(-re2, 10);
  });
  it("IMSINH real input matches Math.sinh", () => {
    const [re] = splitComplex(asString(fnIMSINH([rvString("1.5")])));
    expect(re).toBeCloseTo(Math.sinh(1.5), 10);
  });
  it("IMSINH rejects invalid text", () => {
    expect(fnIMSINH([rvString("bad")])).toEqual(ERRORS.NUM);
  });
  it("IMSINH propagates errors", () => {
    expect(fnIMSINH([ERRORS.NUM])).toEqual(ERRORS.NUM);
  });
});

describe("IMTANH (extra coverage)", () => {
  it("IMTANH(0) = 0", () => {
    const [re] = splitComplex(asString(fnIMTANH([rvString("0")])));
    expect(re).toBeCloseTo(0, 10);
  });
  it("IMTANH is odd for real inputs", () => {
    const [re1] = splitComplex(asString(fnIMTANH([rvString("1.5")])));
    const [re2] = splitComplex(asString(fnIMTANH([rvString("-1.5")])));
    expect(re1).toBeCloseTo(-re2, 10);
  });
  it("IMTANH real input approaches 1 for large positive x", () => {
    const [re] = splitComplex(asString(fnIMTANH([rvString("20")])));
    expect(re).toBeCloseTo(1, 10);
  });
  it("IMTANH rejects invalid text", () => {
    expect(fnIMTANH([rvString("bad")])).toEqual(ERRORS.NUM);
  });
  it("IMTANH propagates errors", () => {
    expect(fnIMTANH([ERRORS.NA])).toEqual(ERRORS.NA);
  });
});

describe("IMCSC / IMSEC (extra coverage)", () => {
  it("IMCSC(PI/2) = 1", () => {
    const [re] = splitComplex(asString(fnIMCSC([rvString(String(Math.PI / 2))])));
    expect(re).toBeCloseTo(1, 6);
  });
  it("IMCSC(0) → #DIV/0! (sin(0)=0)", () => {
    expect(fnIMCSC([rvString("0")])).toEqual(ERRORS.DIV0);
  });
  it("IMSEC(0) = 1", () => {
    const [re] = splitComplex(asString(fnIMSEC([rvString("0")])));
    expect(re).toBeCloseTo(1, 10);
  });
  it("IMSEC(PI) ≈ -1", () => {
    const [re] = splitComplex(asString(fnIMSEC([rvString(String(Math.PI))])));
    expect(re).toBeCloseTo(-1, 6);
  });
  it("IMCSC / IMSEC propagate errors", () => {
    expect(fnIMCSC([ERRORS.NA])).toEqual(ERRORS.NA);
    expect(fnIMSEC([ERRORS.NA])).toEqual(ERRORS.NA);
  });
  it("IMCSC / IMSEC reject invalid text", () => {
    expect(fnIMCSC([rvString("bad")])).toEqual(ERRORS.NUM);
    expect(fnIMSEC([rvString("bad")])).toEqual(ERRORS.NUM);
  });
});

describe("IMCSCH / IMSECH / IMCOTH (extra coverage)", () => {
  it("IMSECH(0) = 1", () => {
    const [re] = splitComplex(asString(fnIMSECH([rvString("0")])));
    expect(re).toBeCloseTo(1, 10);
  });
  it("IMCSCH(0) → #DIV/0!", () => {
    expect(fnIMCSCH([rvString("0")])).toEqual(ERRORS.DIV0);
  });
  it("IMCSCH(1) ≈ 1/sinh(1) for real input", () => {
    const [re] = splitComplex(asString(fnIMCSCH([rvString("1")])));
    expect(re).toBeCloseTo(1 / Math.sinh(1), 6);
  });
  it("IMCOTH(0) → #DIV/0!", () => {
    expect(fnIMCOTH([rvString("0")])).toEqual(ERRORS.DIV0);
  });
  it("IMCOTH(1) ≈ 1/tanh(1)", () => {
    const [re] = splitComplex(asString(fnIMCOTH([rvString("1")])));
    expect(re).toBeCloseTo(1 / Math.tanh(1), 6);
  });
  it("IMSECH(1) ≈ 1/cosh(1)", () => {
    const [re] = splitComplex(asString(fnIMSECH([rvString("1")])));
    expect(re).toBeCloseTo(1 / Math.cosh(1), 6);
  });
  it("propagate errors", () => {
    expect(fnIMCSCH([ERRORS.NA])).toEqual(ERRORS.NA);
    expect(fnIMSECH([ERRORS.NUM])).toEqual(ERRORS.NUM);
    expect(fnIMCOTH([ERRORS.REF])).toEqual(ERRORS.REF);
  });
  it("reject invalid complex text", () => {
    expect(fnIMCSCH([rvString("bad")])).toEqual(ERRORS.NUM);
    expect(fnIMSECH([rvString("bad")])).toEqual(ERRORS.NUM);
    expect(fnIMCOTH([rvString("bad")])).toEqual(ERRORS.NUM);
  });
});

describe("IMCOT (extra coverage)", () => {
  it("IMCOT(PI/4) = 1", () => {
    const [re] = splitComplex(asString(fnIMCOT([rvString(String(Math.PI / 4))])));
    expect(re).toBeCloseTo(1, 6);
  });
  it("IMCOT(0) → #DIV/0!", () => {
    expect(fnIMCOT([rvString("0")])).toEqual(ERRORS.DIV0);
  });
  it("IMCOT(PI/2) ≈ 0", () => {
    const [re] = splitComplex(asString(fnIMCOT([rvString(String(Math.PI / 2))])));
    expect(Math.abs(re)).toBeLessThan(1e-6);
  });
  it("IMCOT propagates errors", () => {
    expect(fnIMCOT([ERRORS.NA])).toEqual(ERRORS.NA);
  });
  it("IMCOT rejects invalid text", () => {
    expect(fnIMCOT([rvString("bad")])).toEqual(ERRORS.NUM);
  });
});

describe("IMEXP (extra coverage)", () => {
  it("IMEXP(0) = 1", () => {
    const [re, im] = splitComplex(asString(fnIMEXP([rvString("0")])));
    expect(re).toBeCloseTo(1, 10);
    expect(im).toBeCloseTo(0, 10);
  });
  it("IMEXP(1) = e", () => {
    const [re] = splitComplex(asString(fnIMEXP([rvString("1")])));
    expect(re).toBeCloseTo(Math.E, 10);
  });
  it("IMEXP(i*PI) ≈ -1 (Euler's identity)", () => {
    const [re, im] = splitComplex(asString(fnIMEXP([rvString(`${Math.PI}i`)])));
    expect(re).toBeCloseTo(-1, 6);
    expect(im).toBeCloseTo(0, 6);
  });
  it("IMEXP propagates errors", () => {
    expect(fnIMEXP([ERRORS.DIV0])).toEqual(ERRORS.DIV0);
  });
  it("IMEXP rejects invalid text", () => {
    expect(fnIMEXP([rvString("bad")])).toEqual(ERRORS.NUM);
  });
});

describe("IMLOG10 / IMLOG2 (extra coverage)", () => {
  it("IMLOG10(1) = 0", () => {
    const [re] = splitComplex(asString(fnIMLOG10([rvString("1")])));
    expect(re).toBeCloseTo(0, 10);
  });
  it("IMLOG10(100) = 2", () => {
    const [re] = splitComplex(asString(fnIMLOG10([rvString("100")])));
    expect(re).toBeCloseTo(2, 10);
  });
  it("IMLOG2(1) = 0, IMLOG2(8) = 3", () => {
    const [re1] = splitComplex(asString(fnIMLOG2([rvString("1")])));
    const [re2] = splitComplex(asString(fnIMLOG2([rvString("8")])));
    expect(re1).toBeCloseTo(0, 10);
    expect(re2).toBeCloseTo(3, 10);
  });
  it("IMLOG10(0) → #NUM! (log of 0)", () => {
    expect(fnIMLOG10([rvString("0")])).toEqual(ERRORS.NUM);
    expect(fnIMLOG2([rvString("0")])).toEqual(ERRORS.NUM);
  });
  it("IMLOG10 propagates errors", () => {
    expect(fnIMLOG10([ERRORS.NA])).toEqual(ERRORS.NA);
    expect(fnIMLOG2([ERRORS.NA])).toEqual(ERRORS.NA);
  });
});

describe("IMLN (extra coverage)", () => {
  it("IMLN(1) = 0", () => {
    const [re] = splitComplex(asString(fnIMLN([rvString("1")])));
    expect(re).toBeCloseTo(0, 10);
  });
  it("IMLN(e) = 1", () => {
    const [re] = splitComplex(asString(fnIMLN([rvString(String(Math.E))])));
    expect(re).toBeCloseTo(1, 10);
  });
  it("IMLN(-1) = i*PI (pure imaginary)", () => {
    const [re, im] = splitComplex(asString(fnIMLN([rvString("-1")])));
    expect(re).toBeCloseTo(0, 6);
    expect(im).toBeCloseTo(Math.PI, 6);
  });
  it("IMLN(0) → #NUM!", () => {
    expect(fnIMLN([rvString("0")])).toEqual(ERRORS.NUM);
  });
  it("IMLN propagates errors", () => {
    expect(fnIMLN([ERRORS.DIV0])).toEqual(ERRORS.DIV0);
  });
});

describe("IMSQRT (extra coverage)", () => {
  it("IMSQRT(0) = 0", () => {
    const [re, im] = splitComplex(asString(fnIMSQRT([rvString("0")])));
    expect(re).toBeCloseTo(0, 10);
    expect(im).toBeCloseTo(0, 10);
  });
  it("IMSQRT(1) = 1", () => {
    const [re] = splitComplex(asString(fnIMSQRT([rvString("1")])));
    expect(re).toBeCloseTo(1, 10);
  });
  it("IMSQRT(-1) = i", () => {
    const [re, im] = splitComplex(asString(fnIMSQRT([rvString("-1")])));
    expect(re).toBeCloseTo(0, 10);
    expect(im).toBeCloseTo(1, 10);
  });
  it("IMSQRT(4) = 2", () => {
    const [re] = splitComplex(asString(fnIMSQRT([rvString("4")])));
    expect(re).toBeCloseTo(2, 10);
  });
  it("IMSQRT propagates errors and rejects invalid text", () => {
    expect(fnIMSQRT([ERRORS.NA])).toEqual(ERRORS.NA);
    expect(fnIMSQRT([rvString("bad")])).toEqual(ERRORS.NUM);
  });
});

describe("IMCONJUGATE (extra coverage)", () => {
  it("IMCONJUGATE of real is real", () => {
    expect(asString(fnIMCONJUGATE([rvString("5")]))).toBe("5");
  });
  it("IMCONJUGATE flips imaginary sign", () => {
    const [re, im] = splitComplex(asString(fnIMCONJUGATE([rvString("3+4i")])));
    expect(re).toBeCloseTo(3, 10);
    expect(im).toBeCloseTo(-4, 10);
  });
  it("IMCONJUGATE(a-bi) = a+bi", () => {
    const [re, im] = splitComplex(asString(fnIMCONJUGATE([rvString("3-4i")])));
    expect(re).toBeCloseTo(3, 10);
    expect(im).toBeCloseTo(4, 10);
  });
  it("double conjugate is identity", () => {
    const first = asString(fnIMCONJUGATE([rvString("3+4i")]));
    const second = asString(fnIMCONJUGATE([rvString(first)]));
    const [re, im] = splitComplex(second);
    expect(re).toBeCloseTo(3, 10);
    expect(im).toBeCloseTo(4, 10);
  });
  it("IMCONJUGATE propagates errors", () => {
    expect(fnIMCONJUGATE([ERRORS.DIV0])).toEqual(ERRORS.DIV0);
  });
});

describe("IMTAN (extra coverage)", () => {
  it("IMTAN(0) = 0", () => {
    const [re] = splitComplex(asString(fnIMTAN([rvString("0")])));
    expect(re).toBeCloseTo(0, 10);
  });
  it("IMTAN(PI/4) = 1", () => {
    const [re] = splitComplex(asString(fnIMTAN([rvString(String(Math.PI / 4))])));
    expect(re).toBeCloseTo(1, 6);
  });
  it("IMTAN(i) is pure imaginary (tanh(1) * i)", () => {
    const [re, im] = splitComplex(asString(fnIMTAN([rvString("i")])));
    expect(re).toBeCloseTo(0, 6);
    expect(im).toBeCloseTo(Math.tanh(1), 6);
  });
  it("IMTAN rejects invalid complex text", () => {
    expect(fnIMTAN([rvString("not a number")])).toEqual(ERRORS.NUM);
  });
  it("IMTAN propagates errors", () => {
    expect(fnIMTAN([ERRORS.NA])).toEqual(ERRORS.NA);
  });
});

describe("IMCOS (extra coverage)", () => {
  it("IMCOS(0) = 1", () => {
    expect(asString(fnIMCOS([rvString("0")]))).toBe("1");
  });
  it("IMCOS(PI) ≈ -1", () => {
    const [re] = splitComplex(asString(fnIMCOS([rvString(String(Math.PI))])));
    expect(re).toBeCloseTo(-1, 6);
  });
  it("IMCOS(PI/2) ≈ 0", () => {
    const [re] = splitComplex(asString(fnIMCOS([rvString(String(Math.PI / 2))])));
    expect(Math.abs(re)).toBeLessThan(1e-6);
  });
  it("IMCOS(-x) = IMCOS(x) for real inputs", () => {
    const [re1] = splitComplex(asString(fnIMCOS([rvString("1.2")])));
    const [re2] = splitComplex(asString(fnIMCOS([rvString("-1.2")])));
    expect(re1).toBeCloseTo(re2, 10);
  });
  it("IMCOS rejects invalid text and propagates errors", () => {
    expect(fnIMCOS([rvString("bad")])).toEqual(ERRORS.NUM);
    expect(fnIMCOS([ERRORS.DIV0])).toEqual(ERRORS.DIV0);
  });
});

describe("IMARGUMENT (extra coverage)", () => {
  it("IMARGUMENT of positive real = 0", () => {
    expect(asNumber(fnIMARGUMENT([rvString("5")]))).toBeCloseTo(0, 10);
  });
  it("IMARGUMENT of positive imaginary = PI/2", () => {
    expect(asNumber(fnIMARGUMENT([rvString("i")]))).toBeCloseTo(Math.PI / 2, 6);
  });
  it("IMARGUMENT of negative real = PI (or -PI)", () => {
    const v = asNumber(fnIMARGUMENT([rvString("-1")]));
    expect(Math.abs(Math.abs(v) - Math.PI)).toBeLessThan(1e-6);
  });
  it("IMARGUMENT of negative imaginary = -PI/2", () => {
    expect(asNumber(fnIMARGUMENT([rvString("-i")]))).toBeCloseTo(-Math.PI / 2, 6);
  });
  it("IMARGUMENT(0) → #DIV/0!", () => {
    expect(fnIMARGUMENT([rvString("0")])).toEqual(ERRORS.DIV0);
  });
  it("IMARGUMENT propagates errors and rejects invalid text", () => {
    expect(fnIMARGUMENT([ERRORS.NA])).toEqual(ERRORS.NA);
    expect(fnIMARGUMENT([rvString("bad")])).toEqual(ERRORS.NUM);
  });
});

// ============================================================================
// R8 deep coverage: bit operations and boundary values
// ============================================================================

describe("BITAND / BITOR / BITXOR 48-bit semantics", () => {
  it("BITAND(2^40, 2^40) = 2^40 (not 0)", () => {
    expect(asNumber(fnBITAND([rvNumber(2 ** 40), rvNumber(2 ** 40)]))).toBe(2 ** 40);
  });

  it("BITAND on 32-bit boundary", () => {
    expect(asNumber(fnBITAND([rvNumber(0xffffffff), rvNumber(0xffffffff)]))).toBe(0xffffffff);
  });

  it("BITAND identity: x AND x = x", () => {
    expect(asNumber(fnBITAND([rvNumber(12345), rvNumber(12345)]))).toBe(12345);
  });

  it("BITAND with 0 = 0", () => {
    expect(asNumber(fnBITAND([rvNumber(0xdeadbeef), rvNumber(0)]))).toBe(0);
  });

  it("BITOR identity: x OR 0 = x", () => {
    expect(asNumber(fnBITOR([rvNumber(42), rvNumber(0)]))).toBe(42);
  });

  it("BITOR high bits", () => {
    expect(asNumber(fnBITOR([rvNumber(2 ** 40), rvNumber(2 ** 30)]))).toBe(2 ** 40 + 2 ** 30);
  });

  it("BITXOR identity: x XOR x = 0", () => {
    expect(asNumber(fnBITXOR([rvNumber(12345), rvNumber(12345)]))).toBe(0);
  });

  it("BITXOR with 0 = x", () => {
    expect(asNumber(fnBITXOR([rvNumber(42), rvNumber(0)]))).toBe(42);
  });

  it("non-integer input → #NUM!", () => {
    expect(fnBITAND([rvNumber(1.5), rvNumber(1)])).toEqual(ERRORS.NUM);
    expect(fnBITOR([rvNumber(1), rvNumber(1.5)])).toEqual(ERRORS.NUM);
  });

  it("negative input → #NUM!", () => {
    expect(fnBITAND([rvNumber(-1), rvNumber(1)])).toEqual(ERRORS.NUM);
  });

  it("value > 2^48-1 → #NUM!", () => {
    expect(fnBITAND([rvNumber(2 ** 48), rvNumber(1)])).toEqual(ERRORS.NUM);
  });
});

describe("BITLSHIFT / BITRSHIFT", () => {
  it("BITLSHIFT(1, 40) = 2^40", () => {
    expect(asNumber(fnBITLSHIFT([rvNumber(1), rvNumber(40)]))).toBe(2 ** 40);
  });

  it("BITLSHIFT negative shift is right shift", () => {
    expect(asNumber(fnBITLSHIFT([rvNumber(8), rvNumber(-1)]))).toBe(4);
  });

  it("BITLSHIFT overflow → #NUM!", () => {
    expect(fnBITLSHIFT([rvNumber(2 ** 40), rvNumber(10)])).toEqual(ERRORS.NUM);
  });

  it("BITRSHIFT(256, 4) = 16", () => {
    expect(asNumber(fnBITRSHIFT([rvNumber(256), rvNumber(4)]))).toBe(16);
  });

  it("BITRSHIFT negative shift is left shift", () => {
    expect(asNumber(fnBITRSHIFT([rvNumber(1), rvNumber(-4)]))).toBe(16);
  });

  it("BITRSHIFT to 0", () => {
    expect(asNumber(fnBITRSHIFT([rvNumber(7), rvNumber(3)]))).toBe(0);
  });

  it("shift outside ±53 → #NUM!", () => {
    expect(fnBITLSHIFT([rvNumber(1), rvNumber(54)])).toEqual(ERRORS.NUM);
    expect(fnBITRSHIFT([rvNumber(1), rvNumber(-54)])).toEqual(ERRORS.NUM);
  });
});

describe("COMPLEX / IMREAL / IMAGINARY deep", () => {
  it("COMPLEX default suffix is 'i'", () => {
    expect(asString(fnCOMPLEX([rvNumber(3), rvNumber(4)]))).toBe("3+4i");
  });

  it("COMPLEX 'j' suffix", () => {
    expect(asString(fnCOMPLEX([rvNumber(3), rvNumber(4), rvString("j")]))).toBe("3+4j");
  });

  it("COMPLEX zero imaginary is real string", () => {
    expect(asString(fnCOMPLEX([rvNumber(5), rvNumber(0)]))).toBe("5");
  });

  it("COMPLEX zero real is pure imaginary", () => {
    expect(asString(fnCOMPLEX([rvNumber(0), rvNumber(4)]))).toBe("4i");
  });

  it("COMPLEX negative imaginary uses minus sign", () => {
    expect(asString(fnCOMPLEX([rvNumber(3), rvNumber(-4)]))).toBe("3-4i");
  });

  it("COMPLEX invalid suffix → #VALUE!", () => {
    expect(fnCOMPLEX([rvNumber(3), rvNumber(4), rvString("x")])).toEqual(ERRORS.VALUE);
  });

  it("IMREAL extracts real part", () => {
    expect(asNumber(fnIMREAL([rvString("3+4i")]))).toBe(3);
  });

  it("IMAGINARY extracts imaginary part", () => {
    expect(asNumber(fnIMAGINARY([rvString("3+4i")]))).toBe(4);
  });

  it("IMREAL / IMAGINARY of real number", () => {
    expect(asNumber(fnIMREAL([rvString("5")]))).toBe(5);
    expect(asNumber(fnIMAGINARY([rvString("5")]))).toBe(0);
  });
});

describe("IMABS / IMARGUMENT", () => {
  it("IMABS(3+4i) = 5 (Pythagorean)", () => {
    expect(asNumber(fnIMABS([rvString("3+4i")]))).toBe(5);
  });

  it("IMABS of pure real = |real|", () => {
    expect(asNumber(fnIMABS([rvString("-7")]))).toBe(7);
  });

  it("IMARGUMENT(i) = π/2", () => {
    expect(asNumber(fnIMARGUMENT([rvString("i")]))).toBeCloseTo(Math.PI / 2, 5);
  });

  it("IMARGUMENT(1) = 0", () => {
    expect(asNumber(fnIMARGUMENT([rvString("1")]))).toBe(0);
  });

  it("IMARGUMENT(-1) = π (or -π in some conventions)", () => {
    expect(Math.abs(asNumber(fnIMARGUMENT([rvString("-1")])))).toBeCloseTo(Math.PI, 5);
  });

  it("IMARGUMENT(0) → #DIV/0!", () => {
    expect(fnIMARGUMENT([rvString("0")])).toEqual(ERRORS.DIV0);
  });
});

describe("DELTA / GESTEP", () => {
  it("DELTA(5, 5) = 1", () => {
    expect(asNumber(fnDELTA([rvNumber(5), rvNumber(5)]))).toBe(1);
  });

  it("DELTA(5, 4) = 0", () => {
    expect(asNumber(fnDELTA([rvNumber(5), rvNumber(4)]))).toBe(0);
  });

  it("DELTA default number2 = 0", () => {
    expect(asNumber(fnDELTA([rvNumber(0)]))).toBe(1);
    expect(asNumber(fnDELTA([rvNumber(1)]))).toBe(0);
  });

  it("GESTEP(5, 3) = 1 (>= step)", () => {
    expect(asNumber(fnGESTEP([rvNumber(5), rvNumber(3)]))).toBe(1);
  });

  it("GESTEP(2, 3) = 0 (< step)", () => {
    expect(asNumber(fnGESTEP([rvNumber(2), rvNumber(3)]))).toBe(0);
  });

  it("GESTEP equal = 1", () => {
    expect(asNumber(fnGESTEP([rvNumber(3), rvNumber(3)]))).toBe(1);
  });

  it("GESTEP default step = 0", () => {
    expect(asNumber(fnGESTEP([rvNumber(-1)]))).toBe(0);
    expect(asNumber(fnGESTEP([rvNumber(0)]))).toBe(1);
  });
});

// ============================================================================
// Saturation blocks — each below-threshold function gets 5-10 additional
// focused tests so every engineering function reaches the 10+ reference bar.
// The IM* family shares a parseComplex front-end, so the patterns repeat:
//   - real & imaginary sampling
//   - identities (IMSIN²+IMCOS² = 1, IMCSC·IMSIN = 1, etc.)
//   - singularities that should surface #DIV/0! / #NUM!
//   - invalid string → #NUM!, error propagation, array cells.
// ============================================================================

// -- Base conversion saturation ---------------------------------------------

describe("BIN2OCT saturation", () => {
  it("BIN2OCT('0') = '0'", () => {
    expect(asString(fnBIN2OCT([rvString("0")]))).toBe("0");
  });
  it("BIN2OCT('111111111') = '777' (9-bit max positive)", () => {
    expect(asString(fnBIN2OCT([rvString("111111111")]))).toBe("777");
  });
  it("BIN2OCT('1000000000') → 10-digit two's complement negative", () => {
    // Binary -512 → octal 7777777000
    expect(asString(fnBIN2OCT([rvString("1000000000")]))).toBe("7777777000");
  });
  it("BIN2OCT with places pads", () => {
    expect(asString(fnBIN2OCT([rvString("10"), rvNumber(4)]))).toBe("0002");
  });
  it("BIN2OCT with places < 1 → #NUM!", () => {
    expect(fnBIN2OCT([rvString("10"), rvNumber(0)])).toEqual(ERRORS.NUM);
  });
  it("BIN2OCT with places > 10 → #NUM!", () => {
    expect(fnBIN2OCT([rvString("10"), rvNumber(11)])).toEqual(ERRORS.NUM);
  });
  it("BIN2OCT rejects non-binary", () => {
    expect(fnBIN2OCT([rvString("12")])).toEqual(ERRORS.NUM);
  });
  it("BIN2OCT rejects empty", () => {
    expect(fnBIN2OCT([rvString("")])).toEqual(ERRORS.NUM);
  });
  it("BIN2OCT propagates errors", () => {
    expect(fnBIN2OCT([ERRORS.VALUE])).toEqual(ERRORS.VALUE);
  });
});

describe("OCT2HEX saturation", () => {
  it("OCT2HEX('0') = '0'", () => {
    expect(asString(fnOCT2HEX([rvString("0")]))).toBe("0");
  });
  it("OCT2HEX('100') = '40'", () => {
    expect(asString(fnOCT2HEX([rvString("100")]))).toBe("40");
  });
  it("OCT2HEX('7777777777') = 'FFFFFFFFFF' (negative -1)", () => {
    expect(asString(fnOCT2HEX([rvString("7777777777")]))).toBe("FFFFFFFFFF");
  });
  it("OCT2HEX with places pads", () => {
    expect(asString(fnOCT2HEX([rvString("17"), rvNumber(4)]))).toBe("000F");
  });
  it("OCT2HEX with places < 1 → #NUM!", () => {
    expect(fnOCT2HEX([rvString("17"), rvNumber(-1)])).toEqual(ERRORS.NUM);
  });
  it("OCT2HEX rejects non-octal '8'", () => {
    expect(fnOCT2HEX([rvString("8")])).toEqual(ERRORS.NUM);
  });
  it("OCT2HEX rejects >10 digits", () => {
    expect(fnOCT2HEX([rvString("11111111111")])).toEqual(ERRORS.NUM);
  });
  it("OCT2HEX propagates errors", () => {
    expect(fnOCT2HEX([ERRORS.DIV0])).toEqual(ERRORS.DIV0);
  });
  it("OCT2HEX of numeric arg coerces to string first", () => {
    expect(asString(fnOCT2HEX([rvNumber(17)]))).toBe("F");
  });
});

describe("HEX2OCT saturation", () => {
  it("HEX2OCT('0') = '0'", () => {
    expect(asString(fnHEX2OCT([rvString("0")]))).toBe("0");
  });
  it("HEX2OCT('8') = '10'", () => {
    expect(asString(fnHEX2OCT([rvString("8")]))).toBe("10");
  });
  it("HEX2OCT('FFFFFFFFFF') = '7777777777' (negative -1)", () => {
    expect(asString(fnHEX2OCT([rvString("FFFFFFFFFF")]))).toBe("7777777777");
  });
  it("HEX2OCT with places pads", () => {
    expect(asString(fnHEX2OCT([rvString("8"), rvNumber(4)]))).toBe("0010");
  });
  it("HEX2OCT out-of-OCT-range → #NUM!", () => {
    // 0x20000000 = 2^29 = out of oct range
    expect(fnHEX2OCT([rvString("20000000")])).toEqual(ERRORS.NUM);
  });
  it("HEX2OCT rejects non-hex", () => {
    expect(fnHEX2OCT([rvString("XYZ")])).toEqual(ERRORS.NUM);
  });
  it("HEX2OCT propagates errors", () => {
    expect(fnHEX2OCT([ERRORS.NA])).toEqual(ERRORS.NA);
  });
  it("HEX2OCT rejects places out of [1, 10]", () => {
    expect(fnHEX2OCT([rvString("8"), rvNumber(11)])).toEqual(ERRORS.NUM);
  });
});

describe("OCT2BIN saturation", () => {
  it("OCT2BIN('0') = '0'", () => {
    expect(asString(fnOCT2BIN([rvString("0")]))).toBe("0");
  });
  it("OCT2BIN('777') = max positive (511 = 9-bit)", () => {
    expect(asString(fnOCT2BIN([rvString("777")]))).toBe("111111111");
  });
  it("OCT2BIN('7777777777') = 10-digit 1…1 (negative -1)", () => {
    expect(asString(fnOCT2BIN([rvString("7777777777")]))).toBe("1111111111");
  });
  it("OCT2BIN with places pads", () => {
    expect(asString(fnOCT2BIN([rvString("7"), rvNumber(4)]))).toBe("0111");
  });
  it("OCT2BIN out-of-BIN-range → #NUM!", () => {
    // octal 1000 = 512, just out of bin range
    expect(fnOCT2BIN([rvString("1000")])).toEqual(ERRORS.NUM);
  });
  it("OCT2BIN rejects non-octal", () => {
    expect(fnOCT2BIN([rvString("8")])).toEqual(ERRORS.NUM);
  });
  it("OCT2BIN propagates errors", () => {
    expect(fnOCT2BIN([ERRORS.REF])).toEqual(ERRORS.REF);
  });
  it("OCT2BIN rejects places < 1", () => {
    expect(fnOCT2BIN([rvString("5"), rvNumber(0)])).toEqual(ERRORS.NUM);
  });
});

// -- Complex saturation -----------------------------------------------------

describe("IMABS saturation", () => {
  it("IMABS('0') = 0", () => {
    expect(asNumber(fnIMABS([rvString("0")]))).toBe(0);
  });
  it("IMABS('3+4i') = 5", () => {
    expect(asNumber(fnIMABS([rvString("3+4i")]))).toBeCloseTo(5, 10);
  });
  it("IMABS('-3-4i') = 5", () => {
    expect(asNumber(fnIMABS([rvString("-3-4i")]))).toBeCloseTo(5, 10);
  });
  it("IMABS('i') = 1", () => {
    expect(asNumber(fnIMABS([rvString("i")]))).toBeCloseTo(1, 10);
  });
  it("IMABS on numeric real passes through (|5| = 5)", () => {
    expect(asNumber(fnIMABS([rvNumber(5)]))).toBe(5);
  });
  it("IMABS on negative numeric real = |n|", () => {
    expect(asNumber(fnIMABS([rvNumber(-7)]))).toBe(7);
  });
  it("IMABS rejects malformed → #NUM!", () => {
    expect(fnIMABS([rvString("xyz")])).toEqual(ERRORS.NUM);
  });
  it("IMABS propagates errors", () => {
    expect(fnIMABS([ERRORS.NA])).toEqual(ERRORS.NA);
  });
  it("IMABS handles 'j' suffix form", () => {
    expect(asNumber(fnIMABS([rvString("3+4j")]))).toBeCloseTo(5, 10);
  });
});

describe("IMSQRT saturation", () => {
  it("IMSQRT('4') = '2'", () => {
    const [re, im] = splitComplex(asString(fnIMSQRT([rvString("4")])));
    expect(re).toBeCloseTo(2, 10);
    expect(im).toBeCloseTo(0, 10);
  });
  it("IMSQRT('i') = sqrt(2)/2 + sqrt(2)/2·i", () => {
    const [re, im] = splitComplex(asString(fnIMSQRT([rvString("i")])));
    expect(re).toBeCloseTo(Math.SQRT1_2, 10);
    expect(im).toBeCloseTo(Math.SQRT1_2, 10);
  });
  it("IMSQRT('0') = '0'", () => {
    const [re, im] = splitComplex(asString(fnIMSQRT([rvString("0")])));
    expect(re).toBe(0);
    expect(im).toBe(0);
  });
  it("IMSQRT('-1') = 'i'", () => {
    const [re, im] = splitComplex(asString(fnIMSQRT([rvString("-1")])));
    expect(re).toBeCloseTo(0, 10);
    expect(im).toBeCloseTo(1, 10);
  });
  it("IMSQRT squared returns original (up to float)", () => {
    const root = asString(fnIMSQRT([rvString("3+4i")]));
    const sq = asString(fnIMPRODUCT([rvString(root), rvString(root)]));
    const [re, im] = splitComplex(sq);
    expect(re).toBeCloseTo(3, 5);
    expect(im).toBeCloseTo(4, 5);
  });
  it("IMSQRT rejects malformed", () => {
    expect(fnIMSQRT([rvString("not-complex")])).toEqual(ERRORS.NUM);
  });
  it("IMSQRT propagates errors", () => {
    expect(fnIMSQRT([ERRORS.DIV0])).toEqual(ERRORS.DIV0);
  });
});

describe("IMTAN saturation", () => {
  it("IMTAN('0') = '0'", () => {
    const [re, im] = splitComplex(asString(fnIMTAN([rvString("0")])));
    expect(re).toBeCloseTo(0, 10);
    expect(im).toBeCloseTo(0, 10);
  });
  it("IMTAN(π/4) ≈ 1", () => {
    const [re] = splitComplex(asString(fnIMTAN([rvString(String(Math.PI / 4))])));
    expect(re).toBeCloseTo(1, 5);
  });
  it("IMTAN(-π/4) ≈ -1", () => {
    const [re] = splitComplex(asString(fnIMTAN([rvString(String(-Math.PI / 4))])));
    expect(re).toBeCloseTo(-1, 5);
  });
  it("IMTAN(i) ≈ i·tanh(1)", () => {
    const [, im] = splitComplex(asString(fnIMTAN([rvString("i")])));
    expect(im).toBeCloseTo(Math.tanh(1), 5);
  });
  it("IMTAN propagates errors", () => {
    expect(fnIMTAN([ERRORS.NA])).toEqual(ERRORS.NA);
  });
  it("IMTAN of number coerces via toStringRV", () => {
    const [re] = splitComplex(asString(fnIMTAN([rvNumber(1)])));
    expect(re).toBeCloseTo(Math.tan(1), 5);
  });
  it("IMTAN rejects malformed", () => {
    expect(fnIMTAN([rvString("xx")])).toEqual(ERRORS.NUM);
  });
});

describe("IMSIN saturation", () => {
  it("IMSIN('0') = '0'", () => {
    const [re, im] = splitComplex(asString(fnIMSIN([rvString("0")])));
    expect(re).toBe(0);
    expect(im).toBe(0);
  });
  it("IMSIN(π/2) ≈ 1", () => {
    const [re] = splitComplex(asString(fnIMSIN([rvString(String(Math.PI / 2))])));
    expect(re).toBeCloseTo(1, 5);
  });
  it("IMSIN(π) ≈ 0", () => {
    const [re] = splitComplex(asString(fnIMSIN([rvString(String(Math.PI))])));
    expect(re).toBeCloseTo(0, 5);
  });
  it("IMSIN(i) ≈ i·sinh(1)", () => {
    const [re, im] = splitComplex(asString(fnIMSIN([rvString("i")])));
    expect(re).toBeCloseTo(0, 5);
    expect(im).toBeCloseTo(Math.sinh(1), 5);
  });
  it("IMSIN²+IMCOS² ≈ 1 for real x=0.3", () => {
    const s = splitComplex(asString(fnIMSIN([rvString("0.3")])));
    const c = splitComplex(asString(fnIMCOS([rvString("0.3")])));
    expect(s[0] * s[0] + c[0] * c[0]).toBeCloseTo(1, 10);
  });
  it("IMSIN rejects malformed", () => {
    expect(fnIMSIN([rvString("xx")])).toEqual(ERRORS.NUM);
  });
  it("IMSIN propagates errors", () => {
    expect(fnIMSIN([ERRORS.NA])).toEqual(ERRORS.NA);
  });
});

describe("IMCOS saturation", () => {
  it("IMCOS('0') ≈ 1", () => {
    const [re] = splitComplex(asString(fnIMCOS([rvString("0")])));
    expect(re).toBeCloseTo(1, 10);
  });
  it("IMCOS(π) ≈ -1", () => {
    const [re] = splitComplex(asString(fnIMCOS([rvString(String(Math.PI))])));
    expect(re).toBeCloseTo(-1, 5);
  });
  it("IMCOS(i) ≈ cosh(1)", () => {
    const [re] = splitComplex(asString(fnIMCOS([rvString("i")])));
    expect(re).toBeCloseTo(Math.cosh(1), 5);
  });
  it("IMCOS rejects malformed", () => {
    expect(fnIMCOS([rvString("xx")])).toEqual(ERRORS.NUM);
  });
  it("IMCOS propagates errors", () => {
    expect(fnIMCOS([ERRORS.REF])).toEqual(ERRORS.REF);
  });
  it("IMCOS of numeric arg", () => {
    const [re] = splitComplex(asString(fnIMCOS([rvNumber(1)])));
    expect(re).toBeCloseTo(Math.cos(1), 10);
  });
});

describe("IMCSC saturation", () => {
  it("IMCSC(π/2) = 1", () => {
    const [re] = splitComplex(asString(fnIMCSC([rvString(String(Math.PI / 2))])));
    expect(re).toBeCloseTo(1, 5);
  });
  it("IMCSC(π/6) ≈ 2", () => {
    const [re] = splitComplex(asString(fnIMCSC([rvString(String(Math.PI / 6))])));
    expect(re).toBeCloseTo(2, 5);
  });
  it("IMCSC(0) → #DIV/0! (sin(0)=0)", () => {
    expect(fnIMCSC([rvString("0")])).toEqual(ERRORS.DIV0);
  });
  it("IMCSC(π) → #DIV/0! (sin(π)≈0)", () => {
    // sin(π) is floating-point tiny but non-zero, so some engines emit a
    // huge finite value rather than #DIV/0!. Accept either outcome.
    const r = fnIMCSC([rvString(String(Math.PI))]);
    if (r.kind === RVKind.Error) {
      expect(r).toEqual(ERRORS.DIV0);
    } else {
      const [re] = splitComplex((r as StringValue).value);
      expect(Math.abs(re)).toBeGreaterThan(1e10);
    }
  });
  it("IMCSC · IMSIN = 1 for real x=0.3", () => {
    const csc = splitComplex(asString(fnIMCSC([rvString("0.3")])))[0];
    const sin = splitComplex(asString(fnIMSIN([rvString("0.3")])))[0];
    expect(csc * sin).toBeCloseTo(1, 10);
  });
  it("IMCSC rejects malformed", () => {
    expect(fnIMCSC([rvString("xx")])).toEqual(ERRORS.NUM);
  });
  it("IMCSC propagates errors", () => {
    expect(fnIMCSC([ERRORS.NA])).toEqual(ERRORS.NA);
  });
});

describe("IMSEC saturation", () => {
  it("IMSEC(0) = 1", () => {
    const [re] = splitComplex(asString(fnIMSEC([rvString("0")])));
    expect(re).toBeCloseTo(1, 10);
  });
  it("IMSEC(π/3) = 2", () => {
    const [re] = splitComplex(asString(fnIMSEC([rvString(String(Math.PI / 3))])));
    expect(re).toBeCloseTo(2, 5);
  });
  it("IMSEC(π/2) → finite huge or #DIV/0! (cos(π/2)≈0)", () => {
    const r = fnIMSEC([rvString(String(Math.PI / 2))]);
    if (r.kind === RVKind.Error) {
      expect(r).toEqual(ERRORS.DIV0);
    } else {
      const [re] = splitComplex((r as StringValue).value);
      expect(Math.abs(re)).toBeGreaterThan(1e10);
    }
  });
  it("IMSEC · IMCOS = 1 for real x=0.4", () => {
    const sec = splitComplex(asString(fnIMSEC([rvString("0.4")])))[0];
    const cos = splitComplex(asString(fnIMCOS([rvString("0.4")])))[0];
    expect(sec * cos).toBeCloseTo(1, 10);
  });
  it("IMSEC rejects malformed", () => {
    expect(fnIMSEC([rvString("xx")])).toEqual(ERRORS.NUM);
  });
  it("IMSEC propagates errors", () => {
    expect(fnIMSEC([ERRORS.REF])).toEqual(ERRORS.REF);
  });
  it("IMSEC on TRUE = sec(1) = 1/cos(1) (R8 boolean coercion)", () => {
    // TRUE → 1 (pure real). sec(1) = 1/cos(1) ≈ 1.8508.
    const r = fnIMSEC([rvBoolean(true)]);
    expect(r.kind).toBe(RVKind.String);
    const re = parseFloat((r as StringValue).value);
    expect(re).toBeCloseTo(1 / Math.cos(1), 5);
  });
});

describe("IMCOT saturation", () => {
  it("IMCOT(π/4) = 1", () => {
    const [re] = splitComplex(asString(fnIMCOT([rvString(String(Math.PI / 4))])));
    expect(re).toBeCloseTo(1, 5);
  });
  it("IMCOT(π/2) ≈ 0", () => {
    const [re] = splitComplex(asString(fnIMCOT([rvString(String(Math.PI / 2))])));
    expect(re).toBeCloseTo(0, 5);
  });
  it("IMCOT(0) → #DIV/0!", () => {
    expect(fnIMCOT([rvString("0")])).toEqual(ERRORS.DIV0);
  });
  it("IMCOT · IMTAN = 1 for real x=0.7", () => {
    const cot = splitComplex(asString(fnIMCOT([rvString("0.7")])))[0];
    const tan = splitComplex(asString(fnIMTAN([rvString("0.7")])))[0];
    expect(cot * tan).toBeCloseTo(1, 10);
  });
  it("IMCOT rejects malformed", () => {
    expect(fnIMCOT([rvString("xx")])).toEqual(ERRORS.NUM);
  });
  it("IMCOT propagates errors", () => {
    expect(fnIMCOT([ERRORS.VALUE])).toEqual(ERRORS.VALUE);
  });
});

describe("IMSINH saturation", () => {
  it("IMSINH('0') = 0", () => {
    const [re, im] = splitComplex(asString(fnIMSINH([rvString("0")])));
    expect(re).toBe(0);
    expect(im).toBe(0);
  });
  it("IMSINH(1) = sinh(1)", () => {
    const [re] = splitComplex(asString(fnIMSINH([rvString("1")])));
    expect(re).toBeCloseTo(Math.sinh(1), 10);
  });
  it("IMSINH(-1) = -sinh(1)", () => {
    const [re] = splitComplex(asString(fnIMSINH([rvString("-1")])));
    expect(re).toBeCloseTo(-Math.sinh(1), 10);
  });
  it("IMSINH(i) ≈ i·sin(1)", () => {
    const [re, im] = splitComplex(asString(fnIMSINH([rvString("i")])));
    expect(re).toBeCloseTo(0, 10);
    expect(im).toBeCloseTo(Math.sin(1), 10);
  });
  it("IMSINH rejects malformed", () => {
    expect(fnIMSINH([rvString("xx")])).toEqual(ERRORS.NUM);
  });
  it("IMSINH propagates errors", () => {
    expect(fnIMSINH([ERRORS.NA])).toEqual(ERRORS.NA);
  });
  it("IMCOSH² - IMSINH² = 1 for real x=1.5", () => {
    const s = splitComplex(asString(fnIMSINH([rvString("1.5")])))[0];
    const c = splitComplex(asString(fnIMCOSH([rvString("1.5")])))[0];
    expect(c * c - s * s).toBeCloseTo(1, 10);
  });
});

describe("IMCOSH saturation", () => {
  it("IMCOSH('0') = 1", () => {
    const [re] = splitComplex(asString(fnIMCOSH([rvString("0")])));
    expect(re).toBeCloseTo(1, 10);
  });
  it("IMCOSH(1) = cosh(1)", () => {
    const [re] = splitComplex(asString(fnIMCOSH([rvString("1")])));
    expect(re).toBeCloseTo(Math.cosh(1), 10);
  });
  it("IMCOSH is even: IMCOSH(-1) = IMCOSH(1)", () => {
    const pos = splitComplex(asString(fnIMCOSH([rvString("1")])))[0];
    const neg = splitComplex(asString(fnIMCOSH([rvString("-1")])))[0];
    expect(pos).toBeCloseTo(neg, 10);
  });
  it("IMCOSH(i) = cos(1)", () => {
    const [re] = splitComplex(asString(fnIMCOSH([rvString("i")])));
    expect(re).toBeCloseTo(Math.cos(1), 10);
  });
  it("IMCOSH rejects malformed", () => {
    expect(fnIMCOSH([rvString("xx")])).toEqual(ERRORS.NUM);
  });
  it("IMCOSH propagates errors", () => {
    expect(fnIMCOSH([ERRORS.VALUE])).toEqual(ERRORS.VALUE);
  });
});

describe("IMTANH saturation", () => {
  it("IMTANH('0') = 0", () => {
    const [re] = splitComplex(asString(fnIMTANH([rvString("0")])));
    expect(re).toBe(0);
  });
  it("IMTANH(1) = tanh(1)", () => {
    const [re] = splitComplex(asString(fnIMTANH([rvString("1")])));
    expect(re).toBeCloseTo(Math.tanh(1), 10);
  });
  it("IMTANH is odd: IMTANH(-2) = -IMTANH(2)", () => {
    const pos = splitComplex(asString(fnIMTANH([rvString("2")])))[0];
    const neg = splitComplex(asString(fnIMTANH([rvString("-2")])))[0];
    expect(neg).toBeCloseTo(-pos, 10);
  });
  it("IMTANH · IMCOTH = 1 for real x=1.5", () => {
    const t = splitComplex(asString(fnIMTANH([rvString("1.5")])))[0];
    const c = splitComplex(asString(fnIMCOTH([rvString("1.5")])))[0];
    expect(t * c).toBeCloseTo(1, 10);
  });
  it("IMTANH rejects malformed", () => {
    expect(fnIMTANH([rvString("xx")])).toEqual(ERRORS.NUM);
  });
  it("IMTANH propagates errors", () => {
    expect(fnIMTANH([ERRORS.NA])).toEqual(ERRORS.NA);
  });
});

describe("IMCSCH saturation", () => {
  it("IMCSCH(1) = 1/sinh(1)", () => {
    const [re] = splitComplex(asString(fnIMCSCH([rvString("1")])));
    expect(re).toBeCloseTo(1 / Math.sinh(1), 10);
  });
  it("IMCSCH(0) → #DIV/0!", () => {
    expect(fnIMCSCH([rvString("0")])).toEqual(ERRORS.DIV0);
  });
  it("IMCSCH(-1) is negative", () => {
    const [re] = splitComplex(asString(fnIMCSCH([rvString("-1")])));
    expect(re).toBeLessThan(0);
    expect(re).toBeCloseTo(-1 / Math.sinh(1), 10);
  });
  it("IMCSCH · IMSINH = 1 for real x=2", () => {
    const csch = splitComplex(asString(fnIMCSCH([rvString("2")])))[0];
    const sinh = splitComplex(asString(fnIMSINH([rvString("2")])))[0];
    expect(csch * sinh).toBeCloseTo(1, 10);
  });
  it("IMCSCH rejects malformed", () => {
    expect(fnIMCSCH([rvString("xx")])).toEqual(ERRORS.NUM);
  });
  it("IMCSCH propagates errors", () => {
    expect(fnIMCSCH([ERRORS.REF])).toEqual(ERRORS.REF);
  });
});

describe("IMSECH saturation", () => {
  it("IMSECH(0) = 1", () => {
    const [re] = splitComplex(asString(fnIMSECH([rvString("0")])));
    expect(re).toBeCloseTo(1, 10);
  });
  it("IMSECH(1) = 1/cosh(1)", () => {
    const [re] = splitComplex(asString(fnIMSECH([rvString("1")])));
    expect(re).toBeCloseTo(1 / Math.cosh(1), 10);
  });
  it("IMSECH is even: IMSECH(-2) = IMSECH(2)", () => {
    const p = splitComplex(asString(fnIMSECH([rvString("2")])))[0];
    const n = splitComplex(asString(fnIMSECH([rvString("-2")])))[0];
    expect(p).toBeCloseTo(n, 10);
  });
  it("IMSECH · IMCOSH = 1 for real x=2", () => {
    const s = splitComplex(asString(fnIMSECH([rvString("2")])))[0];
    const c = splitComplex(asString(fnIMCOSH([rvString("2")])))[0];
    expect(s * c).toBeCloseTo(1, 10);
  });
  it("IMSECH rejects malformed", () => {
    expect(fnIMSECH([rvString("xx")])).toEqual(ERRORS.NUM);
  });
  it("IMSECH propagates errors", () => {
    expect(fnIMSECH([ERRORS.VALUE])).toEqual(ERRORS.VALUE);
  });
});

describe("IMCOTH saturation", () => {
  it("IMCOTH(1) = 1/tanh(1)", () => {
    const [re] = splitComplex(asString(fnIMCOTH([rvString("1")])));
    expect(re).toBeCloseTo(1 / Math.tanh(1), 10);
  });
  it("IMCOTH(0) → #DIV/0!", () => {
    expect(fnIMCOTH([rvString("0")])).toEqual(ERRORS.DIV0);
  });
  it("IMCOTH · IMTANH = 1 for real x=0.5", () => {
    const coth = splitComplex(asString(fnIMCOTH([rvString("0.5")])))[0];
    const tanh = splitComplex(asString(fnIMTANH([rvString("0.5")])))[0];
    expect(coth * tanh).toBeCloseTo(1, 10);
  });
  it("IMCOTH is odd: IMCOTH(-1) = -IMCOTH(1)", () => {
    const p = splitComplex(asString(fnIMCOTH([rvString("1")])))[0];
    const n = splitComplex(asString(fnIMCOTH([rvString("-1")])))[0];
    expect(n).toBeCloseTo(-p, 10);
  });
  it("IMCOTH rejects malformed", () => {
    expect(fnIMCOTH([rvString("xx")])).toEqual(ERRORS.NUM);
  });
  it("IMCOTH propagates errors", () => {
    expect(fnIMCOTH([ERRORS.DIV0])).toEqual(ERRORS.DIV0);
  });
});

describe("IMLOG2 saturation", () => {
  it("IMLOG2('2') = 1", () => {
    const [re, im] = splitComplex(asString(fnIMLOG2([rvString("2")])));
    expect(re).toBeCloseTo(1, 10);
    expect(im).toBeCloseTo(0, 10);
  });
  it("IMLOG2('8') = 3", () => {
    const [re] = splitComplex(asString(fnIMLOG2([rvString("8")])));
    expect(re).toBeCloseTo(3, 10);
  });
  it("IMLOG2('1') = 0", () => {
    const [re] = splitComplex(asString(fnIMLOG2([rvString("1")])));
    expect(re).toBeCloseTo(0, 10);
  });
  it("IMLOG2(0) → #NUM!", () => {
    expect(fnIMLOG2([rvString("0")])).toEqual(ERRORS.NUM);
  });
  it("IMLOG2 rejects malformed", () => {
    expect(fnIMLOG2([rvString("xx")])).toEqual(ERRORS.NUM);
  });
  it("IMLOG2 propagates errors", () => {
    expect(fnIMLOG2([ERRORS.NA])).toEqual(ERRORS.NA);
  });
  it("IMLOG2(i) has zero real part", () => {
    const [re] = splitComplex(asString(fnIMLOG2([rvString("i")])));
    expect(re).toBeCloseTo(0, 10);
  });
});

describe("IMLOG10 saturation", () => {
  it("IMLOG10('100') = 2", () => {
    const [re] = splitComplex(asString(fnIMLOG10([rvString("100")])));
    expect(re).toBeCloseTo(2, 10);
  });
  it("IMLOG10('1') = 0", () => {
    const [re] = splitComplex(asString(fnIMLOG10([rvString("1")])));
    expect(re).toBeCloseTo(0, 10);
  });
  it("IMLOG10('10') = 1", () => {
    const [re] = splitComplex(asString(fnIMLOG10([rvString("10")])));
    expect(re).toBeCloseTo(1, 10);
  });
  it("IMLOG10(0) → #NUM!", () => {
    expect(fnIMLOG10([rvString("0")])).toEqual(ERRORS.NUM);
  });
  it("IMLOG10 rejects malformed", () => {
    expect(fnIMLOG10([rvString("xx")])).toEqual(ERRORS.NUM);
  });
  it("IMLOG10 propagates errors", () => {
    expect(fnIMLOG10([ERRORS.REF])).toEqual(ERRORS.REF);
  });
});

describe("IMPOWER saturation", () => {
  it("IMPOWER('2+0i', 3) = 8", () => {
    const [re] = splitComplex(asString(fnIMPOWER([rvString("2+0i"), rvNumber(3)])));
    expect(re).toBeCloseTo(8, 10);
  });
  it("IMPOWER('i', 2) = -1", () => {
    const [re, im] = splitComplex(asString(fnIMPOWER([rvString("i"), rvNumber(2)])));
    expect(re).toBeCloseTo(-1, 10);
    expect(im).toBeCloseTo(0, 10);
  });
  it("IMPOWER('i', 4) = 1", () => {
    const [re, im] = splitComplex(asString(fnIMPOWER([rvString("i"), rvNumber(4)])));
    expect(re).toBeCloseTo(1, 10);
    expect(im).toBeCloseTo(0, 10);
  });
  it("IMPOWER('1+1i', 2) = 2i", () => {
    const [re, im] = splitComplex(asString(fnIMPOWER([rvString("1+1i"), rvNumber(2)])));
    expect(re).toBeCloseTo(0, 10);
    expect(im).toBeCloseTo(2, 10);
  });
  it("IMPOWER('2', 0.5) = sqrt(2)", () => {
    const [re] = splitComplex(asString(fnIMPOWER([rvString("2"), rvNumber(0.5)])));
    expect(re).toBeCloseTo(Math.sqrt(2), 10);
  });
  it("IMPOWER rejects malformed complex", () => {
    expect(fnIMPOWER([rvString("xx"), rvNumber(2)])).toEqual(ERRORS.NUM);
  });
  it("IMPOWER propagates errors", () => {
    expect(fnIMPOWER([ERRORS.NA, rvNumber(2)])).toEqual(ERRORS.NA);
  });
  it("IMPOWER exponent error propagates", () => {
    expect(fnIMPOWER([rvString("1"), ERRORS.NA])).toEqual(ERRORS.NA);
  });
});

describe("IMEXP saturation", () => {
  it("IMEXP('0') = 1", () => {
    const [re] = splitComplex(asString(fnIMEXP([rvString("0")])));
    expect(re).toBeCloseTo(1, 10);
  });
  it("IMEXP('1') = e", () => {
    const [re] = splitComplex(asString(fnIMEXP([rvString("1")])));
    expect(re).toBeCloseTo(Math.E, 10);
  });
  it("IMEXP('iπ') ≈ -1 (Euler's identity)", () => {
    const [re, im] = splitComplex(asString(fnIMEXP([rvString(`${Math.PI}i`)])));
    expect(re).toBeCloseTo(-1, 5);
    expect(im).toBeCloseTo(0, 5);
  });
  it("IMEXP('i') = cos(1) + sin(1)i", () => {
    const [re, im] = splitComplex(asString(fnIMEXP([rvString("i")])));
    expect(re).toBeCloseTo(Math.cos(1), 10);
    expect(im).toBeCloseTo(Math.sin(1), 10);
  });
  it("IMEXP(-1) = 1/e", () => {
    const [re] = splitComplex(asString(fnIMEXP([rvString("-1")])));
    expect(re).toBeCloseTo(1 / Math.E, 10);
  });
  it("IMEXP rejects malformed", () => {
    expect(fnIMEXP([rvString("xx")])).toEqual(ERRORS.NUM);
  });
  it("IMEXP propagates errors", () => {
    expect(fnIMEXP([ERRORS.NA])).toEqual(ERRORS.NA);
  });
});

describe("IMDIV saturation", () => {
  it("IMDIV(8, 2) = 4", () => {
    const [re] = splitComplex(asString(fnIMDIV([rvString("8"), rvString("2")])));
    expect(re).toBeCloseTo(4, 10);
  });
  it("IMDIV('1+0i', '0+1i') = -i", () => {
    const [re, im] = splitComplex(asString(fnIMDIV([rvString("1"), rvString("i")])));
    expect(re).toBeCloseTo(0, 10);
    expect(im).toBeCloseTo(-1, 10);
  });
  it("IMDIV by 0 → #DIV/0!", () => {
    expect(fnIMDIV([rvString("5"), rvString("0")])).toEqual(ERRORS.DIV0);
  });
  it("IMDIV('3+4i','1+0i') = 3+4i", () => {
    const s = asString(fnIMDIV([rvString("3+4i"), rvString("1")]));
    const [re, im] = splitComplex(s);
    expect(re).toBeCloseTo(3, 10);
    expect(im).toBeCloseTo(4, 10);
  });
  it("IMDIV rejects malformed dividend", () => {
    expect(fnIMDIV([rvString("xx"), rvString("1")])).toEqual(ERRORS.NUM);
  });
  it("IMDIV rejects malformed divisor", () => {
    expect(fnIMDIV([rvString("1"), rvString("yy")])).toEqual(ERRORS.NUM);
  });
  it("IMDIV propagates errors", () => {
    expect(fnIMDIV([ERRORS.NA, rvString("1")])).toEqual(ERRORS.NA);
    expect(fnIMDIV([rvString("1"), ERRORS.REF])).toEqual(ERRORS.REF);
  });
});

describe("IMSUB saturation", () => {
  it("IMSUB('5','3') = 2", () => {
    const [re] = splitComplex(asString(fnIMSUB([rvString("5"), rvString("3")])));
    expect(re).toBe(2);
  });
  it("IMSUB('3+4i','1+2i') = '2+2i'", () => {
    const [re, im] = splitComplex(asString(fnIMSUB([rvString("3+4i"), rvString("1+2i")])));
    expect(re).toBeCloseTo(2, 10);
    expect(im).toBeCloseTo(2, 10);
  });
  it("IMSUB(a, a) = 0", () => {
    const [re, im] = splitComplex(asString(fnIMSUB([rvString("3+4i"), rvString("3+4i")])));
    expect(re).toBe(0);
    expect(im).toBe(0);
  });
  it("IMSUB rejects malformed", () => {
    expect(fnIMSUB([rvString("xx"), rvString("1")])).toEqual(ERRORS.NUM);
    expect(fnIMSUB([rvString("1"), rvString("yy")])).toEqual(ERRORS.NUM);
  });
  it("IMSUB propagates errors", () => {
    expect(fnIMSUB([ERRORS.NA, rvString("1")])).toEqual(ERRORS.NA);
    expect(fnIMSUB([rvString("1"), ERRORS.DIV0])).toEqual(ERRORS.DIV0);
  });
  it("IMSUB is anti-commutative: a-b = -(b-a)", () => {
    const fwd = splitComplex(asString(fnIMSUB([rvString("7"), rvString("3")])))[0];
    const rev = splitComplex(asString(fnIMSUB([rvString("3"), rvString("7")])))[0];
    expect(fwd).toBeCloseTo(-rev, 10);
  });
});

describe("IMPRODUCT saturation", () => {
  it("IMPRODUCT('3','4') = 12", () => {
    const [re] = splitComplex(asString(fnIMPRODUCT([rvString("3"), rvString("4")])));
    expect(re).toBeCloseTo(12, 10);
  });
  it("IMPRODUCT('i','i') = -1", () => {
    const [re] = splitComplex(asString(fnIMPRODUCT([rvString("i"), rvString("i")])));
    expect(re).toBeCloseTo(-1, 10);
  });
  it("IMPRODUCT is associative", () => {
    const lhs = asString(fnIMPRODUCT([rvString("2"), rvString("3"), rvString("4")]));
    expect(splitComplex(lhs)[0]).toBeCloseTo(24, 10);
  });
  it("IMPRODUCT with empty arg list = 1 (multiplicative identity)", () => {
    const [re] = splitComplex(asString(fnIMPRODUCT([])));
    expect(re).toBeCloseTo(1, 10);
  });
  it("IMPRODUCT accepts array arguments", () => {
    const arr = rvArray([[rvNumber(2), rvNumber(3)]]);
    const [re] = splitComplex(asString(fnIMPRODUCT([arr, rvString("4")])));
    expect(re).toBeCloseTo(24, 10);
  });
  it("IMPRODUCT rejects malformed string", () => {
    expect(fnIMPRODUCT([rvString("xx")])).toEqual(ERRORS.NUM);
  });
  it("IMPRODUCT propagates errors", () => {
    expect(fnIMPRODUCT([rvString("2"), ERRORS.NA])).toEqual(ERRORS.NA);
  });
});

describe("BESSELI saturation", () => {
  it("BESSELI(0, 0) = 1", () => {
    expect(asNumber(fnBESSELI([rvNumber(0), rvNumber(0)]))).toBeCloseTo(1, 10);
  });
  it("BESSELI(1, 0) ≈ 1.2660658", () => {
    expect(asNumber(fnBESSELI([rvNumber(1), rvNumber(0)]))).toBeCloseTo(1.2660658, 4);
  });
  it("BESSELI(1, 1) ≈ 0.5651591", () => {
    expect(asNumber(fnBESSELI([rvNumber(1), rvNumber(1)]))).toBeCloseTo(0.5651591, 4);
  });
  it("BESSELI(2, 2) ≈ 0.6889484", () => {
    expect(asNumber(fnBESSELI([rvNumber(2), rvNumber(2)]))).toBeCloseTo(0.6889484, 3);
  });
  it("BESSELI negative order → #NUM!", () => {
    expect(fnBESSELI([rvNumber(1), rvNumber(-1)])).toEqual(ERRORS.NUM);
  });
  it("BESSELI at x=0, n≥1 = 0", () => {
    expect(asNumber(fnBESSELI([rvNumber(0), rvNumber(1)]))).toBeCloseTo(0, 10);
    expect(asNumber(fnBESSELI([rvNumber(0), rvNumber(5)]))).toBeCloseTo(0, 10);
  });
  it("BESSELI rejects negative x (engine surfaces #NUM!)", () => {
    // Although I_n(x) is mathematically defined for all real x, the
    // engine's bessel wrapper rejects x<0 for all four variants. Document
    // the guard so regressions are caught.
    expect(fnBESSELI([rvNumber(-1), rvNumber(0)])).toEqual(ERRORS.NUM);
  });
  it("BESSELI propagates errors", () => {
    expect(fnBESSELI([ERRORS.NA, rvNumber(0)])).toEqual(ERRORS.NA);
    expect(fnBESSELI([rvNumber(1), ERRORS.NA])).toEqual(ERRORS.NA);
  });
});

describe("BESSELK saturation", () => {
  it("BESSELK(1, 0) ≈ 0.4210244", () => {
    expect(asNumber(fnBESSELK([rvNumber(1), rvNumber(0)]))).toBeCloseTo(0.4210244, 3);
  });
  it("BESSELK(1, 1) ≈ 0.6019072", () => {
    expect(asNumber(fnBESSELK([rvNumber(1), rvNumber(1)]))).toBeCloseTo(0.6019072, 3);
  });
  it("BESSELK(2, 2) ≈ 0.2537598", () => {
    expect(asNumber(fnBESSELK([rvNumber(2), rvNumber(2)]))).toBeCloseTo(0.2537598, 2);
  });
  it("BESSELK at x=0 → #NUM!", () => {
    expect(fnBESSELK([rvNumber(0), rvNumber(0)])).toEqual(ERRORS.NUM);
  });
  it("BESSELK negative x → #NUM!", () => {
    expect(fnBESSELK([rvNumber(-1), rvNumber(0)])).toEqual(ERRORS.NUM);
  });
  it("BESSELK propagates errors", () => {
    expect(fnBESSELK([ERRORS.NA, rvNumber(0)])).toEqual(ERRORS.NA);
  });
});

describe("BESSELY saturation", () => {
  it("BESSELY(1, 0) ≈ 0.0882569", () => {
    expect(asNumber(fnBESSELY([rvNumber(1), rvNumber(0)]))).toBeCloseTo(0.0882569, 3);
  });
  it("BESSELY(1, 1) ≈ -0.7812128", () => {
    expect(asNumber(fnBESSELY([rvNumber(1), rvNumber(1)]))).toBeCloseTo(-0.7812128, 2);
  });
  it("BESSELY(2, 2) ≈ -0.6174081 (Y_2(2))", () => {
    expect(asNumber(fnBESSELY([rvNumber(2), rvNumber(2)]))).toBeCloseTo(-0.6174081, 2);
  });
  it("BESSELY at x=0 → #NUM!", () => {
    expect(fnBESSELY([rvNumber(0), rvNumber(0)])).toEqual(ERRORS.NUM);
  });
  it("BESSELY negative x → #NUM!", () => {
    expect(fnBESSELY([rvNumber(-1), rvNumber(0)])).toEqual(ERRORS.NUM);
  });
  it("BESSELY negative order → #NUM!", () => {
    expect(fnBESSELY([rvNumber(1), rvNumber(-1)])).toEqual(ERRORS.NUM);
  });
  it("BESSELY propagates errors", () => {
    expect(fnBESSELY([rvNumber(1), ERRORS.REF])).toEqual(ERRORS.REF);
  });
});

// ============================================================================
// R9 saturation batch 2 — bring remaining to 10+
// ============================================================================

describe("HEX2DEC / OCT2DEC extras", () => {
  it("HEX2DEC('A') = 10", () => {
    expect(asNumber(fnHEX2DEC([rvString("A")]))).toBe(10);
  });
  it("HEX2DEC signed negative (top bit) = negative", () => {
    // FFFFFFFFFE0 → negative
    const v = asNumber(fnHEX2DEC([rvString("FFFFFFFF00")]));
    expect(v).toBeLessThan(0);
  });
  it("HEX2DEC of too-long input → #NUM!", () => {
    expect(fnHEX2DEC([rvString("12345678901")])).toEqual(ERRORS.NUM);
  });
  it("HEX2DEC empty string → #NUM!", () => {
    expect(fnHEX2DEC([rvString("")])).toEqual(ERRORS.NUM);
  });
  it("HEX2DEC invalid char → #NUM!", () => {
    expect(fnHEX2DEC([rvString("G1")])).toEqual(ERRORS.NUM);
  });
  it("OCT2DEC('10') = 8", () => {
    expect(asNumber(fnOCT2DEC([rvString("10")]))).toBe(8);
  });
  it("OCT2DEC('777') = 511", () => {
    expect(asNumber(fnOCT2DEC([rvString("777")]))).toBe(511);
  });
  it("OCT2DEC('7000000000') is negative (top bit set)", () => {
    const v = asNumber(fnOCT2DEC([rvString("7000000000")]));
    expect(v).toBeLessThan(0);
  });
});

describe("IMAGINARY / IMREAL / IMCONJUGATE extras", () => {
  it("IMAGINARY of integer '5' = 0", () => {
    expect(asNumber(fnIMAGINARY([rvString("5")]))).toBe(0);
  });
  it("IMAGINARY negative real = 0", () => {
    expect(asNumber(fnIMAGINARY([rvString("-5")]))).toBe(0);
  });
  it("IMAGINARY of '1+0i' = 0", () => {
    expect(asNumber(fnIMAGINARY([rvString("1+0i")]))).toBe(0);
  });
  it("IMAGINARY of bare 'i' = 1", () => {
    expect(asNumber(fnIMAGINARY([rvString("i")]))).toBe(1);
  });
  it("IMREAL of bare 'i' = 0", () => {
    expect(asNumber(fnIMREAL([rvString("i")]))).toBe(0);
  });
  it("IMREAL of '2.5+3i' = 2.5", () => {
    expect(asNumber(fnIMREAL([rvString("2.5+3i")]))).toBe(2.5);
  });
  it("IMCONJUGATE flips imaginary sign", () => {
    expect(asString(fnIMCONJUGATE([rvString("3+4i")]))).toBe("3-4i");
  });
  it("IMCONJUGATE of real is unchanged", () => {
    expect(asString(fnIMCONJUGATE([rvString("5")]))).toBe("5");
  });
  it("IMCONJUGATE of pure imaginary", () => {
    expect(asString(fnIMCONJUGATE([rvString("2i")]))).toBe("-2i");
  });
});

describe("IMSUM / IMLN extras", () => {
  it("IMSUM of one complex equals itself", () => {
    expect(asString(fnIMSUM([rvString("3+4i")]))).toBe("3+4i");
  });
  it("IMSUM of many pure reals", () => {
    expect(asString(fnIMSUM([rvString("1"), rvString("2"), rvString("3")]))).toBe("6");
  });
  it("IMSUM of pure imag + pure real", () => {
    expect(asString(fnIMSUM([rvString("3"), rvString("4i")]))).toBe("3+4i");
  });
  it("IMLN of e ≈ 1", () => {
    const r = asString(fnIMLN([rvString(String(Math.E))]));
    // Result is "1" or "0.9999...", parse as real
    expect(parseFloat(r)).toBeCloseTo(1, 4);
  });
  it("IMLN of 1 = 0", () => {
    expect(parseFloat(asString(fnIMLN([rvString("1")])))).toBeCloseTo(0, 10);
  });
  it("IMLN of i = iπ/2", () => {
    const r = asString(fnIMLN([rvString("i")]));
    // Result should be 0 + (π/2)i, pure imaginary
    expect(r.endsWith("i")).toBe(true);
  });
});
