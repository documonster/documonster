import { styles } from "@excel/__tests__/shared/styles";
import { copyStyle } from "@excel/utils/copy-style";
import { describe, it, expect } from "vitest";

const style1 = {
  numFmt: styles.numFmts.numFmt1,
  font: styles.fonts.broadwayRedOutline20,
  alignment: styles.namedAlignments.topLeft,
  border: styles.borders.thickRainbow,
  fill: styles.fills.redGreenDarkTrellis
};
const style2 = {
  fill: styles.fills.rgbPathGrad
};
const style3 = {
  protection: { locked: true, hidden: false }
};

describe("copyStyle", () => {
  it("should copy a style deeply", () => {
    const copied = copyStyle(style1);
    expect(copied).toEqual(style1);
    expect(copied!.font).not.toBe(style1.font);
    expect(copied!.font.color).toEqual(style1.font.color);
    expect(copied!.font.color).not.toBe(style1.font.color);
    expect(copied!.alignment).not.toBe(style1.alignment);
    expect(copied!.border).not.toBe(style1.border);
    expect(copied!.border.top).not.toBe(style1.border.top);
    expect(copied!.border.left).not.toBe(style1.border.left);
    expect(copied!.border.bottom).not.toBe(style1.border.bottom);
    expect(copied!.border.right).not.toBe(style1.border.right);
    expect(copied!.border.diagonal).not.toBe(style1.border.diagonal);
    expect(copied!.border.top.color).toEqual(style1.border.top.color);
    expect(copied!.border.top.color).not.toBe(style1.border.top.color);
    expect(copied!.fill).not.toBe(style1.fill);
    expect(copied!.fill.fgColor).toEqual(style1.fill.fgColor);
    expect(copied!.fill.fgColor).not.toBe(style1.fill.fgColor);
    expect(copied!.fill.bgColor).toEqual(style1.fill.bgColor);
    expect(copied!.fill.bgColor).not.toBe(style1.fill.bgColor);

    expect(copyStyle({})).toEqual({});
  });

  it("should copy fill.stops deeply", () => {
    const copied = copyStyle(style2);
    expect(copied!.fill.stops).toEqual(style2.fill.stops);
    expect(copied!.fill.stops).not.toBe(style2.fill.stops);
    expect(copied!.fill.stops[0]).not.toBe(style2.fill.stops[0]);
    expect(copied!.fill.stops[0].color).toEqual(style2.fill.stops[0].color);
    expect(copied!.fill.stops[0].color).not.toBe(style2.fill.stops[0].color);
    expect(copied!.fill.center).toEqual(style2.fill.center);
    expect(copied!.fill.center).not.toBe(style2.fill.center);
  });

  it("should copy protection one level deep", () => {
    const copied = copyStyle(style3);
    expect(copied).toEqual(style3);
    expect(copied!.protection).not.toBe(style3.protection);
  });

  it("should return the argument if a falsy value passed", () => {
    expect(copyStyle(null)).toBe(null);
    expect(copyStyle(undefined)).toBe(undefined);
  });
});
