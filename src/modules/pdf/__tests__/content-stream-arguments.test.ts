/**
 * A page's content stream (`Pdf.PageBuilder#getContentStream`) is public, so its operators refuse an
 * operand of the wrong shape as `PdfArgumentError` rather than failing inside or writing a broken one.
 */
import { Pdf, PdfArgumentError } from "@pdf/index";
import { describe, expect, it } from "vitest";

/* oxlint-disable no-sparse-arrays -- holes are what these cases pass */

function stream() {
  return new Pdf.Builder().addPage().getContentStream();
}

describe("PdfContentStream operands", () => {
  it.each([
    ["setFillColor", "color", () => stream().setFillColor(null as never)],
    ["setStrokeColor", "color", () => stream().setStrokeColor({ r: "1" } as never)],
    ["append", "other", () => stream().append({} as never)],
    ["setDashPattern", "dashArray", () => stream().setDashPattern("3 2" as never)],
    ["showText", "text", () => stream().showText(42 as never)],
    ["showTextHex", "hexString", () => stream().showTextHex(undefined as never)],
    [
      "showTextHexWithAdjustments",
      "parts[0]",
      () => stream().showTextHexWithAdjustments([{}] as never)
    ],
    ["nextLineShowText", "text", () => stream().nextLineShowText(null as never)],
    ["raw", "operator", () => stream().raw(7 as never)]
  ])("#%s names %s", (method, field, call) => {
    expect(call).toThrow(PdfArgumentError);
    expect(call).toThrow(expect.objectContaining({ field }));
    expect(call).toThrow(`PdfContentStream#${method}`);
  });

  it.each([
    ["fillRect", (c: ReturnType<typeof stream>) => c.fillRect(0, 0, 1, 1, null as never)],
    ["drawLine", (c: ReturnType<typeof stream>) => c.drawLine(0, 0, 1, 1, null as never, 1)],
    [
      "drawLine dash",
      (c: ReturnType<typeof stream>) =>
        c.drawLine(0, 0, 1, 1, { r: 0, g: 0, b: 0 }, 1, "2" as never)
    ],
    ["drawImage", (c: ReturnType<typeof stream>) => c.drawImage("a b", 0, 0, 1, 1)]
  ])("%s refuses before it writes anything", (_name, call) => {
    const content = stream();
    expect(() => call(content)).toThrow(PdfArgumentError);
    expect(content.toString()).toBe("");
  });

  it.each([
    [
      "setDashPattern",
      "dashArray[1]",
      (c: ReturnType<typeof stream>) => c.setDashPattern([3, , 2] as never)
    ],
    [
      "showTextHexWithAdjustments",
      "parts[1]",
      (c: ReturnType<typeof stream>) => c.showTextHexWithAdjustments(["<01>", , 10] as never)
    ],
    ["doXObject", "name", (c: ReturnType<typeof stream>) => c.doXObject(1n as never)]
  ])("#%s refuses %s, describing it, and writes nothing", (method, field, call) => {
    const content = stream();
    expect(() => call(content)).toThrow(
      expect.objectContaining({ name: "PdfArgumentError", field })
    );
    expect(() => call(content)).toThrow(`PdfContentStream#${method}: ${field} must be`);
    expect(content.toString()).toBe("");
  });

  it("deferred refuses a callback that is not a function when it is given", () => {
    const content = stream();
    expect(() => content.deferred(42 as never)).toThrow(
      expect.objectContaining({ field: "produce" })
    );
    expect(() => content.deferred(() => "", 42 as never)).toThrow(
      expect.objectContaining({ field: "preflight" })
    );
    expect(content.hasContent()).toBe(false);
  });

  it("still accepts what the builder itself writes", () => {
    const content = stream()
      .setFillColor({ r: 1, g: 0, b: 0 })
      .setDashPattern([3, 2])
      .showText("Hi")
      .append(stream().raw("q"));
    expect(content).toBeDefined();
  });
});
