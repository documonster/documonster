/**
 * `Build.mathFraction`, `mathLimit`, `mathAccent` and `mathBar` take their one setting in an options object; the
 * positional form is a deprecated forward that must build the same node, and both are told apart the same way.
 */
import { Build, DocxArgumentError } from "@word/index";
import { describe, expect, it } from "vitest";

const m = [Build.mathRun("x")];

describe("math builders take their setting in options", () => {
  it("mathFraction", () => {
    expect(Build.mathFraction(m, m, { fractionType: "lin" })).toEqual(
      Build.mathFraction(m, m, "lin")
    );
    expect(Build.mathFraction(m, m, {})).toEqual(Build.mathFraction(m, m));
  });

  it("mathLimit, upper or the default lower", () => {
    expect(Build.mathLimit(m, m, { limitType: "upper" })).toEqual(Build.mathLimit(m, m, "upper"));
    expect(Build.mathLimit(m, m, {})).toMatchObject({ limitType: "lower" });
    expect(Build.mathLimit(m, m)).toMatchObject({ limitType: "lower" });
  });

  it("mathLimit defaults only an absent type: null is refused, as the positional form always did", () => {
    expect(() => Build.mathLimit(m, m, null as never)).toThrow(DocxArgumentError);
    expect(() => Build.mathLimit(m, m, { limitType: null as never })).toThrow(
      expect.objectContaining({ field: "options.limitType" })
    );
  });

  it("mathAccent", () => {
    expect(Build.mathAccent(m, { char: "→" })).toEqual(Build.mathAccent(m, "→"));
    expect(Build.mathAccent(m, {})).toEqual(Build.mathAccent(m));
  });

  it("names the option a wrong value was given for", () => {
    const fields = [
      () => Build.mathFraction(m, m, { fractionType: "over" as never }),
      () => Build.mathLimit(m, m, { limitType: "side" as never }),
      () => Build.mathAccent(m, { char: "^^" }),
      () => Build.mathBar(m, { position: "side" as never })
    ].map(call => {
      try {
        call();
      } catch (error) {
        expect(error).toBeInstanceOf(DocxArgumentError);
        return (error as DocxArgumentError).field;
      }
      return undefined;
    });
    expect(fields).toEqual([
      "options.fractionType",
      "options.limitType",
      "options.char",
      "options.position"
    ]);
  });

  it("reads anything but an options object as the deprecated positional setting", () => {
    for (const call of [
      () => Build.mathFraction(m, m, null as never),
      () => Build.mathLimit(m, m, 5 as never),
      () => Build.mathBar(m, null as never)
    ]) {
      expect(call).toThrow(DocxArgumentError);
    }
    expect(() => Build.mathBar(m, null as never)).toThrow(
      expect.objectContaining({ field: "position" })
    );
  });
});
