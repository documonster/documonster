/**
 * Number format codes: the built-in table, and the escapes a code read out of a file keeps.
 *
 * The table holds a code only for an id every locale shows the same way. The accounting formats 41–44 are the
 * locale's, as the currencies 5–8 are: Excel writes each with a `<numFmt>` declaring its own locale's form (a Chinese
 * Excel `￥`, an English one `$`), so a code read out of a file is the code, and a bare id is kept as an id
 * (`kept-num-fmt.ts`). They once carried the Chinese codes here, so a bare id 42 read as `￥` whatever the file, and a
 * `￥` code was written as id 42 — which an English Excel shows in dollars.
 */
import { defaultNumFormats } from "@excel/xlsx/defaultnumformats";
import { NumFmtXform } from "@excel/xlsx/xform/style/numfmt-xform";
import { describe, expect, it } from "vitest";

/** The four accounting formats as a Chinese Excel writes them — read out of a template it produced. */
const ACCOUNTING = {
  41: '_ * #,##0_ ;_ * \\-#,##0_ ;_ * "-"_ ;_ @_ ',
  42: '_ "￥"* #,##0_ ;_ "￥"* \\-#,##0_ ;_ "￥"* "-"_ ;_ @_ ',
  43: '_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * "-"??_ ;_ @_ ',
  44: '_ "￥"* #,##0.00_ ;_ "￥"* \\-#,##0.00_ ;_ "￥"* "-"??_ ;_ @_ '
} as const;

describe("the built-in number format table", () => {
  it.each(Object.keys(ACCOUNTING))("has no one code for the accounting format %s", id => {
    expect(NumFmtXform.getDefaultFmtCode(Number(id))).toBeUndefined();
  });

  it.each(Object.entries(ACCOUNTING))(
    "does not write the Chinese code for %s as the id",
    (id, code) => {
      expect(defaultNumFormats[Number(id)]).toBeUndefined();
      // Written as the code, it shows as the file's author saw it in every locale; written as the id, it would not.
      expect(NumFmtXform.getDefaultFmtId(code)).toBeUndefined();
    }
  );

  it("has id 47 as Excel shows and writes it, not as ECMA-376's table misprints it", () => {
    expect(NumFmtXform.getDefaultFmtCode(47)).toBe("mm:ss.0");
    expect(NumFmtXform.getDefaultFmtId("mm:ss.0")).toBe(47);
    expect(NumFmtXform.getDefaultFmtId("mmss.0")).toBeUndefined();
  });
});

describe("a format code read out of a file", () => {
  it("keeps its backslash escapes", () => {
    const xform = new NumFmtXform();
    xform.parseOpen({
      name: "numFmt",
      attributes: { numFmtId: "41", formatCode: ACCOUNTING[41] }
    } as never);
    // Verbatim. A backslash escapes the character after it and is part of the code's meaning — unescaping
    // changes what the format *does*, not merely how it is spelled.
    expect((xform.model as { formatCode: string }).formatCode).toBe(ACCOUNTING[41]);
  });
});
