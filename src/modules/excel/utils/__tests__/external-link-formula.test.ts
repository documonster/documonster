/**
 * Unit tests for the external-link formula utilities. These exercise the
 * regex + rewriter in isolation so the workbook-level writer pass can rely
 * on them as a black box.
 */

import { findExternalRefs, rewriteExternalRefs } from "@excel/utils/external-link-formula";
import { describe, expect, it } from "vitest";

describe("external-link-formula: findExternalRefs", () => {
  it("finds a single unquoted filename-form reference", () => {
    const refs = findExternalRefs("[测试.xlsx]Sheet1!A1");
    expect(refs).toHaveLength(1);
    expect(refs[0]).toMatchObject({
      match: "[测试.xlsx]Sheet1!",
      workbook: "测试.xlsx",
      numeric: false,
      index: null,
      sheet: "Sheet1",
      quoted: false,
      start: 0
    });
  });

  it("finds an unquoted numeric-form reference", () => {
    const refs = findExternalRefs("[3]Sheet1!A1");
    expect(refs).toHaveLength(1);
    expect(refs[0]).toMatchObject({
      match: "[3]Sheet1!",
      workbook: "3",
      numeric: true,
      index: 3,
      sheet: "Sheet1"
    });
  });

  it("finds a quoted reference with a spaced sheet name", () => {
    const refs = findExternalRefs("'[Big Book.xlsx]Sheet One'!B12");
    expect(refs).toHaveLength(1);
    expect(refs[0]).toMatchObject({
      workbook: "Big Book.xlsx",
      sheet: "Sheet One",
      quoted: true
    });
  });

  it("decodes doubled single quotes inside a quoted sheet name", () => {
    // "Sheet''s" represents the logical sheet name "Sheet's"
    const refs = findExternalRefs("'[Book.xlsx]Sheet''s Data'!A1");
    expect(refs).toHaveLength(1);
    expect(refs[0].sheet).toBe("Sheet's Data");
  });

  it("finds multiple references in the same formula", () => {
    const refs = findExternalRefs("[a.xlsx]S1!A1+[b.xlsx]S2!B2*[1]S3!C3");
    expect(refs.map(r => r.workbook)).toEqual(["a.xlsx", "b.xlsx", "1"]);
  });

  it("ignores references that appear inside string literals", () => {
    const refs = findExternalRefs('CONCAT("[Book]Sheet!A1", [real.xlsx]S!A1)');
    expect(refs).toHaveLength(1);
    expect(refs[0].workbook).toBe("real.xlsx");
  });

  it("ignores structured references like [@Col] and [#Headers]", () => {
    expect(findExternalRefs("SUM([@Value])")).toHaveLength(0);
    expect(findExternalRefs("SUM([#Headers])")).toHaveLength(0);
    expect(findExternalRefs("Table1[Column Name]")).toHaveLength(0);
  });

  it("returns nothing for formulas with no external refs", () => {
    expect(findExternalRefs("A1+B1")).toEqual([]);
    expect(findExternalRefs("SUM(Sheet1!A1:A5)")).toEqual([]);
    expect(findExternalRefs("")).toEqual([]);
  });

  it("accepts subdirectory paths inside the quoted form", () => {
    // Excel writes path-bearing external refs in the quoted form.
    const refs = findExternalRefs("'[data/src.xlsx]Sheet1'!A1");
    expect(refs).toHaveLength(1);
    expect(refs[0].workbook).toBe("data/src.xlsx");
    expect(refs[0].sheet).toBe("Sheet1");
  });

  it("accepts absolute paths inside the quoted form", () => {
    const refs = findExternalRefs("'[file:///C:/tmp/src.xlsx]Sheet1'!A1");
    expect(refs).toHaveLength(1);
    expect(refs[0].workbook).toBe("file:///C:/tmp/src.xlsx");
  });

  it("does not match `[Column]` structured refs with space-separated suffix", () => {
    // `[Column] !A1` would be malformed but regex must not over-match.
    expect(findExternalRefs("SUM([@Col] + 1)")).toEqual([]);
  });
});

describe("external-link-formula: rewriteExternalRefs", () => {
  it("rewrites a filename reference to the numeric form", () => {
    const out = rewriteExternalRefs("[测试.xlsx]Sheet1!A1", ref =>
      ref.workbook === "测试.xlsx" ? 1 : null
    );
    expect(out).toBe("[1]Sheet1!A1");
  });

  it("preserves the sheet name and tail of the reference", () => {
    const out = rewriteExternalRefs("[a.xlsx]Sheet1!A1:B10+C1", () => 2);
    expect(out).toBe("[2]Sheet1!A1:B10+C1");
  });

  it("rewrites quoted references keeping single quotes intact", () => {
    const out = rewriteExternalRefs("'[a.xlsx]Big Sheet'!$A$1", () => 7);
    expect(out).toBe("'[7]Big Sheet'!$A$1");
  });

  it("rewrites each reference independently within the same formula", () => {
    const out = rewriteExternalRefs("[a.xlsx]S1!A1+[b.xlsx]S2!B2", ref =>
      ref.workbook === "a.xlsx" ? 1 : 2
    );
    expect(out).toBe("[1]S1!A1+[2]S2!B2");
  });

  it("leaves references unchanged when the resolver returns null", () => {
    const out = rewriteExternalRefs("[unknown.xlsx]S!A1+1", () => null);
    expect(out).toBe("[unknown.xlsx]S!A1+1");
  });

  it("leaves numeric references unchanged when resolver returns same index", () => {
    const out = rewriteExternalRefs("[2]S!A1", ref => ref.index);
    expect(out).toBe("[2]S!A1");
  });

  it("does not touch ref tokens inside string literals", () => {
    const out = rewriteExternalRefs('"[a.xlsx]S!A1"+[b.xlsx]S!A1', () => 9);
    expect(out).toBe('"[a.xlsx]S!A1"+[9]S!A1');
  });
});
