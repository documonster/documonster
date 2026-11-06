/**
 * `Query.getTableStyle` returns a named, read-only `ResolvedTableStyle`, as
 * its siblings return `ResolvedParagraphStyle` / `ResolvedRunStyle`; and a
 * resolved numbering level's `format` / `justification` carry the
 * vocabularies they come from.
 */

import { describe, it, expect, expectTypeOf } from "vitest";

import { Document, Query } from "../index";
import type {
  LevelJustification,
  NumberFormat,
  ResolvedNumberingLevel,
  ResolvedTableStyle
} from "../index";

describe("resolved style types", () => {
  it("getTableStyle returns ResolvedTableStyle", () => {
    const d = Document.create();
    Document.useDefaultStyles(d);
    const resolved = Query.getTableStyle(d, "TableGrid");
    expectTypeOf(resolved).toEqualTypeOf<ResolvedTableStyle>();
    expectTypeOf(resolved.chain).toEqualTypeOf<readonly string[]>();
    expect(Array.isArray(resolved.chain)).toBe(true);
  });

  it("ResolvedNumberingLevel names its vocabularies", () => {
    expectTypeOf<ResolvedNumberingLevel["format"]>().toEqualTypeOf<NumberFormat | undefined>();
    expectTypeOf<ResolvedNumberingLevel["justification"]>().toEqualTypeOf<
      LevelJustification | undefined
    >();
  });
});
