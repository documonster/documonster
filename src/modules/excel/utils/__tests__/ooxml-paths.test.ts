import { resolveRelTarget } from "@excel/utils/ooxml-paths";
import { describe, it, expect } from "vitest";

describe("resolveRelTarget", () => {
  it("resolves a relative target with ../", () => {
    expect(resolveRelTarget("xl/worksheets/", "../comments1.xml")).toBe("xl/comments1.xml");
  });

  it("resolves a deeply nested relative target", () => {
    expect(resolveRelTarget("xl/worksheets/", "../../comments1.xml")).toBe("comments1.xml");
  });

  it("strips leading slash for absolute targets", () => {
    expect(resolveRelTarget("xl/worksheets/", "/xl/comments/comment1.xml")).toBe(
      "xl/comments/comment1.xml"
    );
  });

  it("resolves sibling relative target", () => {
    expect(resolveRelTarget("xl/worksheets/", "../tables/table1.xml")).toBe("xl/tables/table1.xml");
  });

  it("resolves target in same directory", () => {
    expect(resolveRelTarget("xl/worksheets/", "sheet2.xml")).toBe("xl/worksheets/sheet2.xml");
  });

  it("resolves target with . segments", () => {
    expect(resolveRelTarget("xl/worksheets/", "./../drawings/drawing1.xml")).toBe(
      "xl/drawings/drawing1.xml"
    );
  });

  it("handles baseDir without trailing slash", () => {
    expect(resolveRelTarget("xl/worksheets", "../comments1.xml")).toBe("xl/comments1.xml");
  });

  it("handles baseDir without trailing slash for same-dir target", () => {
    expect(resolveRelTarget("xl/worksheets", "sheet2.xml")).toBe("xl/worksheets/sheet2.xml");
  });
});
