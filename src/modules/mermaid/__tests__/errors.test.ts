import {
  Mermaid,
  MermaidError,
  MermaidParseError,
  MermaidSyntaxError,
  isMermaidError
} from "@mermaid/index";
import { BaseError } from "@utils/errors";
import { describe, expect, it } from "vitest";

function thrownBy(run: () => unknown): unknown {
  try {
    run();
  } catch (err) {
    return err;
  }
  throw new Error("expected the call to throw");
}

describe("Mermaid errors", () => {
  it("MermaidParseError extends MermaidError, BaseError and Error", () => {
    const error = new MermaidParseError("bad diagram");
    expect(error).toBeInstanceOf(MermaidParseError);
    expect(error).toBeInstanceOf(MermaidError);
    expect(error).toBeInstanceOf(BaseError);
    expect(error).toBeInstanceOf(Error);
    expect(error.name).toBe("MermaidParseError");
    expect(error.message).toBe("bad diagram");
    expect(error.line).toBeUndefined();
    expect(error.column).toBeUndefined();
  });

  it("the deprecated MermaidSyntaxError is the same class, so instanceof answers alike", () => {
    expect(MermaidSyntaxError).toBe(MermaidParseError);
    const thrown = thrownBy(() => Mermaid.parse("zenuml\n title A"));
    expect(thrown).toBeInstanceOf(MermaidSyntaxError);
    expect(new MermaidSyntaxError("x")).toBeInstanceOf(MermaidParseError);
  });

  it("locates the statement that names an unsupported diagram", () => {
    const error = thrownBy(() =>
      Mermaid.parse("%% a comment\n\n   zenuml\n title A")
    ) as MermaidParseError;
    expect(error).toBeInstanceOf(MermaidParseError);
    expect(error).toMatchObject({ line: 3, column: 4 });
    const behindFrontmatter = thrownBy(() =>
      Mermaid.parse("---\ntitle: x\n---\nnope")
    ) as MermaidParseError;
    expect(behindFrontmatter).toMatchObject({ line: 4, column: 1 });
  });

  it("has no location for an empty source, which has no statement", () => {
    const error = thrownBy(() => Mermaid.parse("  \n%% only a comment\n")) as MermaidParseError;
    expect(error).toBeInstanceOf(MermaidParseError);
    expect(error.line).toBeUndefined();
    expect(error.column).toBeUndefined();
  });

  it("takes line and column beside a cause", () => {
    const cause = new Error("inner");
    const error = new MermaidParseError("bad", { line: 2, column: 5, cause });
    expect(error).toMatchObject({ line: 2, column: 5, cause });
  });

  it("MermaidError is named for itself", () => {
    const error = new MermaidError("x");
    expect(error.name).toBe("MermaidError");
    expect(error).toBeInstanceOf(BaseError);
  });

  it("isMermaidError is true for the base class and every subclass", () => {
    expect(isMermaidError(new MermaidError("x"))).toBe(true);
    expect(isMermaidError(new MermaidParseError("x"))).toBe(true);
  });

  it("isMermaidError is true for what the public entry points throw", () => {
    for (const run of [
      () => Mermaid.parse("nonsenseDiagram\n  A --> B"),
      () => Mermaid.toSvg(""),
      () => Mermaid.toDrawList("zenuml\n title A")
    ]) {
      const error = thrownBy(run);
      expect(error).toBeInstanceOf(MermaidParseError);
      expect(isMermaidError(error)).toBe(true);
    }
  });

  it("isMermaidError is false for unrelated errors and non-errors", () => {
    for (const value of [
      new BaseError("x"),
      new Error("x"),
      new TypeError("x"),
      { name: "MermaidParseError", message: "x" },
      "MermaidError",
      null,
      undefined
    ]) {
      expect(isMermaidError(value)).toBe(false);
    }
  });
});

describe("arguments of the wrong kind are refused by name", () => {
  it("parse, toSvg, toDrawList and layoutGraph throw MermaidError naming the function", () => {
    expect(() => Mermaid.parse(42 as never)).toThrow(MermaidError);
    expect(() => Mermaid.parse(null as never)).toThrow(
      /^Mermaid\.parse: source must be Mermaid text; got null/
    );
    expect(() => Mermaid.toSvg(42 as never)).toThrow(MermaidError);
    expect(() => Mermaid.toDrawList({} as never)).toThrow(/not a diagram/);
    const pie = Mermaid.parse('pie\n "A" : 1');
    expect(() => Mermaid.layoutGraph(pie as never)).toThrow(
      /^Mermaid\.layoutGraph: diagram must be a graph diagram.*"pie" diagram/
    );
    expect(Mermaid.layoutGraph(Mermaid.parse("flowchart LR\n a --> b") as never).nodes.length).toBe(
      2
    );
  });
});
