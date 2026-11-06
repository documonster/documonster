/**
 * Mermaid argument checks throw `MermaidArgumentError` (a `MermaidError`) naming the field
 * (X12); an unreadable theme colour is refused (X9); a wrong handle is checked (X10).
 */
import {
  Mermaid,
  MermaidArgumentError,
  MermaidError,
  MermaidParseError,
  centredText,
  layoutFlowchart
} from "@mermaid/index";
import { describe, expect, it } from "vitest";

function argumentError(run: () => unknown): MermaidArgumentError {
  try {
    run();
  } catch (err) {
    expect(err).toBeInstanceOf(MermaidArgumentError);
    expect(err).toBeInstanceOf(MermaidError);
    return err as MermaidArgumentError;
  }
  throw new Error("expected a MermaidArgumentError");
}

describe("MermaidArgumentError", () => {
  it("is a MermaidError with a readonly field", () => {
    const error = new MermaidArgumentError("source", "x");
    expect(error.name).toBe("MermaidArgumentError");
    expect(error.field).toBe("source");
    expect(error).toBeInstanceOf(MermaidError);
  });

  it("names the field for argument and option checks", () => {
    expect(argumentError(() => Mermaid.parse(42 as never)).field).toBe("source");
    expect(argumentError(() => Mermaid.toSvg("graph TD\nA-->B", 42 as never)).field).toBe(
      "options"
    );
    expect(
      argumentError(() => Mermaid.toSvg("graph TD\nA-->B", { theme: { nope: "#fff" } as never }))
        .field
    ).toBe("options.theme.nope");
    expect(argumentError(() => Mermaid.toSvg("graph TD", { fontSize: Number.NaN })).field).toBe(
      "options.fontSize"
    );
    expect(argumentError(() => Mermaid.toDrawList("graph TD", { background: "bogus" })).field).toBe(
      "options.background"
    );
    // The first parameter of toSvg and toDrawList is `source`, a string or a parsed diagram.
    expect(argumentError(() => Mermaid.toSvg(42 as never)).field).toBe("source");
    expect(argumentError(() => Mermaid.toDrawList({} as never)).field).toBe("source");
    expect(argumentError(() => Mermaid.resolveTheme({ nodeFill: 42 as never })).field).toBe(
      "options.nodeFill"
    );
    expect(
      argumentError(() => Mermaid.createCenteredText(null as never, ["a"], { size: 10 })).field
    ).toBe("box");
    expect(argumentError(() => Mermaid.layoutGraph(42 as never)).field).toBe("diagram");
  });

  it("starts each message with the public call (X6)", () => {
    for (const [run, prefix] of [
      [() => Mermaid.parse(42 as never), "Mermaid.parse: "],
      [() => Mermaid.toSvg("graph TD", { padding: -1 }), "Mermaid.toSvg: "],
      [() => Mermaid.toDrawList("graph TD", { padding: -1 }), "Mermaid.toDrawList: "],
      [() => Mermaid.layoutGraph(42 as never), "Mermaid.layoutGraph: "],
      [() => Mermaid.resolveTheme({ nodeFill: "bogus" }), "Mermaid.resolveTheme: "]
    ] as const) {
      expect(argumentError(run).message.startsWith(prefix)).toBe(true);
    }
  });

  it("MermaidParseError takes a cause like every other error class (R4P-S7)", () => {
    const cause = new Error("inner");
    const error = new MermaidParseError("bad", { cause });
    expect(error.name).toBe("MermaidParseError");
    expect(error.cause).toBe(cause);
    expect(error).toBeInstanceOf(MermaidError);
  });

  it("is not used for syntax errors", () => {
    let caught: unknown;
    try {
      Mermaid.parse("not a diagram at all");
    } catch (err) {
      caught = err;
    }
    expect(caught).toBeInstanceOf(MermaidParseError);
    expect(caught).not.toBeInstanceOf(MermaidArgumentError);
  });
});

describe("theme colours are read, not guessed (X9)", () => {
  it("refuses a colour the reader cannot read", () => {
    expect(argumentError(() => Mermaid.resolveTheme({ nodeFill: "zz" })).field).toBe(
      "options.nodeFill"
    );
    expect(argumentError(() => Mermaid.resolveTheme({ palette: ["#fff", "nope"] })).field).toBe(
      "options.palette[1]"
    );
    expect(() => Mermaid.toSvg("graph TD\nA-->B", { theme: { edge: "not-a-colour" } })).toThrow(
      MermaidArgumentError
    );
  });

  it("still accepts every form the reader reads", () => {
    for (const colour of [
      "#f00",
      "#ff000080",
      "red",
      "rgb(1, 2, 3)",
      "hsl(0 50% 50%)",
      "transparent",
      "none"
    ]) {
      expect(() => Mermaid.resolveTheme({ nodeFill: colour })).not.toThrow();
    }
  });
});

describe("handles are checked (X10)", () => {
  it("the deprecated layoutFlowchart and centredText", () => {
    for (const bad of [42, null, {}]) {
      expect(() => layoutFlowchart(bad as never)).toThrow(MermaidArgumentError);
    }
    const black = { r: 0, g: 0, b: 0, a: 1 };
    expect(() => centredText(null as never, ["a"], 10, "Arial", black)).toThrow(
      MermaidArgumentError
    );
    expect(() => centredText(42 as never, ["a"], 10, "Arial", black)).toThrow(MermaidArgumentError);
  });
});

describe("refusals read as every other module's do", () => {
  it("name the call and the field, and show the value", () => {
    const fontSize = argumentError(() => Mermaid.toSvg("flowchart TD\nA-->B", { fontSize: NaN }));
    expect(fontSize.field).toBe("options.fontSize");
    expect(fontSize.message).toBe(
      "Mermaid.toSvg: options.fontSize must be a finite number > 0, got NaN"
    );
    const box = argumentError(() =>
      Mermaid.createCenteredText({ x: 0, y: 0, width: 1, height: 1 }, ["a", 2 as never], {
        size: 10
      })
    );
    expect(box.field).toBe("lines[1]");
    expect(box.message).toBe("Mermaid.createCenteredText: lines[1] must be a string, got 2");
  });

  it("refuse an array box at the box, not at its missing x", () => {
    const err = argumentError(() => Mermaid.createCenteredText([] as never, ["a"], { size: 10 }));
    expect(err.field).toBe("box");
    expect(err.message).toBe(
      "Mermaid.createCenteredText: box must be { x, y, width, height }, got an array"
    );
  });

  it("name the palette entry that is not a colour", () => {
    const err = argumentError(() => Mermaid.resolveTheme({ palette: ["#fff", 3 as never] }));
    expect(err.field).toBe("options.palette[1]");
    expect(err.message).toMatch(
      /^Mermaid\.resolveTheme: options\.palette\[1\] must be a CSS colour .*, got 3$/
    );
  });
});
