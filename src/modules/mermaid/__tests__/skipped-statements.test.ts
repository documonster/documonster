import { Mermaid, MermaidError } from "@mermaid/index";
import type { GanttDiagram, MermaidDiagram } from "@mermaid/index";
import { parseDate } from "@mermaid/parse/dates";
import { describe, expect, it } from "vitest";

const skippedOf = (source: string): MermaidDiagram["skipped"] => Mermaid.parse(source).skipped;

describe("Mermaid.parse reports the statements it skipped", () => {
  it("leaves a clean source without a report", () => {
    expect(skippedOf("flowchart TD\n  A[Start] --> B{Ready?}\n  B -->|yes| C")).toBeUndefined();
    expect("skipped" in Mermaid.parse('pie\n  "A" : 1')).toBe(false);
  });

  it("reports a flowchart statement it could not read, with its position", () => {
    // Each of these used to draw part of the statement, or nothing, without a word.
    expect(skippedOf("flowchart TD\n  A-->")).toEqual([{ text: "A-->", line: 2, column: 3 }]);
    expect(skippedOf("flowchart TD\nA-->B\nthis is ]]] garbage {{\nC-->D")).toEqual([
      { text: "this is ]]] garbage {{", line: 3, column: 1 }
    ]);
    expect(skippedOf("flowchart TD\nA[Start-->B")).toEqual([
      { text: "A[Start-->B", line: 2, column: 1 }
    ]);
  });

  it("gives the column of a statement after a semicolon", () => {
    expect(skippedOf("flowchart TD\nA-->B; C-->")).toEqual([{ text: "C-->", line: 2, column: 8 }]);
  });

  it("does not report what it understood but does not draw", () => {
    expect(
      skippedOf("flowchart TD\nA-->B\nstyle A fill:#f00\nclick A callback\nclassDef x fill:red")
    ).toBeUndefined();
    expect(
      skippedOf(
        "stateDiagram-v2\n[*] --> Still\nStill\nnote right of Still\n  some text\nend note\nstate Moving\n--"
      )
    ).toBeUndefined();
    expect(
      skippedOf("sequenceDiagram\nAlice->>Bob: hi\nactivate Bob\nNote over Bob: ok")
    ).toBeUndefined();
  });

  it("reports across diagram types", () => {
    expect(skippedOf("sequenceDiagram\nAlice-Bob: hi\nAlice->>Bob: ok")).toEqual([
      { text: "Alice-Bob: hi", line: 2, column: 1 }
    ]);
    expect(skippedOf('pie\n"A" : abc\n"B" : 2')).toEqual([
      { text: '"A" : abc', line: 2, column: 1 }
    ]);
    expect(skippedOf("classDiagram\nclass A {\n  +ok()\n}\n!!! nope")).toEqual([
      { text: "!!! nope", line: 5, column: 1 }
    ]);
    expect(skippedOf("erDiagram\nA ||--o{ B : has\n!!!")).toEqual([
      { text: "!!!", line: 3, column: 1 }
    ]);
    expect(skippedOf("journey\ntitle J\nsection S\nTask without score")).toEqual([
      { text: "Task without score", line: 4, column: 1 }
    ]);
    expect(skippedOf('gitGraph\ncommit\ncherry-pick id: "x"')).toEqual([
      { text: 'cherry-pick id: "x"', line: 3, column: 1 }
    ]);
    expect(skippedOf("xychart-beta\nx-axis [a, b]\nbar [1, oops]")).toEqual([
      { text: "bar [1, oops]", line: 3, column: 1 }
    ]);
    expect(skippedOf("sankey-beta\nA,B,10\nA,C,-3")).toEqual([
      { text: "A,C,-3", line: 3, column: 1 }
    ]);
    expect(skippedOf('packet-beta\n0-15: "Src"\nnonsense')).toEqual([
      { text: "nonsense", line: 3, column: 1 }
    ]);
    expect(skippedOf("architecture-beta\nservice a(server)[A]\n???")).toEqual([
      { text: "???", line: 3, column: 1 }
    ]);
    expect(skippedOf('C4Context\nPerson(a, "A")\nRel(a)')).toEqual([
      { text: "Rel(a)", line: 3, column: 1 }
    ]);
    expect(skippedOf("quadrantChart\nA: [0.3, 0.6]\nB: [x, 1]")).toEqual([
      { text: "B: [x, 1]", line: 3, column: 1 }
    ]);
  });
});

describe("gantt dates that do not exist", () => {
  it("are not dates", () => {
    // `Date.UTC` rolled 2024-13-45 over to 2025-02-14.
    expect(parseDate("2024-13-45")).toBeUndefined();
    expect(parseDate("2023-02-29")).toBeUndefined();
    expect(parseDate("2024-02-29")).toBe(Date.UTC(2024, 1, 29));
    expect(parseDate("2024-01-01 25:00")).toBeUndefined();
    // A year below 100 is that year, not 19xx.
    expect(new Date(parseDate("0099-01-01")!).getUTCFullYear()).toBe(99);
  });

  it("leave the task out and report it, instead of drawing it somewhere else", () => {
    const diagram = Mermaid.parse(
      "gantt\ndateFormat YYYY-MM-DD\nsection S\nA :a1, 2024-01-01, 3d\nB :b1, 2024-13-45, 3d"
    ) as GanttDiagram;
    expect(diagram.tasks.map(task => task.label)).toEqual(["A"]);
    expect(diagram.skipped).toEqual([{ text: "B :b1, 2024-13-45, 3d", line: 5, column: 1 }]);
  });

  it("report a dependency on a task that does not exist", () => {
    const diagram = Mermaid.parse(
      "gantt\nsection S\nA :a1, 2024-01-01, 3d\nB :b1, after nope, 2d"
    ) as GanttDiagram;
    expect(diagram.skipped).toEqual([{ text: "B :b1, after nope, 2d", line: 4, column: 1 }]);
  });
});

describe("Mermaid render options are checked", () => {
  const source = "flowchart TD\nA-->B";

  it("rejects options that used to produce an empty or wrong picture", () => {
    // `fontSize: NaN` produced `<svg width="0" height="0">`.
    expect(() => Mermaid.toSvg(source, { fontSize: Number.NaN })).toThrow(MermaidError);
    expect(() => Mermaid.toSvg(source, { fontSize: -5 })).toThrow(/fontSize/);
    expect(() => Mermaid.toSvg(source, null as never)).toThrow(MermaidError);
    expect(() => Mermaid.toSvg(source, { theme: 42 as never })).toThrow(/theme/);
    expect(() => Mermaid.toDrawList(source, { rankGap: Infinity })).toThrow(/rankGap/);
    expect(() => Mermaid.layoutGraph(Mermaid.parse(source) as never, { nodeGap: -1 })).toThrow(
      /nodeGap/
    );
  });

  it("refuses a theme token that does not exist", () => {
    // Mermaid's own name for the node fill; it used to change nothing.
    expect(() => Mermaid.resolveTheme({ primaryColor: "#f00" } as never)).toThrow(/primaryColor/);
    expect(() => Mermaid.resolveTheme(null as never)).toThrow(MermaidError);
    expect(() => Mermaid.resolveTheme({ nodeFill: 1 } as never)).toThrow(/nodeFill/);
    expect(() => Mermaid.resolveTheme({ palette: [] })).toThrow(/palette/);
  });

  it("reads a named theme colour", () => {
    expect(Mermaid.resolveTheme({ nodeFill: "red" }).nodeFill).toEqual({ r: 1, g: 0, b: 0, a: 1 });
    expect(Mermaid.resolveTheme({ nodeFill: "transparent" }).nodeFill.a).toBe(0);
  });

  it("checks createCenteredText", () => {
    const box = { x: 0, y: 0, width: 100, height: 40 };
    expect(() => Mermaid.createCenteredText(box, ["a"], { size: Number.NaN })).toThrow(
      /style.size/
    );
    expect(() => Mermaid.createCenteredText(box, "abc" as never, { size: 12 })).toThrow(/lines/);
    expect(() =>
      Mermaid.createCenteredText({ ...box, x: Number.NaN }, ["a"], { size: 12 })
    ).toThrow(/box.x/);
  });
});
