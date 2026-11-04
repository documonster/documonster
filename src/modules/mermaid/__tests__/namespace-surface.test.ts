import * as MermaidEntry from "@mermaid/index";
import { Mermaid } from "@mermaid/index";
import { describe, expect, it } from "vitest";

describe("documonster/mermaid namespace", () => {
  it("draws through the namespace", () => {
    expect(Mermaid.toSvg("flowchart LR\n  A --> B").startsWith("<svg")).toBe(true);
    const diagram = Mermaid.parse("flowchart LR\n  A --> B");
    if (diagram.kind !== "flowchart") {
      throw new Error("expected a flowchart");
    }
    expect(Mermaid.layoutGraph(diagram).nodes.length).toBe(2);
  });

  it("layoutGraph lays out every graph-shaped diagram, sized as the renderer draws it", () => {
    const sources = {
      state: "stateDiagram-v2\n  [*] --> Idle\n  Idle --> Busy",
      class:
        "classDiagram\n  class Animal {\n    +name: string\n    +eat()\n  }\n  Animal <|-- Dog",
      er: "erDiagram\n  CUSTOMER ||--o{ ORDER : places",
      requirement:
        "requirementDiagram\n  requirement r1 {\n    id: 1\n    text: works\n  }\n  element e1 {\n    type: code\n  }\n  e1 - satisfies -> r1",
      c4: 'C4Context\n  Person(u, "User")\n  System(s, "System")\n  Rel(u, s, "Uses")',
      architecture:
        "architecture-beta\n  service a(server)[A]\n  service b(database)[B]\n  a:R --> L:b"
    };
    for (const [kind, source] of Object.entries(sources)) {
      const diagram = Mermaid.parse(source);
      expect(diagram.kind).toBe(kind);
      if (diagram.kind !== kind) {
        continue;
      }
      const layout = Mermaid.layoutGraph(diagram as Parameters<typeof Mermaid.layoutGraph>[0]);
      expect(layout.nodes.length, kind).toBeGreaterThanOrEqual(2);
      // The same geometry the renderer draws.
      const list = Mermaid.toDrawList(diagram);
      expect([list.width, list.height], kind).toEqual([layout.width, layout.height]);
    }
  });

  it("keeps the deprecated flat names as forwards", () => {
    expect(MermaidEntry.mermaidToDrawList).toBe(Mermaid.toDrawList);
    expect(MermaidEntry.mermaidToSvg).toBe(Mermaid.toSvg);
    expect(MermaidEntry.parseMermaid).toBe(Mermaid.parse);
    const flowchart = Mermaid.parse("flowchart LR\n  A --> B");
    if (flowchart.kind !== "flowchart") {
      throw new Error("expected a flowchart");
    }
    expect(MermaidEntry.layoutFlowchart(flowchart)).toEqual(Mermaid.layoutGraph(flowchart));
    expect(MermaidEntry.centredText).toBeTypeOf("function");
    expect(MermaidEntry.resolveTheme).toBe(Mermaid.resolveTheme);
  });
});
