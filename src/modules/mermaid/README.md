# Mermaid Module

[中文](README_zh.md)

Mermaid diagram text to SVG, pixels, or a PDF page. Twenty-one diagram types, zero dependencies.

```typescript
import { Mermaid } from "documonster/mermaid";
```

## Features

- **Twenty-one diagram types** — flowcharts, state, class, ER, sequence, Gantt, mindmap, git graph, and more
- **No browser, no headless Chrome** — Layout and rendering are pure TypeScript; there is nothing to install and nothing to launch
- **Every backend for free** — The module produces a `DrawList` and implements no renderer, so markup, pixels and PDF pages all come from the same output
- **Layered layout** — Rank, order to reduce crossings, straighten, route: the same algorithm family Mermaid itself uses
- **Stop at any pass** — `Mermaid.parse` for the syntax tree, `Mermaid.layoutGraph` for coordinates, `Mermaid.toDrawList` for the picture
- **Themeable** — Mermaid's own default palette, token for token, with per-token overrides
- **Zero Dependencies** — Pure TypeScript, no external packages
- **Cross-Platform** — Same API in Node.js and browsers

---

## Quick Start

```typescript
import { Mermaid } from "documonster/mermaid";

const svg = Mermaid.toSvg(`flowchart TD
  A[Start] --> B{Ready?}
  B -->|yes| C[Ship it]
  B -->|no| A`);
```

### Any backend, from the same list

The module is a _producer_: it stops at the display list. Nothing in it knows what SVG or a PDF page is, which is exactly what lets one implementation serve every output.

```typescript
import { Mermaid } from "documonster/mermaid";
import { Draw } from "documonster/draw";

const list = Mermaid.toDrawList('pie title Split\n "A" : 60\n "B" : 40');

const svg = Draw.toSvg(list); // markup
const pixels = Draw.toRgba(list, { scale: 2 }); // RGBA
```

A diagram whose labels are not Latin needs a font the rasteriser can find. `Draw.toRgba` discovers one on Node and reports what it could not draw; in a browser you have to supply the bytes. See [the `draw` README](../draw/README.md#fonts-and-how-cjk-gets-drawn).

```typescript
import { Draw } from "documonster/draw";
import { Pdf } from "documonster/pdf";

const doc = new Pdf.Builder();
const page = doc.addPage({ width: 595, height: 842 });
Draw.render(
  list,
  Pdf.createDrawSurface(page, { x: 72, y: 500, width: list.width, height: list.height })
);
const bytes = await doc.toBuffer();
```

---

## Supported diagrams

Twenty-one types, but far fewer layouts — because most of them are the same picture in a different language.

**Laid out as a directed graph** (ranked, ordered, straightened, routed). Only the box shapes and the marks at the ends of the edges differ, which is the part that genuinely differs:

`flowchart` · `stateDiagram` · `classDiagram` · `erDiagram` · `requirementDiagram` · `C4` · `architecture`

**Laid out on their own terms** — a Gantt bar's position comes from the calendar, a journey's dot from a score, a mind-map node from the height of its own subtree, a commit from its lane and its ordinal. None of those is a graph, so none is forced through the graph layout:

`sequenceDiagram` · `gantt` · `timeline` · `journey` · `kanban` · `mindmap` · `gitGraph` · `quadrantChart` · `xychart` · `radar` · `sankey` · `packet` · `block` · `pie`

A box that sizes itself — a class compartment stack, a C4 element with its technology line — tells the layout its size through `LayoutOptions.measureNode`.

---

## Three passes

The parser, the layout and the renderer are separate, and a caller can stop after any of them.

```typescript
import { Mermaid } from "documonster/mermaid";

// 1. Text → syntax tree. Says nothing about geometry.
const diagram = Mermaid.parse("flowchart LR\n  A --> B");

// 2. Tree → coordinates, for the graph-shaped diagrams.
if (diagram.kind === "flowchart") {
  const layout = Mermaid.layoutGraph(diagram, { rankGap: 60 });
  console.log(layout.width, layout.height, layout.nodes, layout.edges);
}

// 3. Tree (or source) → display list.
const list = Mermaid.toDrawList(diagram);
```

Inspecting the tree or re-using the positions are real things to want, which is why both earlier passes are public rather than internal steps.

---

## Options

`Mermaid.toSvg` and `Mermaid.toDrawList` take `MermaidRenderOptions`, which extends `LayoutOptions`:

| Option          | Type                          | Description                                                       |
| --------------- | ----------------------------- | ----------------------------------------------------------------- |
| `theme`         | `ThemeOptions`                | Colour overrides; anything omitted keeps the default              |
| `background`    | `string`                      | Background fill. `"transparent"` (default) leaves no backdrop     |
| `fontSize`      | `number`                      | Label size                                                        |
| `fontFamily`    | `string`                      | Label family                                                      |
| `rankGap`       | `number`                      | Gap between adjacent ranks, along the flow                        |
| `nodeGap`       | `number`                      | Gap between siblings within a rank, across the flow               |
| `maxLabelWidth` | `number`                      | Longest label line before wrapping                                |
| `padding`       | `number`                      | Margin around the diagram                                         |
| `measureNode`   | `(node) => size \| undefined` | Size a node yourself, for a box design this module cannot measure |

The backdrop is `background` (or `theme.background`, which `Mermaid.resolveTheme` takes; `background` wins when both are given). Its default is transparent on purpose: a caller compositing the diagram onto something else gets what they asked for rather than an opaque rectangle.

Colours are CSS: `#rgb`, `#rrggbb`, `#rrggbbaa`, `rgb()`, `hsl()`, the named colours (`white`, `orange`, …) and `transparent`. A theme token this module does not have (Mermaid's own `primaryColor`, say) is refused with a `MermaidArgumentError` rather than ignored, and so is a colour none of those forms reads (`"zz"`, which used to be drawn black) and an option that has no reading — a `NaN` font size, a negative gap.

### Theming

```typescript
const svg = Mermaid.toSvg(source, {
  theme: {
    nodeFill: "#eef2ff",
    nodeStroke: "#4f46e5",
    nodeText: "#1e1b4b",
    edge: "#6366f1",
    palette: ["#4f46e5", "#0ea5e9", "#10b981"]
  },
  background: "#ffffff"
});
```

Tokens omitted fall back to Mermaid's own `base` theme, taken from the values it resolves to rather than eyeballed, so a diagram drawn here and one drawn there are recognisably the same diagram. Text drawn on a palette colour picks ink or paper from that colour's luminance — a palette holds both light and dark entries, and white on Mermaid's pale first slice is unreadable.

---

## Errors

```typescript
import { Mermaid, MermaidParseError } from "documonster/mermaid";

try {
  Mermaid.toSvg("nonsenseDiagram\n  A --> B");
} catch (error) {
  if (error instanceof MermaidParseError) {
    console.log(error.message, error.line, error.column); // … 1 1
  }
}
```

`MermaidParseError` is thrown when the source names no supported diagram type; `line` and
`column` (1-based) locate the statement that names it, and are `undefined` for an empty
source. It is the `*ParseError` every module's parser throws (`XmlParseError`,
`MarkdownParseError`, …). A wrong
argument or option — source that is not text, a malformed options object, a theme colour that
cannot be read, a value that is not a diagram — throws `MermaidArgumentError`. Its message
starts with the call (`Mermaid.toSvg: …`), and its `field` is the path from that call's
parameter name (`options.theme.nodeFill`, `options.fontSize`, `source`). Both extend `MermaidError`, the base class of every
error this module throws (itself a `BaseError`), and `isMermaidError(err)` is the type guard
for it.

Inside a supported diagram the parser is lenient: a statement it cannot read is left out
rather than refused. It is not left out silently — the parsed diagram lists it, with its
position, in `skipped`:

```typescript
const diagram = Mermaid.parse("flowchart TD\n  A --> B\n  B -->");
console.log(diagram.skipped); // [{ text: "B -->", line: 3, column: 3 }]
const svg = Mermaid.toSvg(diagram);
```

`skipped` is absent when every statement was understood. Statements that are understood but
not drawn (`style`, `click`, accessibility titles) are not listed; ones the renderer does not
draw at all (a git `cherry-pick`) are. A Gantt task whose date does not exist (`2024-13-45`)
is listed and not drawn, rather than placed somewhere else.

---

## Notes on the rendering

Two details are worth knowing, because both were bugs first:

- **A rank is re-centred after its overlaps are resolved.** Pushing nodes left-to-right until every pair clears the gap is one-directional and drags the whole rank right, so the trunk of a diagram visibly bends away from a fork. Translating the finished rank back by its mean displacement cannot reintroduce an overlap, because it moves every node equally.
- **A long edge is threaded, not routed around.** An edge spanning several ranks gets a chain of narrow placeholders, one per rank it crosses, so it takes part in the ordering like any other edge and is drawn straight through the lane left for it. Without them it has no say in how the intervening ranks are arranged and has to be routed afterwards around whatever ended up in the way.

Arrowheads are lowered to small filled triangles aimed along the last segment rather than expressed as SVG markers: a marker has no counterpart in a PDF content stream or a scanline rasteriser, and the triangle is the same picture in all three backends.

---

## Examples

`src/modules/mermaid/examples/gallery.ts` renders every supported diagram type in one pass — the fastest way to see what the module draws.
