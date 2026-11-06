/**
 * Argument checks shared by more than one public Mermaid entry.
 *
 * Options objects are checked against schemas (`@utils/schema`), each pinned to its public
 * interface with `SchemaMatches`, so a refusal names the call and the field the same way in
 * every entry: `Mermaid.toSvg: options.fontSize must be a finite number > 0, got NaN`.
 */

import { MermaidArgumentError } from "@mermaid/errors";
import type { LayoutOptions } from "@mermaid/layout/flowchart";
import { check, s } from "@utils/schema";
import type { SchemaCheckOptions, SchemaMatches } from "@utils/schema";
import { isNoPaintColor, parseCssColor } from "@utils/svg-lex";

function argumentError(field: string, message: string): MermaidArgumentError {
  return new MermaidArgumentError(field, message);
}

/** Where a checked value sits: the public call and the path to the value. */
export function at(where: string, field: string): SchemaCheckOptions {
  return { where, field, error: argumentError };
}

const finite = () => s.number({ finite: true });
const positive = () =>
  s
    .number({ finite: true, expected: "a finite number > 0" })
    .refine(v => v > 0, "a finite number > 0");
const nonNegative = () => s.number({ finite: true, min: 0 });
export const optionsObject = { expected: "an options object" };

const CSS_COLOUR =
  "a CSS colour (#rgb, #rrggbb, #rrggbbaa, rgb(), hsl(), a colour name, transparent or none)";

/** A colour string `cssColor` reads — one it cannot read would paint black instead. */
export const colour = () =>
  s
    .string({ expected: CSS_COLOUR })
    .refine(v => isNoPaintColor(v) || parseCssColor(v) !== undefined, CSS_COLOUR);

/** The keys of `LayoutOptions`, for the options objects that extend it. */
export const layoutShape = {
  fontSize: positive().optional(),
  fontFamily: s.string().optional(),
  rankGap: nonNegative().optional(),
  nodeGap: nonNegative().optional(),
  maxLabelWidth: positive().optional(),
  maxEdgeLabelWidth: positive().optional(),
  padding: nonNegative().optional(),
  measureNode: s.function<NonNullable<LayoutOptions["measureNode"]>>().optional()
};
const layoutOptionsSchema = s.object(layoutShape, { ...optionsObject, readonly: true });
const _layoutOptionsMatch: SchemaMatches<typeof layoutOptionsSchema, LayoutOptions> = true;
void _layoutOptionsMatch;

/**
 * Check `LayoutOptions`. A `NaN` font size used to lay the diagram out as a 0×0 picture, and
 * a negative one as a plausible-looking but wrong one.
 */
export function checkLayoutOptions(where: string, options: unknown): void {
  check(layoutOptionsSchema, options, at(where, "options"));
}

const textBoxSchema = s.object(
  { x: finite(), y: finite(), width: finite(), height: finite() },
  { expected: "{ x, y, width, height }" }
);
const linesSchema = s.readonlyArray(s.string(), { expected: "an array of strings" });

/** A text box `{ x, y, width, height }` of finite numbers, and its lines of text. */
export function checkTextBox(where: string, box: unknown, lines: unknown): void {
  check(textBoxSchema, box, at(where, "box"));
  check(linesSchema, lines, at(where, "lines"));
}

const textStyleSizeSchema = s.object(
  { size: positive() },
  { expected: "a text style { size, … }" }
);
const sourceSchema = s.string({ expected: "Mermaid text" });

/** A text style, of which only the size is read: the rest is drawn as given. */
export function checkTextStyleSize(where: string, style: unknown): void {
  check(textStyleSizeSchema, style, at(where, "style"));
}

/** The text of `Mermaid.parse`, refused by name when it is not text. */
export function checkSource(source: unknown): void {
  check(sourceSchema, source, at("Mermaid.parse", "source"));
}

/** A value `Mermaid.parse` returned for a flowchart. */
export function checkFlowchart(where: string, diagram: unknown): void {
  const kind =
    typeof diagram === "object" && diagram !== null
      ? (diagram as { kind?: unknown }).kind
      : undefined;
  if (kind !== "flowchart" || !Array.isArray((diagram as { nodes?: unknown }).nodes)) {
    throw new MermaidArgumentError(
      "diagram",
      `${where} takes a flowchart diagram from Mermaid.parse; got ${
        diagram === null
          ? "null"
          : typeof kind === "string"
            ? `a "${kind}" diagram`
            : typeof diagram
      }.`
    );
  }
}
