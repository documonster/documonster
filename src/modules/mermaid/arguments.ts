/**
 * Argument checks shared by more than one public Mermaid entry.
 */

import { MermaidArgumentError } from "@mermaid/errors";

/** A text box `{ x, y, width, height }` of finite numbers, and its lines of text. */
export function checkTextBox(where: string, box: unknown, lines: unknown): void {
  if (typeof box !== "object" || box === null) {
    throw new MermaidArgumentError("box", `${where}: box must be { x, y, width, height }.`);
  }
  const b = box as Record<string, unknown>;
  for (const key of ["x", "y", "width", "height"] as const) {
    if (typeof b[key] !== "number" || !Number.isFinite(b[key])) {
      throw new MermaidArgumentError(
        `box.${key}`,
        `${where}: box.${key} must be a finite number; got ${String(b[key])}.`
      );
    }
  }
  if (!Array.isArray(lines) || !lines.every(line => typeof line === "string")) {
    throw new MermaidArgumentError("lines", `${where}: lines must be an array of strings.`);
  }
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
