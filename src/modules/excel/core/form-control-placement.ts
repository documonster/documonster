/**
 * How a form control follows the cells under it, in a drawing anchor's terms — kept apart from `form-control`, which
 * builds and edits checkboxes, because every write needs it and a write-only bundle needs none of the rest.
 */
import type { FormCheckboxModel, LegacyControlModel } from "@excel/core/form-control";
import type { EchoNode } from "@excel/xlsx/xform/drawing/generic-echo-xform";

/** How a modelled checkbox follows the cells under it. */
export function formCheckboxPlacement(
  model: FormCheckboxModel
): "oneCell" | "twoCell" | "absolute" {
  return model.editAs ?? "oneCell";
}

/**
 * How a control follows the cells under it, from the two places a file says so: the `<control>` entry's
 * `<anchor moveWithCells sizeWithCells>` (both false when absent) or, without one, the VML shape's
 * `x:MoveWithCells` / `x:SizeWithCells` — which say it the other way round, present when the object does
 * *not* move or size.
 */
export function controlPlacement(
  moveWithCells: boolean,
  sizeWithCells: boolean
): "oneCell" | "twoCell" | "absolute" {
  return sizeWithCells ? "twoCell" : moveWithCells ? "oneCell" : "absolute";
}

/** A preserved control's placement, from its `<control>` entry or else its VML (see {@link controlPlacement}). */
export function legacyControlPlacement(
  control: LegacyControlModel
): "oneCell" | "twoCell" | "absolute" {
  const child = (node: EchoNode | undefined, tag: string) =>
    node?.children.find(c => c.tag === tag);
  const anchor = child(child(control.control, "controlPr"), "anchor");
  if (anchor) {
    const yes = (value: unknown) =>
      value === "1" || value === "true" || value === 1 || value === true;
    return controlPlacement(yes(anchor.attrs.moveWithCells), yes(anchor.attrs.sizeWithCells));
  }
  const data = child(control.vmlShape, "x:ClientData");
  return controlPlacement(
    child(data, "x:MoveWithCells") === undefined,
    child(data, "x:SizeWithCells") === undefined
  );
}
