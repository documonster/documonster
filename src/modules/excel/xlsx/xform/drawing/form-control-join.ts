import type { FormCheckboxModel, LegacyControlModel } from "@excel/core/form-control";
import { controlPlacement } from "@excel/core/form-control-placement";
import type { CtrlPropModel } from "@excel/xlsx/xform/drawing/ctrl-prop-xform";
import type { SheetControlEntry } from "@excel/xlsx/xform/sheet/controls-xform";

/**
 * A sheet's checkboxes from the three places a file describes them.
 *
 * The VML shape is the base: it draws the control and carries its state, link, caption and a
 * pixel-rounded anchor. The sheet's `<controls>` entry for the same shape id — absent in an XLSB, and in
 * files written before Excel 2010 — adds the name and an exact EMU anchor. Its ctrlProp part holds the
 * state Excel 2010+ reads when it differs from the VML, so where it exists it wins for the state, the
 * link and the 3-D flag.
 */
export function joinFormControls(
  vmlControls: readonly FormCheckboxModel[],
  entries: readonly SheetControlEntry[],
  ctrlPropFor: (rId: string) => CtrlPropModel | undefined
): FormCheckboxModel[] {
  const byShapeId = new Map(entries.map(entry => [entry.shapeId, entry]));
  return vmlControls.map(control => {
    const entry = byShapeId.get(control.shapeId);
    const props = entry?.rId === undefined ? undefined : ctrlPropFor(entry.rId);
    const joined: FormCheckboxModel = { ...control, tl: { ...control.tl }, br: { ...control.br } };
    if (entry !== undefined) {
      if (entry.name !== undefined) {
        joined.name = entry.name;
      }
      if (entry.print !== undefined) {
        joined.print = entry.print;
      }
      if (entry.tl !== undefined && entry.br !== undefined) {
        joined.tl = { ...entry.tl };
        joined.br = { ...entry.br };
      }
      if (entry.moveWithCells !== undefined) {
        joined.editAs = controlPlacement(entry.moveWithCells, entry.sizeWithCells ?? false);
      }
    }
    if (props !== undefined) {
      joined.checked = props.checked;
      joined.noThreeD = props.noThreeD;
      if (props.link !== undefined) {
        joined.link = props.link;
      }
    }
    return joined;
  });
}

/**
 * A sheet's other legacy controls, joined the same way but kept rather than interpreted: each VML shape
 * takes the `<control>` entry naming its id and the ctrlProp part that entry reaches, verbatim. A shape
 * no entry names keeps only its VML, as an XLSB's always does.
 */
export function joinLegacyControls(
  vmlControls: readonly LegacyControlModel[],
  entries: readonly SheetControlEntry[],
  ctrlPropFor: (rId: string) => CtrlPropModel | undefined
): LegacyControlModel[] {
  const byShapeId = new Map(entries.map(entry => [entry.shapeId, entry]));
  return vmlControls.map(control => {
    const entry = byShapeId.get(control.shapeId);
    if (entry?.xml === undefined) {
      return { ...control };
    }
    const ctrlPropXml = entry.rId === undefined ? undefined : ctrlPropFor(entry.rId)?.xml;
    return {
      ...control,
      ...(entry.name === undefined ? {} : { name: entry.name }),
      ...(entry.tl === undefined || entry.br === undefined
        ? {}
        : { tl: { ...entry.tl }, br: { ...entry.br } }),
      control: entry.xml,
      ...(ctrlPropXml === undefined ? {} : { ctrlPropXml })
    };
  });
}
