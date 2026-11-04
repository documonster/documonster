/**
 * Pairing an XLSB's ctrlProp parts with the controls they describe — a read-only concern, kept out of
 * `form-control-join.ts` because the XLSX writer imports that statically and this needs the checkbox
 * state parser from `core/form-control`.
 */
import type { LegacyControlModel } from "@excel/core/form-control";
import { formCheckboxStateFromValue } from "@excel/core/form-control";
import type { EchoNode } from "@excel/xlsx/xform/drawing/generic-echo-xform";

/**
 * The `formControlPr` attributes that restate a VML `<x:ClientData>` child ([MS-XLSX] CT_FormControlPr
 * against [ISO/IEC29500-4] §19.4.2.12): a value compared as written, a flag as present-or-not.
 */
const VALUE_FIELDS: Record<string, string> = {
  fmlaLink: "x:FmlaLink",
  fmlaRange: "x:FmlaRange",
  fmlaGroup: "x:FmlaGroup",
  fmlaTxbx: "x:FmlaTxbx",
  val: "x:Val",
  min: "x:Min",
  max: "x:Max",
  inc: "x:Inc",
  page: "x:Page",
  sel: "x:Sel",
  selType: "x:SelType",
  dx: "x:Dx",
  dropLines: "x:DropLines",
  dropStyle: "x:DropStyle"
};
const FLAG_FIELDS: Record<string, string> = {
  horiz: "x:Horiz",
  noThreeD: "x:NoThreeD",
  firstButton: "x:FirstButton",
  multiSel: "x:MultiSel"
};

/** A shape's `<x:ClientData>` children, each tag with its trimmed text. */
function clientDataOf(shape: EchoNode): Map<string, string> {
  const clientData = shape.children.find(child => child.tag === "x:ClientData");
  const leaves = new Map<string, string>();
  for (const child of clientData?.children ?? []) {
    const content = child.content ?? (child.text === undefined ? [] : [child.text]);
    leaves.set(
      child.tag,
      content
        .filter(item => typeof item === "string")
        .join("")
        .trim()
    );
  }
  return leaves;
}

/** A reference or value as both parts may spell it: `$E$1` and `e1` agree. */
function normalise(value: string): string {
  return value.replace(/[$\s]/g, "").toLowerCase();
}

/** How many fields a ctrlProp part and a shape agree on, and how many they contradict. */
function agreement(
  attributes: Readonly<Record<string, string>>,
  clientData: ReadonlyMap<string, string>
): { agree: number; contradict: number } {
  let agree = 0;
  let contradict = 0;
  for (const [attribute, element] of Object.entries(VALUE_FIELDS)) {
    const prop = attributes[attribute];
    const vml = clientData.get(element);
    if (prop !== undefined && vml !== undefined) {
      if (normalise(prop) === normalise(vml)) {
        agree++;
      } else {
        contradict++;
      }
    }
  }
  for (const [attribute, element] of Object.entries(FLAG_FIELDS)) {
    const prop = /^(1|true)$/i.test(attributes[attribute] ?? "");
    if (prop === clientData.has(element)) {
      agree += prop ? 1 : 0;
    } else {
      contradict++;
    }
  }
  if (attributes.checked !== undefined || clientData.has("x:Checked")) {
    if (
      formCheckboxStateFromValue(attributes.checked) ===
      formCheckboxStateFromValue(clientData.get("x:Checked"))
    ) {
      agree++;
    } else {
      contradict++;
    }
  }
  return { agree, contradict };
}

/**
 * Give each preserved control the ctrlProp part that describes it, in a package that does not say which
 * part that is — an XLSB.
 *
 * [MS-XLSB] has no key for it. The worksheet's records name an ActiveX control's part (`BrtActiveX`
 * carries a shape id and a relationship id, and requires `ObjectType="Pict"`), but no record names a form
 * control's; §2.1.7 lists a Control Properties part as the target only of a "SpreadsheetML control" — the
 * XLSX `<control>` element — and §2.1.7.62 does not list it among the worksheet's relationships at all.
 * What an XLSB does carry is the control itself, whole, in its VML `<x:ClientData>`, and a ctrlProp part
 * restates those same fields. So each part goes to the control of its type whose shape agrees with it:
 * none of the shared fields contradicted, as many as possible agreeing. Only controls that agree equally —
 * whose shared fields are identical, so either pairing says the same thing — fall back to relationship
 * order, matched against shape order, which is the order this library's writer allocates both in. A part
 * that contradicts every control of its type describes none of them and is not kept.
 */
export function pairCtrlProps(
  controls: readonly LegacyControlModel[],
  ctrlProps: readonly {
    readonly xml: string;
    readonly attributes: Readonly<Record<string, string>>;
  }[]
): void {
  const clientData = controls.map(control => clientDataOf(control.vmlShape));
  const candidates: { prop: number; control: number; agree: number }[] = [];
  ctrlProps.forEach((ctrlProp, prop) => {
    const type = ctrlProp.attributes.objectType?.toLowerCase();
    controls.forEach((control, index) => {
      if (control.ctrlPropXml !== undefined || control.objectType.toLowerCase() !== type) {
        return;
      }
      const { agree, contradict } = agreement(ctrlProp.attributes, clientData[index]);
      if (contradict === 0) {
        candidates.push({ prop, control: index, agree });
      }
    });
  });
  // Best agreement first; among equals, relationship order against shape order.
  candidates.sort((a, b) => b.agree - a.agree || a.prop - b.prop || a.control - b.control);
  const takenProps = new Set<number>();
  const takenControls = new Set<number>();
  for (const { prop, control } of candidates) {
    if (takenProps.has(prop) || takenControls.has(control)) {
      continue;
    }
    takenProps.add(prop);
    takenControls.add(control);
    (controls[control] as { ctrlPropXml?: string }).ctrlPropXml = ctrlProps[prop].xml;
  }
}
