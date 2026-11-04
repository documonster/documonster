import type { FormControlAnchor } from "@excel/core/form-control";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { EchoNode } from "@excel/xlsx/xform/drawing/generic-echo-xform";
import { GenericEchoXform } from "@excel/xlsx/xform/drawing/generic-echo-xform";
import { parseXsdBoolean } from "@excel/xlsx/xform/xsd-values";
import type { ParseOpenTag } from "@xml/types";

/** One `<control>` of a sheet's `<controls>`: the bridge from a VML shape to its ctrlProp part. */
interface SheetControlEntry {
  shapeId: number;
  /** The sheet relationship reaching `xl/ctrlProps/ctrlPropN.xml`. */
  rId?: string;
  name?: string;
  /** `<controlPr print>`; absent means printed. */
  print?: boolean;
  /** `<controlPr><anchor>`, in EMU — exact where the VML anchor is rounded to pixels. */
  tl?: FormControlAnchor;
  br?: FormControlAnchor;
  /** `<controlPr><anchor moveWithCells sizeWithCells>`, when there is an `<anchor>`; both false when absent. */
  moveWithCells?: boolean;
  sizeWithCells?: boolean;
  /** The `<control>` element itself, for a control this library keeps rather than models. */
  xml?: EchoNode;
}

const CORNER_FIELDS: Record<string, keyof FormControlAnchor> = {
  "xdr:col": "col",
  "xdr:colOff": "colOff",
  "xdr:row": "row",
  "xdr:rowOff": "rowOff"
};

/**
 * Reads a sheet's `<controls>` (Office 2010+, inside `mc:AlternateContent`). Parse only: the writer emits
 * the element itself from the form-control model.
 *
 * The wrapping `mc:AlternateContent` elements are not handled here: the sheet xform ignores elements it
 * has no xform for, so `<controls>` is reached by its own name, and each `<control>` is found whether or
 * not it sits in a nested `mc:Choice`.
 */
class SheetControlsXform extends BaseXform<SheetControlEntry[]> {
  private entry?: SheetControlEntry;
  private corner?: Partial<FormControlAnchor>;
  private field?: keyof FormControlAnchor;
  private value = "";
  private echo?: GenericEchoXform;

  get tag(): string {
    return "controls";
  }

  parseOpen(node: ParseOpenTag): boolean {
    this.echo?.parseOpen(node);
    switch (node.name) {
      case this.tag:
        this.model ??= [];
        break;
      case "control": {
        const shapeId = parseInt(node.attributes.shapeId ?? "", 10);
        this.entry = Number.isFinite(shapeId)
          ? {
              shapeId,
              ...(node.attributes["r:id"] ? { rId: node.attributes["r:id"] } : {}),
              ...(node.attributes.name ? { name: node.attributes.name } : {})
            }
          : undefined;
        if (this.entry) {
          this.echo = new GenericEchoXform("control");
          this.echo.parseOpen(node);
        }
        break;
      }
      case "controlPr":
        if (this.entry && node.attributes.print !== undefined) {
          this.entry.print = !/^(0|false)$/i.test(node.attributes.print);
        }
        break;
      case "anchor":
        if (this.entry) {
          this.entry.moveWithCells = parseXsdBoolean(node.attributes.moveWithCells) ?? false;
          this.entry.sizeWithCells = parseXsdBoolean(node.attributes.sizeWithCells) ?? false;
        }
        break;
      case "from":
      case "to":
        this.corner = {};
        break;
      default:
        if (this.corner && node.name in CORNER_FIELDS) {
          this.field = CORNER_FIELDS[node.name];
          this.value = "";
        }
    }
    return true;
  }

  parseText(text: string): void {
    this.echo?.parseText(text);
    if (this.field) {
      this.value += text;
    }
  }

  parseClose(name: string): boolean {
    if (this.echo && !this.echo.parseClose(name)) {
      if (this.entry && this.echo.model) {
        this.entry.xml = this.echo.model;
      }
      this.echo = undefined;
    }
    switch (name) {
      case this.tag:
        return false;
      case "control":
        if (this.entry) {
          // Excel may also list the control in an `mc:Fallback`; the first entry for a shape is kept.
          if (!this.model!.some(entry => entry.shapeId === this.entry!.shapeId)) {
            this.model!.push(this.entry);
          }
          this.entry = undefined;
        }
        return true;
      case "from":
      case "to": {
        const corner = this.corner;
        this.corner = undefined;
        if (
          this.entry &&
          corner &&
          Object.values(CORNER_FIELDS).every(field => Number.isFinite(corner[field]))
        ) {
          this.entry[name === "from" ? "tl" : "br"] = corner as FormControlAnchor;
        }
        return true;
      }
      default:
        if (this.field && this.corner) {
          this.corner[this.field] = parseInt(this.value.trim(), 10);
          this.field = undefined;
        }
        return true;
    }
  }

  reset(): void {
    super.reset();
    this.model = undefined;
    this.echo = undefined;
  }
}

export { SheetControlsXform };
export type { SheetControlEntry };
