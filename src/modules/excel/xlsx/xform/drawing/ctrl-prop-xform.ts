import type { CheckboxState, FormCheckboxModel } from "@excel/core/form-control";
import { formCheckboxStateFromValue } from "@excel/core/form-control";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";
import { XmlWriter } from "@xml/writer";

/**
 * Control Properties Xform - Generates ctrlProp*.xml for form controls
 *
 * Each form control (checkbox, button, etc.) has an associated ctrlProp file
 * that stores its properties like objectType, checked state, and linked cell.
 */

/** The properties of one control, as `xl/ctrlProps/ctrlPropN.xml` states them. */
interface CtrlPropModel {
  /** `CheckBox`, `Button`, `Drop`, … */
  objectType?: string;
  checked: CheckboxState;
  link?: string;
  noThreeD: boolean;
  /** The part as read, kept for a control this library does not model. */
  xml?: string;
  /** Every attribute of `formControlPr`, for matching the part to the VML shape it describes. */
  attributes?: Record<string, string>;
}

class CtrlPropXform extends BaseXform<FormCheckboxModel> {
  private _checkedToXmlValue(checked: FormCheckboxModel["checked"]): string {
    switch (checked) {
      case "Checked":
        return "1";
      case "Mixed":
        return "2";
      case "Unchecked":
      default:
        return "0";
    }
  }

  get tag(): string {
    return "formControlPr";
  }

  render(xmlStream: XmlSink, model?: FormCheckboxModel): void {
    const renderModel = (model || this.model)!;

    const attrs: Record<string, string> = {
      xmlns: "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
      objectType: "CheckBox",
      // Excel tends to serialize this as numeric state (0/1/2), matching VML x:Checked.
      checked: this._checkedToXmlValue(renderModel.checked),
      lockText: "1"
    };

    // Add linked cell reference
    if (renderModel.link) {
      attrs.fmlaLink = renderModel.link;
    }

    // Add noThreeD for flat appearance
    if (renderModel.noThreeD) {
      attrs.noThreeD = "1";
    }

    xmlStream.openXml({ version: "1.0", encoding: "UTF-8", standalone: "yes" });
    xmlStream.leafNode(this.tag, attrs);
  }

  /**
   * Generate XML string directly (convenience method)
   * Uses render() internally to ensure consistency
   */
  toXml(model: FormCheckboxModel): string {
    const xmlStream = new XmlWriter();
    this.render(xmlStream, model);
    return xmlStream.xml;
  }

  /** What a parsed `formControlPr` says; read by the worksheet reconcile. */
  parsed?: CtrlPropModel;

  parseOpen(node: ParseOpenTag): boolean {
    // The part's default namespace is x14's, which the read normalises to its `x14:` prefix.
    if (node.name === this.tag || node.name === `x14:${this.tag}`) {
      const { objectType, checked, fmlaLink, noThreeD } = node.attributes;
      this.parsed = {
        objectType,
        checked: formCheckboxStateFromValue(checked),
        ...(fmlaLink ? { link: fmlaLink } : {}),
        noThreeD: noThreeD === "1" || noThreeD === "true",
        attributes: Object.fromEntries(
          Object.entries(node.attributes).filter(([name]) => !name.startsWith("xmlns"))
        )
      };
    }
    return true;
  }

  parseText(): void {
    // formControlPr carries everything a checkbox needs in attributes.
  }

  parseClose(name: string): boolean {
    return name !== this.tag && name !== `x14:${this.tag}`;
  }
}

export { CtrlPropXform };
export type { CtrlPropModel };
