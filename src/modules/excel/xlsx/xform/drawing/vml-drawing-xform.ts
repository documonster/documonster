import { StdDocAttributes } from "@xml/writer";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { VmlShapeXform } from "@excel/xlsx/xform/comment/vml-shape-xform";
import { FormCheckbox, type FormCheckboxModel } from "@excel/form-control";

/**
 * Unified VML Drawing Xform - Combines Notes (comments) and Form Controls
 *
 * Excel uses a single VML file per worksheet that can contain:
 * - Comment/note shapes (shapetype 202)
 * - Form control shapes (checkbox shapetype 201, etc.)
 *
 * This unified xform renders both into a single VML file.
 */

interface VmlDrawingModel {
  /** Comment/note shapes */
  comments?: any[];
  /** Form control checkboxes */
  formControls?: FormCheckboxModel[];
}

class VmlDrawingXform extends BaseXform<VmlDrawingModel> {
  declare public map: { [key: string]: any };
  declare public parser: any;

  constructor() {
    super();
    this.map = {
      "v:shape": new VmlShapeXform()
    };
    this.model = { comments: [], formControls: [] };
  }

  get tag(): string {
    return "xml";
  }

  /**
   * Render VML drawing containing both notes and form controls
   */
  render(xmlStream: any, model?: VmlDrawingModel): void {
    const renderModel = (model || this.model)!;
    const comments = renderModel.comments;
    const formControls = renderModel.formControls;
    const hasComments = comments && comments.length > 0;
    const hasFormControls = formControls && formControls.length > 0;

    xmlStream.openXml(StdDocAttributes);
    xmlStream.openNode(this.tag, VmlDrawingXform.DRAWING_ATTRIBUTES);

    // Shape layout - shared by both notes and form controls
    xmlStream.openNode("o:shapelayout", { "v:ext": "edit" });
    xmlStream.leafNode("o:idmap", { "v:ext": "edit", data: 1 });
    xmlStream.closeNode();

    // Shapetype 202 for notes/comments
    if (hasComments) {
      xmlStream.openNode("v:shapetype", {
        id: "_x0000_t202",
        coordsize: "21600,21600",
        "o:spt": 202,
        path: "m,l,21600r21600,l21600,xe"
      });
      xmlStream.leafNode("v:stroke", { joinstyle: "miter" });
      xmlStream.leafNode("v:path", { gradientshapeok: "t", "o:connecttype": "rect" });
      xmlStream.closeNode();
    }

    // Shapetype 201 for form control checkboxes
    if (hasFormControls) {
      xmlStream.openNode("v:shapetype", {
        id: "_x0000_t201",
        coordsize: "21600,21600",
        "o:spt": "201",
        path: "m,l,21600r21600,l21600,xe"
      });
      xmlStream.leafNode("v:stroke", { joinstyle: "miter" });
      xmlStream.leafNode("v:path", {
        shadowok: "f",
        "o:extrusionok": "f",
        strokeok: "f",
        fillok: "f",
        "o:connecttype": "rect"
      });
      xmlStream.leafNode("o:lock", { "v:ext": "edit", shapetype: "t" });
      xmlStream.closeNode();
    }

    // Render comment shapes
    if (hasComments) {
      for (let i = 0; i < comments.length; i++) {
        this.map["v:shape"].render(xmlStream, comments[i], i);
      }
    }

    // Render form control shapes
    if (hasFormControls) {
      for (const control of formControls) {
        this._renderCheckboxShape(xmlStream, control);
      }
    }

    xmlStream.closeNode();
  }

  /**
   * Render a checkbox form control shape
   */
  private _renderCheckboxShape(xmlStream: any, control: FormCheckboxModel): void {
    const shapeAttrs: Record<string, string> = {
      id: `_x0000_s${control.shapeId}`,
      type: "#_x0000_t201",
      style: FormCheckbox.getVmlStyle(control),
      "o:insetmode": "auto",
      fillcolor: "buttonFace [67]",
      strokecolor: "windowText [64]",
      "o:preferrelative": "t",
      filled: "f",
      stroked: "f"
    };

    xmlStream.openNode("v:shape", shapeAttrs);

    // Fill element
    xmlStream.leafNode("v:fill", { "o:detectmouseclick": "t" });

    // Lock element
    xmlStream.leafNode("o:lock", { "v:ext": "edit", text: "t" });

    // Textbox for label
    if (control.text) {
      xmlStream.openNode("v:textbox", {
        style: "mso-direction-alt:auto",
        "o:singleclick": "t"
      });
      xmlStream.openNode("div", { style: "text-align:left" });
      xmlStream.openNode("font", { face: "Tahoma", size: "160", color: "auto" });
      xmlStream.writeText(control.text);
      xmlStream.closeNode(); // font
      xmlStream.closeNode(); // div
      xmlStream.closeNode(); // v:textbox
    }

    // ClientData - the core of the checkbox control
    xmlStream.openNode("x:ClientData", { ObjectType: "Checkbox" });

    // Match Excel's VML patterns (similar to Note ClientData): include positioning and cell address.
    // Omitting these can cause Excel to repair the sheet by dropping all legacy controls.
    xmlStream.leafNode("x:MoveWithCells");
    xmlStream.leafNode("x:SizeWithCells");

    // Anchor position
    xmlStream.openNode("x:Anchor");
    xmlStream.writeText(FormCheckbox.getVmlAnchor(control));
    xmlStream.closeNode();

    // Protection / text locking
    xmlStream.leafNode("x:Locked", undefined, "False");
    xmlStream.leafNode("x:LockText", undefined, "True");

    // Print settings
    xmlStream.leafNode("x:PrintObject", undefined, control.print ? "True" : "False");
    xmlStream.leafNode("x:AutoFill", undefined, "False");
    xmlStream.leafNode("x:AutoLine", undefined, "False");
    xmlStream.leafNode("x:TextHAlign", undefined, "Left");
    xmlStream.leafNode("x:TextVAlign", undefined, "Center");

    // Linked cell
    if (control.link) {
      xmlStream.leafNode("x:FmlaLink", undefined, control.link);
    }

    // 3D appearance
    if (control.noThreeD) {
      xmlStream.leafNode("x:NoThreeD");
    }

    // Checked state (0 = unchecked, 1 = checked, 2 = mixed)
    xmlStream.leafNode("x:Checked", undefined, String(FormCheckbox.getVmlCheckedValue(control)));

    // Cell address (0-based row/column)
    xmlStream.leafNode("x:Row", undefined, String(control.tl.row));
    xmlStream.leafNode("x:Column", undefined, String(control.tl.col));

    xmlStream.closeNode(); // x:ClientData
    xmlStream.closeNode(); // v:shape
  }

  // Parsing - delegate to VmlShapeXform for notes
  parseOpen(node: any): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {
          comments: [],
          formControls: []
        };
        break;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model!.comments!.push(this.parser.model);
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        return false;
      default:
        return true;
    }
  }

  static DRAWING_ATTRIBUTES = {
    "xmlns:v": "urn:schemas-microsoft-com:vml",
    "xmlns:o": "urn:schemas-microsoft-com:office:office",
    "xmlns:x": "urn:schemas-microsoft-com:office:excel"
  };
}

export { VmlDrawingXform };
export type { VmlDrawingModel };
