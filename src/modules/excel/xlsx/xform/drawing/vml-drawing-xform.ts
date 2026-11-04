import type { FormCheckboxModel, LegacyControlModel } from "@excel/core/form-control";
import {
  formCheckboxVmlAnchor,
  formCheckboxVmlCheckedValue,
  formCheckboxVmlStyle
} from "@excel/core/form-control";
import { formCheckboxPlacement } from "@excel/core/form-control-placement";
import { allocateVmlShapeIds, createVmlShapeIdSpace, vmlIdMap } from "@excel/utils/vml-shape-ids";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ShapeModel } from "@excel/xlsx/xform/comment/vml-shape-xform";
import { VmlShapeXform } from "@excel/xlsx/xform/comment/vml-shape-xform";
import type { EchoNode } from "@excel/xlsx/xform/drawing/generic-echo-xform";
import { GenericEchoXform } from "@excel/xlsx/xform/drawing/generic-echo-xform";
import {
  VmlFormControlCapture,
  legacyControlFromVml
} from "@excel/xlsx/xform/drawing/vml-form-control-capture";
import { xmlValue } from "@xml/encode";
import type { ParseOpenTag, XmlSink } from "@xml/types";
import { StdDocAttributes } from "@xml/writer";

/**
 * Unified VML Drawing Xform - Combines Notes (comments) and Form Controls
 *
 * Excel uses a single VML file per worksheet that can contain:
 * - Comment/note shapes (shapetype 202)
 * - Form control shapes (checkbox shapetype 201, etc.)
 *
 * This unified xform renders both into a single VML file.
 */

/** Header/footer image model for VML rendering. */
interface VmlHeaderImageModel {
  /** rId referencing the image in the VML drawing's .rels */
  imageRelId: string;
  /** Image width in points */
  width?: number;
  /** Image height in points */
  height?: number;
  /** Excel header/footer section id: LH, CH, RH, LF, CF, or RF. */
  position?: string;
  /**
   * The `<v:shape>` as read, so a rewrite keeps what Excel's Format Picture dialog put on it — the
   * picture's title, washout (`gain`/`blacklevel`), grayscale, bilevel and crop — and anything else the
   * shape carries. Its id, relationship and size are still this model's.
   */
  shape?: EchoNode;
}

interface VmlDrawingModel {
  /** Comment/note shapes */
  comments?: ShapeModel[];
  /** Form control checkboxes */
  formControls?: FormCheckboxModel[];
  /**
   * Every other legacy control, kept as read. On write each carries the shape id it is to be drawn
   * with, already allocated.
   */
  legacyControls?: LegacyControlModel[];
  /** Header/footer image (for watermark in header mode) */
  headerImage?: VmlHeaderImageModel;
  /** All header/footer image shapes parsed from a foreign workbook. */
  headerImages?: VmlHeaderImageModel[];
  /**
   * Shape ids for the notes and header/footer images, in order, from the write's workbook-wide
   * allocation (`@excel/utils/vml-shape-ids`). The form controls carry their own. Absent, the part
   * is numbered as if it had the workbook to itself.
   */
  shapeIds?: readonly number[];
}

class VmlDrawingXform extends BaseXform<VmlDrawingModel> {
  declare public map: Record<string, BaseXform>;
  declare public parser?: BaseXform;

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
  render(xmlStream: XmlSink, model?: VmlDrawingModel): void {
    const renderModel = (model || this.model)!;
    const comments = renderModel.comments;
    const formControls = renderModel.formControls;
    const declaredHeaderImages =
      renderModel.headerImages ?? (renderModel.headerImage ? [renderModel.headerImage] : []);
    // Excel holds at most one image per header/footer section, and a VML part
    // must not repeat a shape id. Keep the last declaration per section.
    const headerImages = [
      ...new Map(
        declaredHeaderImages.map(image => [image.position ?? "CH", image] as const)
      ).values()
    ];
    const hasComments = comments && comments.length > 0;
    const hasFormControls = formControls && formControls.length > 0;
    const hasHeaderImage = headerImages.length > 0;
    const legacyControls = renderModel.legacyControls ?? [];

    const noteCount = hasComments ? comments.length : 0;
    const controlIds = [
      ...(hasFormControls ? formControls.map(control => control.shapeId) : []),
      ...legacyControls.map(control => control.shapeId)
    ];
    const shapeIds =
      renderModel.shapeIds ??
      allocateVmlShapeIds(createVmlShapeIdSpace(), {
        fixed: controlIds,
        count: noteCount + headerImages.length
      }).free;
    const noteIds = shapeIds.slice(0, noteCount);
    const headerImageIds = shapeIds.slice(noteCount);

    xmlStream.openXml(StdDocAttributes);
    // A preserved shape may use a prefix the source root declared and this one does not.
    const rootAttributes: Record<string, string> = { ...VmlDrawingXform.DRAWING_ATTRIBUTES };
    for (const control of legacyControls) {
      for (const [name, uri] of Object.entries(control.vmlNamespaces ?? {})) {
        rootAttributes[name] ??= uri;
      }
    }
    xmlStream.openNode(this.tag, rootAttributes);

    // Shape layout - shared by both notes and form controls
    xmlStream.openNode("o:shapelayout", { "v:ext": "edit" });
    xmlStream.leafNode("o:idmap", {
      "v:ext": "edit",
      data: vmlIdMap([...noteIds, ...controlIds, ...headerImageIds])
    });
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

    // Shapetype 75 for header/footer image (watermark)
    if (hasHeaderImage) {
      xmlStream.openNode("v:shapetype", {
        id: "_x0000_t75",
        coordsize: "21600,21600",
        "o:spt": "75",
        "o:preferrelative": "t",
        path: "m@4@5l@4@11@9@11@9@5xe",
        filled: "f",
        stroked: "f"
      });
      xmlStream.leafNode("v:stroke", { joinstyle: "miter" });
      xmlStream.openNode("v:formulas");
      xmlStream.leafNode("v:f", { eqn: "if lineDrawn pixelLineWidth 0" });
      xmlStream.leafNode("v:f", { eqn: "sum @0 1 0" });
      xmlStream.leafNode("v:f", { eqn: "sum 0 0 @1" });
      xmlStream.leafNode("v:f", { eqn: "prod @2 1 2" });
      xmlStream.leafNode("v:f", { eqn: "prod @3 21600 pixelWidth" });
      xmlStream.leafNode("v:f", { eqn: "prod @3 21600 pixelHeight" });
      xmlStream.leafNode("v:f", { eqn: "sum @0 0 1" });
      xmlStream.leafNode("v:f", { eqn: "prod @6 1 2" });
      xmlStream.leafNode("v:f", { eqn: "prod @7 21600 pixelWidth" });
      xmlStream.leafNode("v:f", { eqn: "sum @8 21600 0" });
      xmlStream.leafNode("v:f", { eqn: "prod @7 21600 pixelHeight" });
      xmlStream.leafNode("v:f", { eqn: "sum @10 21600 0" });
      xmlStream.closeNode(); // v:formulas
      xmlStream.leafNode("v:path", {
        "o:extrusionok": "f",
        gradientshapeok: "t",
        "o:connecttype": "rect"
      });
      xmlStream.leafNode("o:lock", { "v:ext": "edit", aspectratio: "t" });
      xmlStream.closeNode(); // v:shapetype
    }

    // The shapetypes the preserved controls name, unless one above already declared that id.
    const declaredTypes = new Set<string>([
      ...(hasComments ? ["_x0000_t202"] : []),
      ...(hasFormControls ? ["_x0000_t201"] : []),
      ...(hasHeaderImage ? ["_x0000_t75"] : [])
    ]);
    const echo = new GenericEchoXform("v:shapetype");
    for (const control of legacyControls) {
      const shapeType = control.vmlShapeType;
      const id =
        shapeType?.attrs.id === undefined ? undefined : xmlValue(shapeType.attrs.id, "undefined");
      if (shapeType && id !== undefined && !declaredTypes.has(id)) {
        declaredTypes.add(id);
        echo.render(xmlStream, shapeType);
      }
    }

    // Render comment shapes
    if (hasComments) {
      for (let i = 0; i < comments.length; i++) {
        (this.map["v:shape"] as VmlShapeXform).render(xmlStream, comments[i], noteIds[i]);
      }
    }

    // Render form control shapes
    if (hasFormControls) {
      for (const control of formControls) {
        this._renderCheckboxShape(xmlStream, control);
      }
    }

    for (const control of legacyControls) {
      echo.render(xmlStream, withVmlShapeId(control.vmlShape, control.shapeId));
    }

    // Render header/footer image shape
    if (hasHeaderImage) {
      for (let i = 0; i < headerImages.length; i++) {
        this._renderHeaderImageShape(xmlStream, headerImages[i], headerImageIds[i]);
      }
    }

    xmlStream.closeNode();
  }

  /**
   * Render a header/footer image shape for watermark
   */
  private _renderHeaderImageShape(
    xmlStream: XmlSink,
    headerImage: VmlHeaderImageModel,
    shapeId: number
  ): void {
    const width = headerImage.width ?? 467.25;
    const height = headerImage.height ?? 311.25;

    const position = headerImage.position ?? "CH";
    if (headerImage.shape) {
      new GenericEchoXform("v:shape").render(
        xmlStream,
        redrawHeaderShape(headerImage.shape, {
          id: position,
          spid: `_x0000_s${shapeId}`,
          relId: headerImage.imageRelId,
          width,
          height
        })
      );
      return;
    }
    xmlStream.openNode("v:shape", {
      id: position,
      "o:spid": `_x0000_s${shapeId}`,
      type: "#_x0000_t75",
      style: `position:absolute;margin-left:0;margin-top:0;width:${width}pt;height:${height}pt;z-index:1`
    });
    xmlStream.leafNode("v:imagedata", {
      "o:relid": headerImage.imageRelId,
      "o:title": "watermark"
    });
    xmlStream.leafNode("o:lock", { "v:ext": "edit", rotation: "t" });
    xmlStream.closeNode(); // v:shape
  }

  /**
   * Render a checkbox form control shape
   */
  private _renderCheckboxShape(xmlStream: XmlSink, control: FormCheckboxModel): void {
    const shapeAttrs: Record<string, string> = {
      id: `_x0000_s${control.shapeId}`,
      type: "#_x0000_t201",
      style: control.vmlStyle ?? formCheckboxVmlStyle(control),
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

    // Positioning, as Excel writes it: each element is present when the control does *not* move, or does
    // *not* size, with the cells under it — so the default "move but don't size" carries `x:SizeWithCells` only.
    const placement = formCheckboxPlacement(control);
    if (placement === "absolute") {
      xmlStream.leafNode("x:MoveWithCells");
    }
    if (placement !== "twoCell") {
      xmlStream.leafNode("x:SizeWithCells");
    }

    // Anchor position
    xmlStream.openNode("x:Anchor");
    xmlStream.writeText(formCheckboxVmlAnchor(control));
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
    xmlStream.leafNode("x:Checked", undefined, String(formCheckboxVmlCheckedValue(control)));

    // Cell address (0-based row/column)
    xmlStream.leafNode("x:Row", undefined, String(control.tl.row));
    xmlStream.leafNode("x:Column", undefined, String(control.tl.col));

    xmlStream.closeNode(); // x:ClientData
    xmlStream.closeNode(); // v:shape
  }

  // Parsing - delegate to VmlShapeXform for notes, handle header images directly
  parseOpen(node: ParseOpenTag): boolean {
    if (this._shapeType) {
      this._shapeType.parseOpen(node);
      return true;
    }
    if (this.parser) {
      this._control?.open(node);
      this._shape?.parseOpen(node);
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
        this._shapeTypes = new Map();
        this._namespaces = Object.fromEntries(
          Object.entries(node.attributes).filter(
            ([name]) => name.startsWith("xmlns:") && !(name in VmlDrawingXform.DRAWING_ATTRIBUTES)
          )
        );
        break;
      case "v:shapetype":
        this._shapeType = new GenericEchoXform("v:shapetype");
        this._shapeType.parseOpen(node);
        break;
      case "v:shape":
        // Check if this is a header image shape (type="#_x0000_t75")
        if (node.attributes.type === "#_x0000_t75") {
          this._parsingHeaderImage = true;
          this._headerShape = new GenericEchoXform("v:shape");
          this._headerShape.parseOpen(node);
          this._headerImagePosition = node.attributes.id;
          // Extract width/height from style
          const style = node.attributes.style || "";
          const widthMatch = /width:([0-9.]+)pt/.exec(style);
          const heightMatch = /height:([0-9.]+)pt/.exec(style);
          this._headerImageWidth = widthMatch ? parseFloat(widthMatch[1]) : undefined;
          this._headerImageHeight = heightMatch ? parseFloat(heightMatch[1]) : undefined;
        } else {
          // Regular shape — delegate to VmlShapeXform (comments). A form control is a `v:shape` too, and
          // which one it is only shows at `<x:ClientData ObjectType>`, so the control is captured alongside
          // and the two are told apart when the shape closes.
          this._control = new VmlFormControlCapture(node.attributes);
          this._shape = new GenericEchoXform("v:shape");
          this._shape.parseOpen(node);
          this.parser = this.map[node.name];
          if (this.parser) {
            this.parser.parseOpen(node);
          }
        }
        break;
      default:
        this._headerShape?.parseOpen(node);
        if (this._parsingHeaderImage && node.name === "v:imagedata") {
          this._headerImageRelId = node.attributes["o:relid"];
        } else {
          this.parser = this.map[node.name];
          if (this.parser) {
            this.parser.parseOpen(node);
          }
        }
        break;
    }
    return true;
  }

  parseText(text: string): void {
    if (this._shapeType) {
      this._shapeType.parseText(text);
      return;
    }
    this._control?.text(text);
    this._shape?.parseText(text);
    this._headerShape?.parseText(text);
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this._shapeType) {
      if (!this._shapeType.parseClose(name)) {
        const shapeType = this._shapeType.model;
        this._shapeType = undefined;
        if (shapeType?.attrs.id !== undefined) {
          this._shapeTypes.set(String(shapeType.attrs.id), shapeType);
        }
      }
      return true;
    }
    if (this.parser) {
      this._control?.close(name);
      this._shape?.parseClose(name);
      if (!this.parser.parseClose(name)) {
        const control = this._control;
        const shape = this._shape?.model;
        this._control = undefined;
        this._shape = undefined;
        const checkbox = control?.objectType === "Checkbox" ? control.toCheckbox() : undefined;
        if (checkbox) {
          this.model!.formControls!.push(checkbox);
        } else if (control !== undefined && !control.isNote) {
          // Any other control — and a checkbox too damaged to model — is kept as it was read.
          const legacy = shape && legacyControlFromVml(control, shape);
          if (legacy) {
            (this.model!.legacyControls ??= []).push(legacy);
          }
        } else if (control === undefined || control.isNote) {
          // Only a note is a comment's box. Another control (a button, a list box) is not modelled, and
          // handing it to the comments merged it into whichever note shared its index.
          this.model!.comments!.push(this.parser.model);
        }
        this.parser = undefined;
      }
      return true;
    }
    let headerShape: EchoNode | undefined;
    if (this._headerShape) {
      if (this._headerShape.parseClose(name)) {
        return true;
      }
      headerShape = this._headerShape.model;
      this._headerShape = undefined;
    }
    switch (name) {
      case "v:shape":
        if (this._parsingHeaderImage && this._headerImageRelId) {
          const headerImage: VmlHeaderImageModel = {
            imageRelId: this._headerImageRelId,
            width: this._headerImageWidth,
            height: this._headerImageHeight,
            position: this._headerImagePosition,
            ...(headerShape ? { shape: headerShape } : {})
          };
          this.model!.headerImages ??= [];
          this.model!.headerImages.push(headerImage);
          this.model!.headerImage ??= headerImage;
        }
        this._parsingHeaderImage = false;
        this._headerImageRelId = undefined;
        this._headerImageWidth = undefined;
        this._headerImageHeight = undefined;
        this._headerImagePosition = undefined;
        return true;
      case this.tag:
        for (const control of this.model!.legacyControls ?? []) {
          const shapeType = this._shapeTypes.get(
            String(control.vmlShape.attrs.type ?? "").replace(/^#/, "")
          );
          if (shapeType) {
            control.vmlShapeType = shapeType;
          }
          if (Object.keys(this._namespaces).length > 0) {
            control.vmlNamespaces = this._namespaces;
          }
        }
        return false;
      default:
        return true;
    }
  }

  /** The form control the shape being parsed may turn out to be. */
  private _control?: VmlFormControlCapture;
  /** The shape being parsed, captured whole in case it is a control this library does not model. */
  private _shape?: GenericEchoXform;
  /** The header/footer picture's shape being captured. */
  private _headerShape?: GenericEchoXform;
  /** The `<v:shapetype>` being captured. */
  private _shapeType?: GenericEchoXform;
  /** The part's shapetypes by id, for the preserved controls that name one. */
  private _shapeTypes = new Map<string, EchoNode>();
  /** The root's namespace declarations beyond the three this writer always declares. */
  private _namespaces: Record<string, string> = {};

  // Internal state for parsing header image shapes
  private _parsingHeaderImage = false;
  private _headerImageRelId?: string;
  private _headerImageWidth?: number;
  private _headerImageHeight?: number;
  private _headerImagePosition?: string;

  static DRAWING_ATTRIBUTES = {
    "xmlns:v": "urn:schemas-microsoft-com:vml",
    "xmlns:o": "urn:schemas-microsoft-com:office:office",
    "xmlns:x": "urn:schemas-microsoft-com:office:excel"
  };
}

/**
 * A header/footer picture's shape as read, given the section, shape id, relationship and size this
 * write draws it with. Everything else on the shape and its `<v:imagedata>` is kept.
 */
function redrawHeaderShape(
  shape: EchoNode,
  place: { id: string; spid: string; relId: string; width: number; height: number }
): EchoNode {
  const style = String(
    shape.attrs.style ?? "position:absolute;margin-left:0;margin-top:0;z-index:1"
  );
  const sized = [`width:${place.width}pt`, `height:${place.height}pt`].reduce(
    (current, declaration) => {
      const property = declaration.slice(0, declaration.indexOf(":"));
      const pattern = new RegExp(`(^|;)\\s*${property}\\s*:[^;]*`);
      return pattern.test(current)
        ? current.replace(pattern, `$1${declaration}`)
        : `${current.replace(/;?\s*$/, "")};${declaration}`;
    },
    style
  );
  const redraw = (node: EchoNode): EchoNode =>
    node.tag === "v:imagedata"
      ? { ...node, attrs: { ...node.attrs, "o:relid": place.relId } }
      : node;
  const content = (shape.content ?? shape.children).map(item =>
    typeof item === "object" ? redraw(item) : item
  );
  return {
    ...shape,
    attrs: { ...shape.attrs, id: place.id, "o:spid": place.spid, style: sized },
    children: shape.children.map(redraw),
    content
  };
}

/** `shape` drawn as `_x0000_s{shapeId}`, in whichever attribute carried its id. */
function withVmlShapeId(shape: EchoNode, shapeId: number): EchoNode {
  const spid = `_x0000_s${shapeId}`;
  const attrs = { ...shape.attrs };
  if (attrs["o:spid"] === undefined && /^_x0000_s\d+$/.test(String(attrs.id ?? ""))) {
    attrs.id = spid;
  } else {
    attrs["o:spid"] = spid;
  }
  return { ...shape, attrs };
}

export { VmlDrawingXform };
export type { VmlDrawingModel };
