import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { VmlClientDataXform } from "@excel/xlsx/xform/comment/vml-client-data-xform";
import { VmlTextboxXform } from "@excel/xlsx/xform/comment/vml-textbox-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

interface ShapeModel {
  note: {
    margins?: {
      insetmode?: string;
      inset?: number[];
    };
    width?: number;
    height?: number;
  };
  refAddress?: any;
}

/** Default comment box geometry in points (matches legacy Excel notes). */
const DEFAULT_NOTE_WIDTH_PT = 97.8;
const DEFAULT_NOTE_HEIGHT_PT = 59.1;

class VmlShapeXform extends BaseXform {
  declare public map: { [key: string]: any };
  declare public parser: any;
  declare public model: any;

  constructor() {
    super();
    this.map = {
      "v:textbox": new VmlTextboxXform(),
      "x:ClientData": new VmlClientDataXform()
    };
  }

  get tag(): string {
    return "v:shape";
  }

  render(xmlStream: XmlSink, model: ShapeModel, index?: number): void {
    xmlStream.openNode("v:shape", VmlShapeXform.V_SHAPE_ATTRIBUTES(model, index ?? 0));

    xmlStream.leafNode("v:fill", { color2: "infoBackground [80]" });
    xmlStream.leafNode("v:shadow", { color: "none [81]", obscured: "t" });
    xmlStream.leafNode("v:path", { "o:connecttype": "none" });
    this.map["v:textbox"].render(xmlStream, model);
    this.map["x:ClientData"].render(xmlStream, model);

    xmlStream.closeNode();
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {
          margins: {
            insetmode: node.attributes["o:insetmode"]
          },
          anchor: "",
          editAs: "",
          protection: {}
        };
        {
          // Recover the comment box geometry from the VML style string
          // (e.g. "...width:120pt;height:80pt;..."). Only surface width/height
          // when they differ from the legacy defaults, so untouched notes keep
          // a clean model (and stay byte-compatible with prior behaviour).
          const style: string = node.attributes.style ?? "";
          const width = VmlShapeXform.parseStyleLength(style, "width");
          const height = VmlShapeXform.parseStyleLength(style, "height");
          if (width !== undefined && width !== DEFAULT_NOTE_WIDTH_PT) {
            this.model.width = width;
          }
          if (height !== undefined && height !== DEFAULT_NOTE_HEIGHT_PT) {
            this.model.height = height;
          }
        }
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
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        this.model.margins.inset = this.map["v:textbox"].model && this.map["v:textbox"].model.inset;
        this.model.protection =
          this.map["x:ClientData"].model && this.map["x:ClientData"].model.protection;
        this.model.anchor = this.map["x:ClientData"].model && this.map["x:ClientData"].model.anchor;
        this.model.editAs = this.map["x:ClientData"].model && this.map["x:ClientData"].model.editAs;
        if (this.map["x:ClientData"].model) {
          this.model.row = this.map["x:ClientData"].model.row;
          this.model.col = this.map["x:ClientData"].model.col;
        }
        return false;
      default:
        return true;
    }
  }

  /**
   * Extract a points-valued length (e.g. `width:120pt`) from a VML style
   * string. Returns `undefined` when the property is absent or not in `pt`.
   */
  static parseStyleLength(style: string, prop: "width" | "height"): number | undefined {
    const match = new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*([0-9.]+)pt`, "i").exec(style);
    if (!match) {
      return undefined;
    }
    const value = parseFloat(match[1]);
    return Number.isFinite(value) ? value : undefined;
  }

  static V_SHAPE_ATTRIBUTES(model: ShapeModel, index: number): any {
    const width = model.note?.width ?? DEFAULT_NOTE_WIDTH_PT;
    const height = model.note?.height ?? DEFAULT_NOTE_HEIGHT_PT;
    return {
      id: `_x0000_s${1025 + index}`,
      type: "#_x0000_t202",
      style: `position:absolute; margin-left:105.3pt;margin-top:10.5pt;width:${width}pt;height:${height}pt;z-index:1;visibility:hidden`,
      fillcolor: "infoBackground [80]",
      strokecolor: "none [81]",
      "o:insetmode": model.note.margins && model.note.margins.insetmode
    };
  }
}

export { VmlShapeXform };
