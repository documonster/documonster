import type { Color } from "@excel/types";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

// Color encapsulates translation from color model to/from xlsx
class ColorXform extends BaseXform<Partial<Color>> {
  declare private name: string;

  constructor(name?: string) {
    super();

    // this.name controls the xm node name
    this.name = name ?? "color";
  }

  get tag(): string {
    return this.name;
  }

  render(xmlStream: XmlSink, model?: Partial<Color>): boolean {
    if (model) {
      xmlStream.openNode(this.name);
      if (model.argb) {
        // `rgb` is `ST_UnsignedIntHex`, four bytes: a six-digit colour gives no alpha, and an opaque one is the only
        // sensible reading — the reading the XLSB writer has always made (`xlsb/color.ts`), so both formats agree.
        xmlStream.addAttribute("rgb", model.argb.length === 6 ? `FF${model.argb}` : model.argb);
      } else if (model.theme !== undefined) {
        xmlStream.addAttribute("theme", model.theme);
        if (model.tint !== undefined) {
          xmlStream.addAttribute("tint", model.tint);
        }
      } else if (model.indexed !== undefined) {
        xmlStream.addAttribute("indexed", model.indexed);
      } else {
        xmlStream.addAttribute("auto", "1");
      }
      xmlStream.closeNode();
      return true;
    }
    return false;
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (node.name === this.name) {
      if (node.attributes.rgb) {
        this.model = { argb: node.attributes.rgb };
      } else if (node.attributes.theme) {
        this.model = { theme: parseInt(node.attributes.theme, 10) };
        if (node.attributes.tint) {
          this.model.tint = parseFloat(node.attributes.tint);
        }
      } else if (node.attributes.indexed) {
        this.model = { indexed: parseInt(node.attributes.indexed, 10) };
      } else {
        this.model = undefined;
      }
      return true;
    }
    return false;
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }
}

export { ColorXform };
