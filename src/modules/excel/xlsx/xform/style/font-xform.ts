import type { Color } from "@excel/types";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { BooleanXform } from "@excel/xlsx/xform/simple/boolean-xform";
import { IntegerXform } from "@excel/xlsx/xform/simple/integer-xform";
import { StringXform } from "@excel/xlsx/xform/simple/string-xform";
import { ColorXform } from "@excel/xlsx/xform/style/color-xform";
import { UnderlineXform } from "@excel/xlsx/xform/style/underline-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

interface FontModel {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | string;
  charset?: number;
  color?: Partial<Color>;
  condense?: boolean;
  extend?: boolean;
  family?: number;
  outline?: boolean;
  vertAlign?: string;
  scheme?: string;
  shadow?: boolean;
  strike?: boolean;
  size?: number;
  name?: string;
}

interface FontOptions {
  tagName: string;
  fontNameTag: string;
}

// Font encapsulates translation from font model to xlsx
class FontXform extends BaseXform {
  declare private options: FontOptions;
  declare public parser?: BaseXform;
  declare private renderOrder: string[];

  constructor(options?: FontOptions) {
    super();

    this.options = options || FontXform.OPTIONS;

    // Define properties in render order (Excel's expected order)
    const fontProperties = [
      { tag: "b", prop: "bold", xform: new BooleanXform({ tag: "b", attr: "val" }) },
      { tag: "i", prop: "italic", xform: new BooleanXform({ tag: "i", attr: "val" }) },
      { tag: "u", prop: "underline", xform: new UnderlineXform() },
      { tag: "strike", prop: "strike", xform: new BooleanXform({ tag: "strike", attr: "val" }) },
      {
        tag: "condense",
        prop: "condense",
        xform: new BooleanXform({ tag: "condense", attr: "val" })
      },
      { tag: "extend", prop: "extend", xform: new BooleanXform({ tag: "extend", attr: "val" }) },
      { tag: "outline", prop: "outline", xform: new BooleanXform({ tag: "outline", attr: "val" }) },
      { tag: "shadow", prop: "shadow", xform: new BooleanXform({ tag: "shadow", attr: "val" }) },
      { tag: "sz", prop: "size", xform: new IntegerXform({ tag: "sz", attr: "val" }) },
      { tag: "color", prop: "color", xform: new ColorXform() },
      {
        tag: this.options.fontNameTag,
        prop: "name",
        xform: new StringXform({ tag: this.options.fontNameTag, attr: "val" })
      },
      { tag: "family", prop: "family", xform: new IntegerXform({ tag: "family", attr: "val" }) },
      { tag: "scheme", prop: "scheme", xform: new StringXform({ tag: "scheme", attr: "val" }) },
      { tag: "charset", prop: "charset", xform: new IntegerXform({ tag: "charset", attr: "val" }) },
      {
        tag: "vertAlign",
        prop: "vertAlign",
        xform: new StringXform({ tag: "vertAlign", attr: "val" })
      }
    ];

    // Build map and renderOrder from single source of truth
    this.map = Object.fromEntries(
      fontProperties.map(p => [p.tag, { prop: p.prop, xform: p.xform }])
    );
    this.renderOrder = fontProperties.map(p => p.tag);
  }

  get tag(): string {
    return this.options.tagName;
  }

  render(xmlStream: XmlSink, model: FontModel): void {
    const { map, renderOrder } = this;

    xmlStream.openNode(this.options.tagName);
    renderOrder.forEach(tag => {
      map![tag].xform.render(xmlStream, model[map![tag].prop as keyof FontModel]);
    });
    xmlStream.closeNode();
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    if (this.map![node.name]) {
      this.parser = this.map![node.name].xform;
      // Child xform parseOpen returns a boolean (BaseXform's base signature is
      // void); the child here always reports whether it consumed the node.
      return (this.parser as unknown as { parseOpen(n: ParseOpenTag): boolean }).parseOpen(node);
    }
    switch (node.name) {
      case this.options.tagName:
        this.model = {};
        return true;
      default:
        return false;
    }
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this.parser && !this.parser.parseClose(name)) {
      const item = this.map![name];
      if (this.parser.model) {
        this.model[item.prop] = this.parser.model;
      }
      this.parser = undefined;
      return true;
    }
    switch (name) {
      case this.options.tagName:
        return false;
      default:
        return true;
    }
  }

  static OPTIONS: FontOptions = {
    tagName: "font",
    fontNameTag: "name"
  };
}

export { FontXform };
