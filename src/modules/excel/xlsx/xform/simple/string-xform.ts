import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlAttributes, XmlSink } from "@xml/types";

interface StringXformOptions {
  tag: string;
  attr?: string;
  attrs?: XmlAttributes;
}

class StringXform extends BaseXform {
  declare private tag: string;
  declare private attr?: string;
  declare private attrs?: XmlAttributes;
  declare private text: string[];

  constructor(options: StringXformOptions) {
    super();

    this.tag = options.tag;
    this.attr = options.attr;
    this.attrs = options.attrs;
    this.text = [];
  }

  render(xmlStream: XmlSink, model?: string): void {
    if (model !== undefined) {
      xmlStream.openNode(this.tag);
      if (this.attrs) {
        xmlStream.addAttributes(this.attrs);
      }
      if (this.attr) {
        xmlStream.addAttribute(this.attr, model);
      } else {
        xmlStream.writeText(model);
      }
      xmlStream.closeNode();
    }
  }

  parseOpen(node: ParseOpenTag): void {
    if (node.name === this.tag) {
      if (this.attr) {
        this.model = node.attributes[this.attr];
      } else {
        this.text = [];
      }
    }
  }

  parseText(text: string): void {
    if (!this.attr) {
      this.text.push(text);
    }
  }

  parseClose(): boolean {
    if (!this.attr) {
      this.model = this.text.join("");
    }
    return false;
  }
}

export { StringXform };
