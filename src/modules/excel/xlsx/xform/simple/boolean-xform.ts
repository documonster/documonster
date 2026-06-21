import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { XmlSink } from "@xml/types";

interface BooleanXformOptions {
  tag: string;
  attr: string;
}

class BooleanXform extends BaseXform {
  declare private tag: string;
  declare private attr: string;

  constructor(options: BooleanXformOptions) {
    super();

    this.tag = options.tag;
    this.attr = options.attr;
  }

  render(xmlStream: XmlSink, model?: boolean): void {
    if (model) {
      xmlStream.openNode(this.tag);
      xmlStream.closeNode();
    }
  }

  parseOpen(node: any): void {
    if (node.name === this.tag) {
      this.model = true;
    }
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }
}

export { BooleanXform };
