import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

class CfvoXform extends BaseXform {
  get tag(): string {
    return "cfvo";
  }

  render(xmlStream: XmlSink, model: any): void {
    xmlStream.leafNode(this.tag, {
      type: model.type,
      val: model.value
    });
  }

  parseOpen(node: ParseOpenTag): void {
    this.model = {
      type: node.attributes.type,
      value: BaseXform.toFloatValue(node.attributes.val)
    };
  }

  parseClose(name: string): boolean {
    return name !== this.tag;
  }
}

export { CfvoXform };
