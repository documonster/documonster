import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

interface CustomFilterModel {
  val: string;
  operator?: string;
}

class CustomFilterXform extends BaseXform<CustomFilterModel> {
  constructor() {
    super();
    this.model = { val: "" };
  }

  get tag(): string {
    return "customFilter";
  }

  render(xmlStream: XmlSink, model: CustomFilterModel): void {
    xmlStream.leafNode(this.tag, {
      val: model.val,
      operator: model.operator
    });
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (node.name === this.tag) {
      this.model = {
        val: node.attributes.val,
        operator: node.attributes.operator
      };
      return true;
    }
    return false;
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }
}

export { CustomFilterXform };
