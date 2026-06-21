import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { XmlSink } from "@xml/types";

interface FilterModel {
  val: string;
}

class FilterXform extends BaseXform<FilterModel> {
  constructor() {
    super();
    this.model = { val: "" };
  }

  get tag(): string {
    return "filter";
  }

  render(xmlStream: XmlSink, model: FilterModel): void {
    xmlStream.leafNode(this.tag, {
      val: model.val
    });
  }

  parseOpen(node: any): boolean {
    if (node.name === this.tag) {
      this.model = {
        val: node.attributes.val
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

export { FilterXform };
