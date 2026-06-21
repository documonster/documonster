import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { EMU_PER_PX } from "@utils/units";
import type { ParseOpenTag, XmlSink } from "@xml/types";

interface ExtModel {
  width: number;
  height: number;
}

class ExtXform extends BaseXform<ExtModel> {
  declare private tag: string;
  declare public map: Record<string, BaseXform>;

  constructor(options: { tag: string }) {
    super();

    this.tag = options.tag;
    this.map = {};
    this.model = { width: 0, height: 0 };
  }

  render(xmlStream: XmlSink, model: ExtModel): void {
    xmlStream.openNode(this.tag);

    const width = Math.floor(model.width * EMU_PER_PX);
    const height = Math.floor(model.height * EMU_PER_PX);

    xmlStream.addAttribute("cx", width);
    xmlStream.addAttribute("cy", height);

    xmlStream.closeNode();
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (node.name === this.tag) {
      this.model = {
        width: parseInt(node.attributes.cx ?? "0", 10) / EMU_PER_PX,
        height: parseInt(node.attributes.cy ?? "0", 10) / EMU_PER_PX
      };
      return true;
    }
    return false;
  }

  parseText(_text?: string): void {}

  parseClose(_name?: string): boolean {
    return false;
  }
}

export { ExtXform };
