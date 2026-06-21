import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

class VmlPositionXform extends BaseXform<{ [key: string]: boolean }> {
  declare private _model: { tag?: string };

  constructor(model?: { tag?: string }) {
    super();
    this._model = model || {};
    this.model = {};
  }

  get tag(): string {
    return this._model?.tag ?? "";
  }

  render(xmlStream: XmlSink, model: string, type?: string[]): void {
    if (type && model === type[2]) {
      xmlStream.leafNode(this.tag);
    } else if (type && this.tag === "x:SizeWithCells" && model === type[1]) {
      xmlStream.leafNode(this.tag);
    }
  }

  parseOpen(node: ParseOpenTag): boolean {
    switch (node.name) {
      case this.tag:
        this.model = {};
        this.model[this.tag] = true;
        return true;
      default:
        return false;
    }
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }
}

export { VmlPositionXform };
