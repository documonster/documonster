import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

class VmlProtectionXform extends BaseXform {
  declare private _model: { tag?: string };
  declare private text: string;

  constructor(model?: { tag?: string }) {
    super();
    this._model = model || {};
    this.text = "";
  }

  get tag(): string {
    return this._model?.tag ?? "";
  }

  render(xmlStream: XmlSink, model: any): void {
    xmlStream.leafNode(this.tag, undefined, model);
  }

  parseOpen(node: ParseOpenTag): boolean {
    switch (node.name) {
      case this.tag:
        this.text = "";
        return true;
      default:
        return false;
    }
  }

  parseText(text: string): void {
    this.text = text;
  }

  parseClose(): boolean {
    return false;
  }
}

export { VmlProtectionXform };
