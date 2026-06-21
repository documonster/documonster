import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { XmlSink } from "@xml/types";

class FormulaXform extends BaseXform {
  get tag(): string {
    return "formula";
  }

  render(xmlStream: XmlSink, model: any): void {
    xmlStream.leafNode(this.tag, undefined, model);
  }

  parseOpen(): void {
    this.model = "";
  }

  parseText(text: string): void {
    this.model += text;
  }

  parseClose(name: string): boolean {
    return name !== this.tag;
  }
}

export { FormulaXform };
