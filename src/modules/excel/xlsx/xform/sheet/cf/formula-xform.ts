import { fromFileFormula, toFileFormula } from "@excel/utils/future-functions";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { XmlSink } from "@xml/types";

class FormulaXform extends BaseXform<string> {
  get tag(): string {
    return "formula";
  }

  render(xmlStream: XmlSink, model?: string): void {
    xmlStream.leafNode(this.tag, undefined, model === undefined ? undefined : toFileFormula(model));
  }

  parseOpen(): void {
    this.model = "";
  }

  parseText(text: string): void {
    this.model += text;
  }

  parseClose(name: string): boolean {
    if (name === this.tag && typeof this.model === "string") {
      this.model = fromFileFormula(this.model);
    }
    return name !== this.tag;
  }
}

export { FormulaXform };
