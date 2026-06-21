import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { XmlSink } from "@xml/types";

interface CalcPropertiesModel {
  fullCalcOnLoad?: boolean;
  iterate?: boolean;
  iterateCount?: number;
  iterateDelta?: number;
}

class WorkbookCalcPropertiesXform extends BaseXform {
  render(xmlStream: XmlSink, model: CalcPropertiesModel): void {
    xmlStream.leafNode("calcPr", {
      calcId: 171027,
      fullCalcOnLoad: model.fullCalcOnLoad ? 1 : undefined,
      iterate: model.iterate ? 1 : undefined,
      iterateCount: model.iterateCount !== undefined ? model.iterateCount : undefined,
      iterateDelta: model.iterateDelta !== undefined ? model.iterateDelta : undefined
    });
  }

  parseOpen(node: any): boolean {
    if (node.name === "calcPr") {
      const attrs = node.attributes ?? {};
      this.model = {
        fullCalcOnLoad: attrs.fullCalcOnLoad === "1",
        iterate: attrs.iterate === "1" ? true : undefined,
        iterateCount:
          attrs.iterateCount !== undefined ? parseInt(attrs.iterateCount, 10) : undefined,
        iterateDelta: attrs.iterateDelta !== undefined ? parseFloat(attrs.iterateDelta) : undefined
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

export { WorkbookCalcPropertiesXform };
