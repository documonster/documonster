import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

interface SheetFormatPropertiesModel {
  defaultRowHeight: number;
  dyDescent?: number;
  outlineLevelRow: number;
  outlineLevelCol: number;
  defaultColWidth?: number;
  customHeight?: boolean;
}

class SheetFormatPropertiesXform extends BaseXform {
  get tag(): string {
    return "sheetFormatPr";
  }

  render(xmlStream: XmlSink, model?: SheetFormatPropertiesModel): void {
    if (model) {
      const attributes: any = {
        defaultRowHeight: model.defaultRowHeight,
        // Only output outlineLevelRow/Col when non-zero (matches Excel behavior)
        outlineLevelRow: model.outlineLevelRow || undefined,
        outlineLevelCol: model.outlineLevelCol || undefined,
        // Only output dyDescent if explicitly set (MS extension, not ECMA-376 standard)
        "x14ac:dyDescent":
          model.dyDescent !== undefined && model.dyDescent !== 0 ? model.dyDescent : undefined
      };
      // Only output defaultColWidth if explicitly set
      if (model.defaultColWidth) {
        attributes.defaultColWidth = model.defaultColWidth;
      }

      // Only output customHeight if it was present in the original file
      if (model.customHeight) {
        attributes.customHeight = "1";
      }

      if (Object.values(attributes).some((value: any) => value !== undefined)) {
        xmlStream.leafNode("sheetFormatPr", attributes);
      }
    }
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (node.name === "sheetFormatPr") {
      this.model = {
        defaultRowHeight: parseFloat(node.attributes.defaultRowHeight ?? "0"),
        dyDescent:
          node.attributes["x14ac:dyDescent"] !== undefined
            ? parseFloat(node.attributes["x14ac:dyDescent"])
            : undefined,
        outlineLevelRow: parseInt(node.attributes.outlineLevelRow ?? "0", 10),
        outlineLevelCol: parseInt(node.attributes.outlineLevelCol ?? "0", 10)
      };
      if (node.attributes.defaultColWidth) {
        this.model.defaultColWidth = parseFloat(node.attributes.defaultColWidth);
      }
      if (node.attributes.customHeight === "1") {
        this.model.customHeight = true;
      }
      return true;
    }
    return false;
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }
}

export { SheetFormatPropertiesXform };
