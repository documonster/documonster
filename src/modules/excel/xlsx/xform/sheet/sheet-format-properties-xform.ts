import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { parseXsdFloatOrText, parseXsdIntOrText } from "@excel/xlsx/xform/xsd-values";
import type { ParseOpenTag, XmlAttributes, XmlSink } from "@xml/types";

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
      const attributes: XmlAttributes = {
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

      if (Object.values(attributes).some((value: unknown) => value !== undefined)) {
        xmlStream.leafNode("sheetFormatPr", attributes);
      }
    }
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (node.name === "sheetFormatPr") {
      // A value that is not a number keeps its text, for the reader to set aside (`sheet-settings-checks`).
      this.model = {
        defaultRowHeight: parseXsdFloatOrText(node.attributes.defaultRowHeight) ?? 0,
        dyDescent: parseXsdFloatOrText(node.attributes["x14ac:dyDescent"]),
        outlineLevelRow: parseXsdIntOrText(node.attributes.outlineLevelRow) ?? 0,
        outlineLevelCol: parseXsdIntOrText(node.attributes.outlineLevelCol) ?? 0
      } as SheetFormatPropertiesModel;
      if (node.attributes.defaultColWidth) {
        this.model.defaultColWidth = parseXsdFloatOrText(node.attributes.defaultColWidth);
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
