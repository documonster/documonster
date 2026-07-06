import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

// <cellStyle name="Heading 1" xfId="1" builtinId="16"/>
//
// A named cell style entry that maps a style name to a cellStyleXfs index.
export interface CellStyleModel {
  name: string;
  xfId: number;
  builtinId?: number;
  /** Hidden from the Excel "Cell Styles" gallery. Preserved for round-trip. */
  hidden?: boolean;
  /** Marks a customised built-in style. Preserved for round-trip. */
  customBuiltin?: boolean;
  /** Outline level (used by RowLevel_/ColLevel_ styles). Preserved for round-trip. */
  iLevel?: number;
}

// CellStyleXform translates a single <cellStyle> element to/from xlsx.
class CellStyleXform extends BaseXform {
  get tag(): string {
    return "cellStyle";
  }

  render(xmlStream: XmlSink, model: CellStyleModel): void {
    const attributes: Record<string, string | number> = {
      name: model.name,
      xfId: model.xfId
    };
    if (model.builtinId !== undefined) {
      attributes.builtinId = model.builtinId;
    }
    if (model.iLevel !== undefined) {
      attributes.iLevel = model.iLevel;
    }
    if (model.hidden) {
      attributes.hidden = "1";
    }
    if (model.customBuiltin) {
      attributes.customBuiltin = "1";
    }
    xmlStream.leafNode("cellStyle", attributes);
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (node.name === "cellStyle") {
      const model: CellStyleModel = {
        name: node.attributes.name,
        xfId: parseInt(node.attributes.xfId, 10)
      };
      if (node.attributes.builtinId !== undefined) {
        model.builtinId = parseInt(node.attributes.builtinId, 10);
      }
      if (node.attributes.iLevel !== undefined) {
        model.iLevel = parseInt(node.attributes.iLevel, 10);
      }
      if (node.attributes.hidden === "1") {
        model.hidden = true;
      }
      if (node.attributes.customBuiltin === "1") {
        model.customBuiltin = true;
      }
      this.model = model;
      return true;
    }
    return false;
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }
}

export { CellStyleXform };
