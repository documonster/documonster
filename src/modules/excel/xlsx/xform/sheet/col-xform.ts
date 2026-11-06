import type { Style } from "@excel/types";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { parseBoolean } from "@utils/utils";
import type { ParseOpenTag, XmlSink } from "@xml/types";

interface ColModel {
  min: number;
  max: number;
  width?: number;
  styleId?: number;
  hidden?: boolean;
  bestFit?: boolean;
  /** `customWidth` as read; absent for a column a caller made, where a width implies it. */
  customWidth?: boolean;
  outlineLevel?: number;
  collapsed?: boolean;
  style?: Partial<Style>;
}

/** Style-manager surface used by the column xform's prepare/reconcile. */
interface ColXformOptions {
  styles: {
    addStyleModel(model: Partial<Style>): number;
    getStyleModel(id: number): Style | null;
  };
}

class ColXform extends BaseXform {
  get tag(): string {
    return "col";
  }

  prepare(model: ColModel, options: ColXformOptions): void {
    const styleId = options.styles.addStyleModel(model.style || {});
    if (styleId) {
      model.styleId = styleId;
    }
  }

  render(xmlStream: XmlSink, model: ColModel): void {
    xmlStream.openNode("col");
    xmlStream.addAttribute("min", model.min);
    xmlStream.addAttribute("max", model.max);
    // Zero is a width — Excel hides a column by giving it one — so only an absent width goes unwritten.
    if (model.width !== undefined) {
      xmlStream.addAttribute("width", model.width);
    }
    if (model.styleId) {
      xmlStream.addAttribute("style", model.styleId);
    }
    if (model.hidden) {
      xmlStream.addAttribute("hidden", "1");
    }
    if (model.bestFit) {
      xmlStream.addAttribute("bestFit", "1");
    }
    if (model.outlineLevel) {
      xmlStream.addAttribute("outlineLevel", model.outlineLevel);
    }
    if (model.collapsed) {
      xmlStream.addAttribute("collapsed", "1");
    }
    // Only where the file said so, or where a caller set the width: written on every column, it turned a width Excel
    // had sized to the default font into one the author chose.
    if (model.customWidth ?? model.width !== undefined) {
      xmlStream.addAttribute("customWidth", "1");
    }
    xmlStream.closeNode();
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (node.name === "col") {
      const model: ColModel = (this.model = {
        min: parseInt(node.attributes.min ?? "0", 10),
        max: parseInt(node.attributes.max ?? "0", 10),
        width:
          node.attributes.width === undefined ? undefined : parseFloat(node.attributes.width ?? "0")
      });
      if (node.attributes.style) {
        model.styleId = parseInt(node.attributes.style, 10);
      }
      if (parseBoolean(node.attributes.hidden)) {
        model.hidden = true;
      }
      if (parseBoolean(node.attributes.bestFit)) {
        model.bestFit = true;
      }
      // Recorded only where it differs from what the width implies, which is what the writer assumes without it.
      const customWidth = parseBoolean(node.attributes.customWidth);
      if (customWidth !== (model.width !== undefined)) {
        model.customWidth = customWidth;
      }
      if (node.attributes.outlineLevel) {
        model.outlineLevel = parseInt(node.attributes.outlineLevel, 10);
      }
      if (parseBoolean(node.attributes.collapsed)) {
        model.collapsed = true;
      }
      return true;
    }
    return false;
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }

  reconcile(model: ColModel, options: ColXformOptions): void {
    // reconcile column styles
    if (model.styleId !== undefined) {
      model.style = options.styles.getStyleModel(model.styleId) ?? undefined;
    }
  }
}

export { ColXform };
