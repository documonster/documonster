import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { CellXform } from "@excel/xlsx/xform/sheet/cell-xform";
import { parseBoolean } from "@utils/utils";
import { colCache } from "@excel/utils/col-cache";
import { MaxItemsExceededError } from "@excel/errors";

interface RowXformOptions {
  maxItems?: number;
}

interface RowModel {
  number: number;
  min?: number;
  max?: number;
  cells: any[];
  styleId?: number;
  hidden?: boolean;
  bestFit?: boolean;
  height?: number;
  customHeight?: boolean;
  outlineLevel?: number;
  collapsed?: boolean;
  style?: any;
  dyDescent?: number;
}

class RowXform extends BaseXform<RowModel> {
  declare private maxItems?: number;
  declare public map: { [key: string]: any };
  declare public parser: any;
  declare private numRowsSeen: number;
  declare private lastCellCol: number;

  constructor(options?: RowXformOptions) {
    super();

    this.maxItems = options && options.maxItems;
    this.map = {
      c: new CellXform()
    };
  }

  get tag(): string {
    return "row";
  }

  reset(): void {
    super.reset();
    this.numRowsSeen = 0;
    this.lastCellCol = 0;
  }

  prepare(model: RowModel, options: any): void {
    const styleId = options.styles.addStyleModel(model.style);
    if (styleId) {
      model.styleId = styleId;
    }
    const cellXform = this.map.c;
    model.cells.forEach((cellModel: any) => {
      cellXform.prepare(cellModel, options);
    });
  }

  render(xmlStream: any, model?: RowModel, options?: any): void {
    if (!model) {
      return;
    }
    xmlStream.openNode("row");
    xmlStream.addAttribute("r", model.number);
    if (model.height != null && model.height > 0) {
      xmlStream.addAttribute("ht", model.height);
      if (model.customHeight !== false) {
        xmlStream.addAttribute("customHeight", "1");
      }
    } else if (model.height === 0) {
      // height=0 signals auto-height: write a minimal ht hint without
      // customHeight so Excel recalculates the row height on open.
      xmlStream.addAttribute("ht", 1);
    }
    if (model.hidden) {
      xmlStream.addAttribute("hidden", "1");
    }
    if (model.min! > 0 && model.max! > 0 && model.min! <= model.max!) {
      xmlStream.addAttribute("spans", `${model.min}:${model.max}`);
    }
    if (model.styleId) {
      xmlStream.addAttribute("s", model.styleId);
      xmlStream.addAttribute("customFormat", "1");
    }
    // Output dyDescent if present (MS extension for font descent)
    if (model.dyDescent !== undefined) {
      xmlStream.addAttribute("x14ac:dyDescent", model.dyDescent);
    }
    if (model.outlineLevel) {
      xmlStream.addAttribute("outlineLevel", model.outlineLevel);
    }
    if (model.collapsed) {
      xmlStream.addAttribute("collapsed", "1");
    }

    const cellXform = this.map.c;
    model.cells.forEach((cellModel: any) => {
      cellXform.render(xmlStream, cellModel, options);
    });

    xmlStream.closeNode();
  }

  parseOpen(node: any): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    if (node.name === "row") {
      this.numRowsSeen += 1;
      // Reset lastCellCol for each new row
      this.lastCellCol = 0;
      const spans = node.attributes.spans;
      let spanMin: number | undefined;
      let spanMax: number | undefined;
      if (spans) {
        const colonIdx = spans.indexOf(":");
        spanMin = parseInt(spans, 10); // parses up to non-digit
        spanMax = colonIdx > -1 ? parseInt(spans.substring(colonIdx + 1), 10) : undefined;
      }
      // If r attribute is missing, use numRowsSeen as the row number
      const rowNumber = node.attributes.r ? parseInt(node.attributes.r, 10) : this.numRowsSeen;
      const model: RowModel = (this.model = {
        number: rowNumber,
        min: spanMin,
        max: spanMax,
        cells: []
      });
      if (node.attributes.s) {
        model.styleId = parseInt(node.attributes.s, 10);
      }
      if (parseBoolean(node.attributes.hidden)) {
        model.hidden = true;
      }
      if (parseBoolean(node.attributes.bestFit)) {
        model.bestFit = true;
      }
      if (node.attributes.ht) {
        model.height = parseFloat(node.attributes.ht);
      }
      if (parseBoolean(node.attributes.customHeight)) {
        model.customHeight = true;
      }
      if (node.attributes.outlineLevel) {
        model.outlineLevel = parseInt(node.attributes.outlineLevel, 10);
      }
      if (parseBoolean(node.attributes.collapsed)) {
        model.collapsed = true;
      }
      if (node.attributes["x14ac:dyDescent"] !== undefined) {
        model.dyDescent = parseFloat(node.attributes["x14ac:dyDescent"]);
      }
      return true;
    }

    this.parser = this.map[node.name];
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    return false;
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        const cellModel = this.parser.model;
        // If cell has address, extract column number from it
        // Otherwise, calculate address based on position
        if (cellModel.address) {
          const decoded = colCache.decodeAddress(cellModel.address);
          this.lastCellCol = decoded.col;
        } else {
          // No r attribute, calculate address from position
          this.lastCellCol += 1;
          cellModel.address = colCache.encodeAddress(this.model!.number, this.lastCellCol);
        }
        this.model!.cells.push(cellModel);
        if (this.maxItems && this.model!.cells.length > this.maxItems) {
          throw new MaxItemsExceededError("column", this.maxItems);
        }
        this.parser = undefined;
      }
      return true;
    }
    return false;
  }

  reconcile(model: RowModel, options: any): void {
    model.style = model.styleId !== undefined ? options.styles.getStyleModel(model.styleId) : {};
    if (model.styleId !== undefined) {
      model.styleId = undefined;
    }

    const cellXform = this.map.c;
    model.cells.forEach((cellModel: any) => {
      cellXform.reconcile(cellModel, options);
    });
  }
}

export { RowXform };
