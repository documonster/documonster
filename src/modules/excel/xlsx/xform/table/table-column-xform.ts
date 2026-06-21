import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { decodeOoxmlEscape, encodeOoxmlAttr } from "@utils/utils";
import type { XmlSink } from "@xml/types";

interface TableColumnModel {
  id?: number;
  name: string;
  totalsRowLabel?: string;
  totalsRowFunction?: string;
  totalsRowFormula?: string;
  calculatedColumnFormula?: string;
  dxfId?: string;
}

// Child elements of <tableColumn> whose text content we capture
type ChildTextTag = "calculatedColumnFormula" | "totalsRowFormula";

class TableColumnXform extends BaseXform<TableColumnModel> {
  private _childTag: ChildTextTag | undefined;
  private _childText = "";

  constructor() {
    super();
    this.model = { name: "" };
  }

  get tag(): string {
    return "tableColumn";
  }

  prepare(model: TableColumnModel, options: { index: number }): void {
    model.id = options.index + 1;
  }

  private _renderAttributes(model: TableColumnModel) {
    return {
      id: model.id!.toString(),
      name: encodeOoxmlAttr(model.name),
      totalsRowLabel: model.totalsRowLabel ? encodeOoxmlAttr(model.totalsRowLabel) : undefined,
      // Excel doesn't output totalsRowFunction when value is 'none'
      totalsRowFunction: model.totalsRowFunction === "none" ? undefined : model.totalsRowFunction,
      dxfId: model.dxfId
    };
  }

  render(xmlStream: XmlSink, model: TableColumnModel): void {
    // `<totalsRowFormula>` is only valid when `totalsRowFunction="custom"`
    // (or absent entirely, treated as custom). For the built-in
    // functions — `sum` / `average` / `count` / `countNums` / `max` /
    // `min` / `stdDev` / `var` — Excel GENERATES the SUBTOTAL formula
    // itself at open time. Emitting a redundant
    // `<totalsRowFormula>SUBTOTAL(101,Table[Col])</totalsRowFormula>`
    // alongside `totalsRowFunction="average"` makes Excel report
    // "Removed Records: Table from /xl/tables/tableN.xml" because the
    // schema allows the child only with a custom function. The
    // library's own `Table` class eagerly populates
    // `column.totalsRowFormula = getFormula(column)` for every
    // non-label column so rows render correctly in-memory; we gate
    // the *emission* here instead of mutating that in-memory cache.
    const isCustomFn =
      model.totalsRowFunction === undefined || model.totalsRowFunction === "custom";
    const emitTotalsFormula = !!model.totalsRowFormula && isCustomFn;
    if (model.calculatedColumnFormula || emitTotalsFormula) {
      xmlStream.openNode(this.tag, this._renderAttributes(model));
      if (model.calculatedColumnFormula) {
        xmlStream.leafNode("calculatedColumnFormula", undefined, model.calculatedColumnFormula);
      }
      if (emitTotalsFormula) {
        xmlStream.leafNode("totalsRowFormula", undefined, model.totalsRowFormula);
      }
      xmlStream.closeNode();
    } else {
      xmlStream.leafNode(this.tag, this._renderAttributes(model));
    }
  }

  parseOpen(node: any): boolean {
    if (node.name === this.tag) {
      const { attributes } = node;
      this.model = {
        name: decodeOoxmlEscape(attributes.name),
        totalsRowLabel: attributes.totalsRowLabel
          ? decodeOoxmlEscape(attributes.totalsRowLabel)
          : undefined,
        totalsRowFunction: attributes.totalsRowFunction,
        dxfId: attributes.dxfId
      };
      return true;
    }
    // Recognise child elements whose text content we want to capture
    if (node.name === "calculatedColumnFormula" || node.name === "totalsRowFormula") {
      this._childTag = node.name;
      this._childText = "";
    }
    return true;
  }

  parseText(text: string): void {
    if (this._childTag) {
      this._childText += text;
    }
  }

  parseClose(name: string): boolean {
    if (name === this.tag) {
      return false;
    }
    // Closing a recognised child element — store captured text
    if (this._childTag && name === this._childTag) {
      this.model![this._childTag] = this._childText;
      this._childTag = undefined;
    }
    return true;
  }
}

export { TableColumnXform };
