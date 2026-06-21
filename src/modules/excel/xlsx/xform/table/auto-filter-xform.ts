import { XlsxParseError } from "@excel/errors";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { FilterColumnXform } from "@excel/xlsx/xform/table/filter-column-xform";
import type { FilterColumnModel } from "@excel/xlsx/xform/table/filter-column-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

interface AutoFilterModel {
  autoFilterRef: string;
  columns: FilterColumnModel[];
}

class AutoFilterXform extends BaseXform<AutoFilterModel> {
  declare public map: { [key: string]: FilterColumnXform };
  declare public parser?: BaseXform;

  constructor() {
    super();

    this.map = {
      filterColumn: new FilterColumnXform()
    };
    this.model = { autoFilterRef: "", columns: [] };
  }

  get tag(): string {
    return "autoFilter";
  }

  prepare(model: AutoFilterModel): void {
    model.columns.forEach((column, index) => {
      this.map.filterColumn.prepare(column, { index });
    });
  }

  render(xmlStream: XmlSink, model: AutoFilterModel): void {
    xmlStream.openNode(this.tag, {
      ref: model.autoFilterRef
    });

    // Only emit `<filterColumn>` for columns that carry actual filter
    // state. Real Excel only emits the child when a filter is applied
    // (`customFilters` / `filters` / `dynamicFilter`) or when the
    // author explicitly set the filter-button visibility (either
    // `filterButton: true` or `filterButton: false`). Columns that
    // never touched `filterButton` (i.e. `undefined`) default to
    // Excel's "show button" behaviour and should emit nothing —
    // emitting an empty `<filterColumn hiddenButton="1"/>` for
    // every such column makes Excel reject the table with
    // "Removed Records: Table from /xl/tables/tableN.xml".
    model.columns.forEach(column => {
      if (
        column?.customFilters !== undefined ||
        column?.filters !== undefined ||
        column?.dynamicFilter !== undefined ||
        column?.filterButton !== undefined
      ) {
        this.map.filterColumn.render(xmlStream, column);
      }
    });

    xmlStream.closeNode();
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.model = {
          autoFilterRef: node.attributes.ref,
          columns: []
        };
        return true;

      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parseOpen(node);
          return true;
        }
        throw new XlsxParseError(
          "autoFilter",
          `Unexpected xml node in parseOpen: ${JSON.stringify(node)}`
        );
    }
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model!.columns.push(this.parser.model);
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        return false;
      default:
        throw new XlsxParseError("autoFilter", `Unexpected xml node in parseClose: ${name}`);
    }
  }
}

export { AutoFilterXform };
