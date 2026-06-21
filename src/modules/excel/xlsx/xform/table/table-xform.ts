import type {
  CellFormulaValue,
  CellValue,
  Style,
  TableColumnProperties,
  TableStyleProperties
} from "@excel/types";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { ListXform } from "@excel/xlsx/xform/list-xform";
import { AutoFilterXform } from "@excel/xlsx/xform/table/auto-filter-xform";
import { TableColumnXform } from "@excel/xlsx/xform/table/table-column-xform";
import { TableStyleInfoXform } from "@excel/xlsx/xform/table/table-style-info-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";
import { StdDocAttributes } from "@xml/writer";

interface TableModel {
  id?: number;
  name: string;
  displayName?: string;
  ref?: string;
  tableRef: string;
  totalsRow?: boolean;
  headerRow?: boolean;
  columns?: TableColumnProperties[];
  rows?: Array<Array<CellValue | CellFormulaValue>>;
  autoFilterRef?: string;
  style?: TableStyleProperties;
}

interface TableXformOptions {
  styles?: {
    getDxfStyle(id: number): Partial<Style> | undefined;
  };
}

class TableXform extends BaseXform<TableModel> {
  declare public map: Record<string, BaseXform>;
  declare public parser?: BaseXform;

  constructor() {
    super();

    this.map = {
      autoFilter: new AutoFilterXform(),
      tableColumns: new ListXform({
        tag: "tableColumns",
        count: true,
        empty: true,
        childXform: new TableColumnXform()
      }),
      tableStyleInfo: new TableStyleInfoXform()
    };
    this.model = {
      id: 0,
      name: "",
      tableRef: "",
      columns: []
    };
  }

  prepare(model: TableModel, options: TableXformOptions): void {
    this.map.autoFilter.prepare(model);
    this.map.tableColumns.prepare(model.columns, options);
  }

  get tag(): string {
    return "table";
  }

  render(xmlStream: XmlSink, model: TableModel): void {
    xmlStream.openXml(StdDocAttributes);
    xmlStream.openNode(this.tag, {
      ...TableXform.TABLE_ATTRIBUTES,
      id: model.id,
      name: model.name,
      displayName: model.displayName || model.name,
      ref: model.tableRef,
      totalsRowCount: model.totalsRow ? "1" : undefined,
      // Excel doesn't output headerRowCount when it's 1 (default) or when there's a header row
      headerRowCount: model.headerRow ? undefined : "0"
    });

    this.map.autoFilter.render(xmlStream, model);
    this.map.tableColumns.render(xmlStream, model.columns);
    this.map.tableStyleInfo.render(xmlStream, model.style);

    xmlStream.closeNode();
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    const { name, attributes } = node;
    switch (name) {
      case this.tag:
        this.reset();
        this.model = {
          name: attributes.name,
          displayName: attributes.displayName || attributes.name,
          tableRef: attributes.ref,
          totalsRow: attributes.totalsRowCount === "1",
          // ECMA-376: headerRowCount defaults to 1, so missing attribute means has header
          headerRow: attributes.headerRowCount !== "0"
        };
        break;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        this.model!.columns = this.map!.tableColumns.model as TableColumnProperties[];
        {
          const autoFilterModel = this.map!.autoFilter.model as
            | { autoFilterRef?: string; columns: { filterButton?: boolean }[] }
            | undefined;
          if (autoFilterModel) {
            this.model!.autoFilterRef = autoFilterModel.autoFilterRef;
            autoFilterModel.columns.forEach((column, index) => {
              (
                this.model!.columns![index] as TableColumnProperties & { filterButton?: boolean }
              ).filterButton = column.filterButton;
            });
          }
        }
        this.model!.style = this.map!.tableStyleInfo.model as TableStyleProperties;
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  }

  reconcile(model: TableModel, options: TableXformOptions): void {
    // Map tableRef to ref for Table constructor compatibility
    if (model.tableRef && !model.ref) {
      model.ref = model.tableRef;
    }
    // Add empty rows array if not present (tables loaded from file don't have row data)
    if (!model.rows) {
      model.rows = [];
    }
    // fetch the dfxs from styles
    const styles = options.styles;
    if (styles) {
      model.columns!.forEach(columnModel => {
        // dxfId is a transient (de)serialisation field carried on the column.
        const column = columnModel as TableColumnProperties & { dxfId?: number };
        if (column.dxfId !== undefined) {
          column.style = styles.getDxfStyle(column.dxfId);
        }
      });
    }
  }

  static TABLE_ATTRIBUTES = {
    xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  };
}

export { TableXform };
