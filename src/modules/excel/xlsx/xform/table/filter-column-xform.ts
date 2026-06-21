import { XlsxParseError } from "@excel/errors";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { ListXform } from "@excel/xlsx/xform/list-xform";
import { CustomFilterXform } from "@excel/xlsx/xform/table/custom-filter-xform";
import { FilterXform } from "@excel/xlsx/xform/table/filter-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

interface FilterColumnModel {
  colId?: string;
  filterButton: boolean;
  customFilters?: any[];
}

class FilterColumnXform extends BaseXform<FilterColumnModel> {
  declare public map: { [key: string]: ListXform };
  declare public parser?: BaseXform;

  constructor() {
    super();

    this.map = {
      customFilters: new ListXform({
        tag: "customFilters",
        count: false,
        empty: true,
        childXform: new CustomFilterXform()
      }),
      filters: new ListXform({
        tag: "filters",
        count: false,
        empty: true,
        childXform: new FilterXform()
      })
    };
    this.model = { filterButton: false };
  }

  get tag(): string {
    return "filterColumn";
  }

  prepare(model: FilterColumnModel, options: { index: number }): void {
    model.colId = options.index.toString();
  }

  render(xmlStream: XmlSink, model: FilterColumnModel): void {
    if (model.customFilters) {
      xmlStream.openNode(this.tag, {
        colId: model.colId,
        hiddenButton: model.filterButton ? "0" : "1"
      });

      this.map.customFilters.render(xmlStream, model.customFilters);

      xmlStream.closeNode();
      return;
    }
    xmlStream.leafNode(this.tag, {
      colId: model.colId,
      hiddenButton: model.filterButton ? "0" : "1"
    });
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    const { attributes } = node;
    switch (node.name) {
      case this.tag:
        this.model = {
          filterButton: attributes.hiddenButton === "0"
        };
        return true;
      case "dynamicFilter":
        // Ignore dynamicFilter nodes - we don't need to preserve them for reading
        return true;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parseOpen(node);
          return true;
        }
        throw new XlsxParseError(
          "filterColumn",
          `Unexpected xml node in parseOpen: ${JSON.stringify(node)}`
        );
    }
  }

  parseText(): void {}

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        this.model!.customFilters = this.map!.customFilters.model;
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  }
}

export { FilterColumnXform };
