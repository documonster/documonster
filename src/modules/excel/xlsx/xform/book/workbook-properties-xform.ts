import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { XmlSink } from "@xml/types";

interface WorkbookPropertiesModel {
  date1904?: boolean;
}

class WorkbookPropertiesXform extends BaseXform {
  render(xmlStream: XmlSink, model: WorkbookPropertiesModel): void {
    xmlStream.leafNode("workbookPr", {
      date1904: model.date1904 ? 1 : undefined,
      // Excel doesn't output defaultThemeVersion
      filterPrivacy: 1
    });
  }

  parseOpen(node: any): boolean {
    if (node.name === "workbookPr") {
      this.model = {
        date1904: node.attributes.date1904 === "1"
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

export { WorkbookPropertiesXform };
