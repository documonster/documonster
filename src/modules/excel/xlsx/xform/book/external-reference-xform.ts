/**
 * Xform for a single `<externalReference r:id="..."/>` element inside
 * `<externalReferences>` in `xl/workbook.xml`. Each `<externalReference>`
 * maps positionally (in document order, 1-based) to an `[N]Sheet!Ref`
 * prefix in formula strings.
 */

import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

export interface ExternalReferenceModel {
  rId: string;
}

class ExternalReferenceXform extends BaseXform<ExternalReferenceModel> {
  render(xmlStream: XmlSink, model: ExternalReferenceModel): void {
    xmlStream.leafNode("externalReference", { "r:id": model.rId });
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (node.name === "externalReference") {
      this.model = { rId: node.attributes["r:id"] ?? "" };
      return true;
    }
    return false;
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }
}

export { ExternalReferenceXform };
