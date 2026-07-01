import type { AutoFilter } from "@excel/types";
import { colCache } from "@excel/utils/col-cache";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

class AutoFilterXform extends BaseXform<AutoFilter> {
  get tag(): string {
    return "autoFilter";
  }

  render(xmlStream: XmlSink, model?: AutoFilter): void {
    if (model) {
      if (typeof model === "string") {
        // assume range
        xmlStream.leafNode("autoFilter", { ref: model });
      } else {
        const getAddress = function (addr: string | { row: number; col: number }): string {
          if (typeof addr === "string") {
            return addr;
          }
          return colCache.getAddress(addr.row, addr.col).address;
        };

        const firstAddress = getAddress(model.from);
        const secondAddress = getAddress(model.to);
        if (firstAddress && secondAddress) {
          xmlStream.leafNode("autoFilter", { ref: `${firstAddress}:${secondAddress}` });
        }
      }
    }
  }

  parseOpen(node: ParseOpenTag): void {
    if (node.name === "autoFilter") {
      this.model = node.attributes.ref;
    }
  }
}

export { AutoFilterXform };
