import type { IgnoredError } from "@excel/types";
import { BaseXform } from "@excel/xlsx/xform/base-xform";

/**
 * Boolean attribute names supported on <ignoredError> elements.
 * These correspond to the OOXML spec CT_IgnoredError attributes.
 */
const BOOL_ATTRS = [
  "numberStoredAsText",
  "formula",
  "formulaRange",
  "unlockedFormula",
  "emptyCellReference",
  "listDataValidation",
  "calculatedColumn",
  "evalError",
  "twoDigitTextYear"
] as const;

/**
 * Xform for the <ignoredErrors> element in a worksheet.
 *
 * Renders:
 * ```xml
 * <ignoredErrors>
 *   <ignoredError sqref="A1:B10" numberStoredAsText="1" />
 * </ignoredErrors>
 * ```
 */
class IgnoredErrorsXform extends BaseXform {
  declare public model: IgnoredError[];

  get tag(): string {
    return "ignoredErrors";
  }

  render(xmlStream: any, model: IgnoredError[] | undefined): void {
    if (!model || model.length === 0) {
      return;
    }
    xmlStream.openNode("ignoredErrors");
    for (const entry of model) {
      const attrs: Record<string, string | number> = { sqref: entry.ref };
      for (const attr of BOOL_ATTRS) {
        if (entry[attr]) {
          attrs[attr] = 1;
        }
      }
      xmlStream.leafNode("ignoredError", attrs);
    }
    xmlStream.closeNode();
  }

  parseOpen(node: any): boolean {
    switch (node.name) {
      case "ignoredErrors":
        this.model = [];
        return true;
      case "ignoredError": {
        const entry: IgnoredError = {
          ref: node.attributes.sqref ?? ""
        };
        for (const attr of BOOL_ATTRS) {
          if (node.attributes[attr] === "1" || node.attributes[attr] === "true") {
            entry[attr] = true;
          }
        }
        this.model.push(entry);
        return true;
      }
      default:
        return true;
    }
  }

  parseText(): void {
    // no text content
  }

  parseClose(name: string): boolean {
    switch (name) {
      case "ignoredErrors":
        return false;
      default:
        return true;
    }
  }
}

export { IgnoredErrorsXform };
