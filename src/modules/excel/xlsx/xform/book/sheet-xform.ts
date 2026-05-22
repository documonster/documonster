import type { WorksheetState } from "@excel/types";
import { BaseXform } from "@excel/xlsx/xform/base-xform";

const VALID_STATES: Set<string> = new Set(["visible", "hidden", "veryHidden"]);

const RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function parseWorksheetState(raw: string | undefined): WorksheetState {
  const state = raw || "visible";
  return VALID_STATES.has(state) ? (state as WorksheetState) : "visible";
}

function parseSheetId(raw: string | undefined): number | undefined {
  if (raw === undefined) {
    return undefined;
  }
  const id = parseInt(raw, 10);
  // OOXML constrains `sheetId` to a positive integer. Anything that
  // doesn't parse to one — empty string, alphabetic, zero, negative,
  // overflowing — must not propagate as `NaN`/`0`/`-1` because those
  // would later seed pseudo keys like `_worksheets["NaN"]` (the same
  // family of bug as issue #166's `_worksheets["undefined"]`).
  return Number.isInteger(id) && id > 0 ? id : undefined;
}

interface SheetModel {
  id: number | undefined;
  name: string;
  state: WorksheetState;
  rId: string | undefined;
}

class WorksheetXform extends BaseXform {
  // All prefixes the workbook root binds to the OOXML relationships
  // namespace. Conventionally a workbook declares `xmlns:r=…`, so the
  // sheet uses `r:id="rId1"`. But the prefix is only a label and a
  // workbook may legally bind any prefix (or several) to that
  // namespace. `WorkbookXform` populates this list from the
  // `<workbook>` root; `r` is the safe fallback when the workbook
  // declares no relationships binding at all.
  relationshipsPrefixes: readonly string[] = ["r"];

  render(xmlStream: any, model: SheetModel): void {
    xmlStream.leafNode("sheet", {
      name: model.name,
      sheetId: model.id,
      // Excel doesn't output state when it's 'visible' (default)
      state: model.state === "visible" ? undefined : model.state,
      "r:id": model.rId
    });
  }

  parseOpen(node: any): boolean {
    if (node.name === "sheet") {
      this.model = {
        name: node.attributes.name,
        id: parseSheetId(node.attributes.sheetId),
        state: parseWorksheetState(node.attributes.state),
        rId: this._extractRelId(node)
      };
      return true;
    }
    return false;
  }

  /**
   * Locate the relationship id on a `<sheet>` element. Tries every
   * prefix the workbook root bound to the relationships namespace,
   * then any prefix the `<sheet>` element itself rebinds locally.
   * Returns `undefined` if no relationship id is present — callers
   * (the workbook reconciler) will treat such a `<sheet>` as a
   * half-broken declaration that can't be bound to a worksheet part.
   */
  private _extractRelId(node: any): string | undefined {
    const attrs = node.attributes ?? {};
    for (const prefix of this.relationshipsPrefixes) {
      const value = attrs[`${prefix}:id`];
      if (value !== undefined) {
        return value;
      }
    }
    // Local-scope fallback: a `<sheet>` element occasionally redeclares
    // the relationships namespace under a fresh prefix. Scan its own
    // attributes for `xmlns:X="…/relationships"` and look up `X:id`.
    for (const attrName of Object.keys(attrs)) {
      if (attrName.startsWith("xmlns:") && attrs[attrName] === RELATIONSHIPS_NS) {
        const localPrefix = attrName.slice("xmlns:".length);
        const candidate = attrs[`${localPrefix}:id`];
        if (candidate !== undefined) {
          return candidate;
        }
      }
    }
    return undefined;
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }
}

export { WorksheetXform };
