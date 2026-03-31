import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { WorksheetState } from "@excel/types";

const VALID_STATES: Set<string> = new Set(["visible", "hidden", "veryHidden"]);

function parseWorksheetState(raw: string | undefined): WorksheetState {
  const state = raw || "visible";
  return VALID_STATES.has(state) ? (state as WorksheetState) : "visible";
}

interface SheetModel {
  id: number;
  name: string;
  state: WorksheetState;
  rId: string;
}

class WorksheetXform extends BaseXform {
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
        id: parseInt(node.attributes.sheetId, 10),
        state: parseWorksheetState(node.attributes.state),
        rId: node.attributes["r:id"]
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

export { WorksheetXform };
