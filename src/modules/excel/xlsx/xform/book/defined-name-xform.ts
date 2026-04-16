import type { DefinedNameModel } from "@excel/defined-names";
import { BaseXform } from "@excel/xlsx/xform/base-xform";

class DefinedNamesXform extends BaseXform {
  declare private _parsedName?: string;
  declare private _parsedLocalSheetId?: string;
  declare private _parsedText: string[];

  constructor() {
    super();
    this._parsedText = [];
  }

  render(xmlStream: any, model: DefinedNameModel): void {
    // <definedNames>
    //   <definedName name="name">text</definedName>
    //   <definedName name="_xlnm.Print_Area" localSheetId="0">text</definedName>
    // </definedNames>
    xmlStream.openNode("definedName", {
      name: model.name,
      localSheetId: model.localSheetId
    });
    // For opaque names, write the rawText verbatim to preserve round-trip fidelity.
    // For reference/formula names, join the ranges array as before.
    if (model.kind === "opaque" && model.rawText) {
      xmlStream.writeText(model.rawText);
    } else {
      xmlStream.writeText(model.ranges.join(","));
    }
    xmlStream.closeNode();
  }

  parseOpen(node: any): boolean {
    switch (node.name) {
      case "definedName":
        this._parsedName = node.attributes.name;
        this._parsedLocalSheetId = node.attributes.localSheetId;
        this._parsedText = [];
        return true;
      default:
        return false;
    }
  }

  parseText(text: string): void {
    this._parsedText.push(text);
  }

  /**
   * Stage 1 of the two-phase defined name design: the XLSX layer only
   * preserves the raw XML text.  Semantic classification (reference vs
   * formula vs opaque) is deferred to `DefinedNames.set model()`.
   */
  parseClose(): boolean {
    const rawText = this._parsedText.join("");

    const model: DefinedNameModel = {
      name: this._parsedName!,
      ranges: [],
      rawText: rawText.trim() || undefined
    };

    if (this._parsedLocalSheetId !== undefined) {
      model.localSheetId = parseInt(this._parsedLocalSheetId, 10);
    }
    this.model = model;
    return false;
  }
}

export { DefinedNamesXform };
