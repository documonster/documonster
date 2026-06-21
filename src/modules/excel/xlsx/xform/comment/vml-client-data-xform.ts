import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { VmlPositionXform } from "@excel/xlsx/xform/comment/style/vml-position-xform";
import { VmlProtectionXform } from "@excel/xlsx/xform/comment/style/vml-protection-xform";
import { VmlAnchorXform } from "@excel/xlsx/xform/comment/vml-anchor-xform";
import type { XmlSink } from "@xml/types";

const POSITION_TYPE = ["twoCells", "oneCells", "absolute"];

interface Protection {
  locked?: string;
  lockText?: string;
}

interface ClientDataModel {
  anchor: any;
  protection: Protection;
  editAs: string;
  row?: number;
  col?: number;
}

interface RenderModel {
  note: {
    protection: Protection;
    editAs: string;
  };
  refAddress: {
    row: number;
    col: number;
  };
}

class VmlClientDataXform extends BaseXform<ClientDataModel> {
  declare public map: { [key: string]: any };
  declare public parser: any;
  /** Name of the current simple leaf element being parsed (e.g. "x:Row"). */
  private _leafName: string | undefined;
  /** Accumulated text for the current leaf element. */
  private _leafText = "";

  constructor() {
    super();
    this.map = {
      "x:Anchor": new VmlAnchorXform(),
      "x:Locked": new VmlProtectionXform({ tag: "x:Locked" }),
      "x:LockText": new VmlProtectionXform({ tag: "x:LockText" }),
      "x:SizeWithCells": new VmlPositionXform({ tag: "x:SizeWithCells" }),
      "x:MoveWithCells": new VmlPositionXform({ tag: "x:MoveWithCells" })
    };
    this.model = { anchor: [], protection: {}, editAs: "" };
  }

  get tag(): string {
    return "x:ClientData";
  }

  render(xmlStream: XmlSink, model: RenderModel): void {
    const { protection, editAs } = model.note;
    xmlStream.openNode(this.tag, { ObjectType: "Note" });
    this.map["x:MoveWithCells"].render(xmlStream, editAs, POSITION_TYPE);
    this.map["x:SizeWithCells"].render(xmlStream, editAs, POSITION_TYPE);
    this.map["x:Anchor"].render(xmlStream, model);
    this.map["x:Locked"].render(xmlStream, protection.locked);
    xmlStream.leafNode("x:AutoFill", undefined, "False");
    this.map["x:LockText"].render(xmlStream, protection.lockText);
    xmlStream.leafNode("x:Row", undefined, model.refAddress.row - 1);
    xmlStream.leafNode("x:Column", undefined, model.refAddress.col - 1);
    xmlStream.closeNode();
  }

  parseOpen(node: any): boolean {
    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {
          anchor: [],
          protection: {},
          editAs: ""
        };
        break;
      case "x:Row":
      case "x:Column":
        this._leafName = node.name;
        this._leafText = "";
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
    if (this._leafName) {
      this._leafText += text;
    } else if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this._leafName) {
      if (name === this._leafName) {
        const value = parseInt(this._leafText, 10);
        if (name === "x:Row") {
          // VML uses 0-based row; convert to 1-based to match comment ref
          this.model!.row = value + 1;
        } else if (name === "x:Column") {
          // VML uses 0-based col; convert to 1-based to match comment ref
          this.model!.col = value + 1;
        }
        this._leafName = undefined;
        this._leafText = "";
      }
      return true;
    }
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        this.normalizeModel();
        return false;
      default:
        return true;
    }
  }

  normalizeModel(): void {
    const position = Object.assign(
      {},
      this.map["x:MoveWithCells"].model,
      this.map["x:SizeWithCells"].model
    );
    const len = Object.keys(position).length;
    this.model!.editAs = POSITION_TYPE[len];
    this.model!.anchor = this.map!["x:Anchor"].text;
    this.model!.protection.locked = this.map!["x:Locked"].text;
    this.model!.protection.lockText = this.map!["x:LockText"].text;
  }
}

export { VmlClientDataXform };
