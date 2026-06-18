import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { RichTextXform } from "@excel/xlsx/xform/strings/rich-text-xform";
import { TextXform } from "@excel/xlsx/xform/strings/text-xform";

interface NoteText {
  font?: any;
  text: string;
}

interface CommentNote {
  texts: NoteText[];
}

interface CommentModel {
  type: string;
  note: CommentNote;
  ref: string;
  authorId?: number;
}

class CommentXform extends BaseXform<CommentModel> {
  declare public parser: any;
  declare private _richTextXform?: RichTextXform;
  declare private _textXform?: TextXform;

  constructor(model?: CommentModel) {
    super();
    this.model = model || { type: "note", note: { texts: [] }, ref: "" };
  }

  get tag(): string {
    return "r";
  }

  get richTextXform(): RichTextXform {
    if (!this._richTextXform) {
      this._richTextXform = new RichTextXform();
    }
    return this._richTextXform;
  }

  get textXform(): TextXform {
    if (!this._textXform) {
      this._textXform = new TextXform();
    }
    return this._textXform;
  }

  render(xmlStream: any, model?: CommentModel): void {
    const renderModel = model || this.model;

    xmlStream.openNode("comment", {
      ref: renderModel!.ref,
      authorId: renderModel!.authorId ?? 0
    });
    xmlStream.openNode("text");
    if (renderModel && renderModel.note && renderModel.note.texts) {
      renderModel.note.texts.forEach(text => {
        this.richTextXform.render(xmlStream, text);
      });
    }
    xmlStream.closeNode();
    xmlStream.closeNode();
  }

  parseOpen(node: any): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case "comment":
        this.model = {
          type: "note",
          note: {
            texts: []
          },
          ref: node.attributes.ref,
          authorId: node.attributes.authorId != null ? Number(node.attributes.authorId) : undefined
        };
        return true;
      case "r":
        this.parser = this.richTextXform;
        this.parser.parseOpen(node);
        return true;
      case "t":
        // Legacy comments (e.g. produced by other tools) may store the
        // body as a bare <t> directly under <text> with no <r> run wrapper.
        // This is valid for the CT_Rst type, so treat it like a run without font.
        this.parser = this.textXform;
        this.parser.parseOpen(node);
        return true;
      default:
        return false;
    }
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        // The active sub-parser has finished. Collect its result.
        if (this.parser === this._richTextXform) {
          // <r> run: model is already a { font?, text } run.
          this.model!.note.texts.push(this.parser.model);
        } else {
          // Bare <t> body (e.g. from other tools): wrap the plain string
          // as a single run without font, mirroring a <r><t> run.
          this.model!.note.texts.push({ text: this.parser.model });
        }
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case "comment":
        return false;
      default:
        return true;
    }
  }
}

export { CommentXform };
