import { XlsxParseError } from "@excel/errors";
import type { RichText } from "@excel/types";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { PhoneticTextXform } from "@excel/xlsx/xform/strings/phonetic-text-xform";
import { RichTextXform } from "@excel/xlsx/xform/strings/rich-text-xform";
import { TextXform } from "@excel/xlsx/xform/strings/text-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

// <si>
//   <r></r><r></r>...
// </si>
// <si>
//   <t></t>
// </si>

type SharedStringModel = string | { richText: RichText[] };

interface SharedStringXformOptions {
  /** The element read: `si` in the shared strings, `is` in a cell. Both are `CT_Rst`. Defaults to `si`. */
  readonly tag?: string;
  /** The error for content `CT_Rst` does not allow, given what is wrong with it. */
  readonly fail?: (detail: string) => Error;
}

/**
 * A rich string (`CT_Rst`): `<t>?`, then `<r>*`, then phonetic runs (`<rPh>*`, `<phoneticPr>?`), which are not modelled.
 *
 * Read by structure rather than by collecting text. An element's text may arrive in pieces, and gathering every `<t>`
 * into one string read a phonetic run's `<t>` — the reading of a Japanese string — as part of the string itself, and a
 * second `<t>` as more of the first. A `<t>` that `CT_Rst` does not allow — a second one, one after a run, a second
 * in one run — is refused; a `<t>` before runs (which the schema does allow) is the first run's text.
 */
class SharedStringXform extends BaseXform<SharedStringModel> {
  declare public map: Record<string, BaseXform>;
  declare public parser?: BaseXform;
  declare private readonly elementTag: string;
  declare private readonly fail: (detail: string) => Error;
  /** The `<t>` read, if one was. */
  declare private text: string | undefined;
  /** The `<r>`s read, if any were. */
  declare private runs: RichText[] | undefined;
  /** Depth inside a child that is not modelled (`<phoneticPr>`, a foreign element), whose subtree is skipped. */
  declare private skipping: number;

  constructor(model?: SharedStringModel, options: SharedStringXformOptions = {}) {
    super();

    this.model = model;
    this.elementTag = options.tag ?? "si";
    this.fail =
      options.fail ??
      (detail => new XlsxParseError(this.elementTag, `<${this.elementTag}> ${detail}`));
    this.skipping = 0;

    this.map = {
      r: new RichTextXform(),
      t: new TextXform(),
      rPh: new PhoneticTextXform()
    };
  }

  get tag(): string {
    return this.elementTag;
  }

  render(xmlStream: XmlSink, model?: SharedStringModel): void {
    xmlStream.openNode(this.tag);
    if (
      model &&
      typeof model === "object" &&
      Object.prototype.hasOwnProperty.call(model, "richText") &&
      model.richText
    ) {
      if (model.richText.length) {
        model.richText.forEach(text => {
          this.map.r.render(xmlStream, text);
        });
      } else {
        this.map.t.render(xmlStream, "");
      }
    } else if (model !== undefined && model !== null) {
      this.map.t.render(xmlStream, model as string);
    }
    xmlStream.closeNode();
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (this.parser) {
      if (
        node.name === "t" &&
        this.parser === this.map.r &&
        (this.parser.model as RichText | undefined)?.text !== undefined
      ) {
        throw this.fail("has a run with more than one <t>");
      }
      this.parser.parseOpen(node);
      return true;
    }
    if (this.skipping > 0) {
      this.skipping++;
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.text = undefined;
        this.runs = undefined;
        this.skipping = 0;
        return true;
      case "t":
        if (this.text !== undefined) {
          throw this.fail("has more than one <t>");
        }
        if (this.runs !== undefined) {
          throw this.fail("has a <t> after an <r>");
        }
        break;
      case "r":
      case "rPh":
        break;
      default:
        this.skipping = 1;
        return true;
    }
    this.parser = this.map[node.name];
    this.parser.parseOpen(node);
    return true;
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        switch (name) {
          case "r":
            (this.runs ??= []).push(this.parser.model as RichText);
            break;
          case "t":
            this.text = this.parser.model as string;
            break;
          default:
            break;
        }
        this.parser = undefined;
      }
      return true;
    }
    if (this.skipping > 0) {
      this.skipping--;
      return true;
    }
    if (name === this.tag) {
      this.model = this.result();
      return false;
    }
    return true;
  }

  /**
   * What has been read of an element cut short — a cell's `<is>`, ended by a `<c>` nested in it — with the run still
   * open kept if it has text, and the parser reset for the next element.
   */
  endEarly(): SharedStringModel {
    const run =
      this.parser === this.map.r ? (this.parser.model as RichText | undefined) : undefined;
    if (run?.text !== undefined) {
      (this.runs ??= []).push(run);
    }
    (this.map.r as RichTextXform).parser = undefined;
    (this.map.rPh as PhoneticTextXform).parser = undefined;
    this.parser = undefined;
    this.skipping = 0;
    return (this.model = this.result());
  }

  private result(): SharedStringModel {
    if (this.runs === undefined) {
      return this.text ?? "";
    }
    return { richText: this.text === undefined ? this.runs : [{ text: this.text }, ...this.runs] };
  }
}

export { SharedStringXform };
