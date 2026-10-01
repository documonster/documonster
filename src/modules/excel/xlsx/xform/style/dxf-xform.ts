import type { Alignment, Borders, Fill, Font, NumFmt, Protection } from "@excel/types";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { PreservedSubtreeXform } from "@excel/xlsx/xform/preserved-xml-xform";
import { AlignmentXform } from "@excel/xlsx/xform/style/alignment-xform";
import { BorderXform } from "@excel/xlsx/xform/style/border-xform";
import { dxfExtensionsOf, setDxfExtensions } from "@excel/xlsx/xform/style/dxf-extensions";
import { FillXform } from "@excel/xlsx/xform/style/fill-xform";
import { FontXform } from "@excel/xlsx/xform/style/font-xform";
import { NumFmtXform } from "@excel/xlsx/xform/style/numfmt-xform";
import { ProtectionXform } from "@excel/xlsx/xform/style/protection-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

// <xf numFmtId="[numFmtId]" fontId="[fontId]" fillId="[fillId]" borderId="[xf.borderId]" xfId="[xfId]">
//   Optional <alignment>
//   Optional <protection>
// </xf>

interface DxfModel {
  alignment?: Partial<Alignment>;
  border?: Partial<Borders>;
  fill?: Fill;
  font?: Partial<Font>;
  /**
   * Either the format code (what the writer is handed) or the parsed
   * `{ id, formatCode }` pair (what {@link DxfXform.parseClose} produces when a
   * `dxf` is read back from a file). {@link DxfXform.render} accepts both —
   * before that it stringified the object into `formatCode="[object Object]"`,
   * corrupting the number format of every conditional format on re-save.
   */
  numFmt?: string | NumFmt;
  numFmtId?: number;
  protection?: Partial<Protection>;
}

// Style assists translation from style model to/from xlsx
class DxfXform extends BaseXform {
  declare public map: Record<string, BaseXform>;
  declare public parser?: BaseXform;
  /**
   * Depth inside a child this xform neither models nor preserves. Only `extLst` is legal after the six facets, so
   * this is for a producer that wrote something the schema does not allow: skipped whole, so that an element
   * nested inside it which happens to share a facet's name is not read as that facet.
   */
  declare private unknownDepth: number;

  constructor() {
    super();

    this.map = {
      alignment: new AlignmentXform(true),
      border: new BorderXform(),
      // Both in their differential form: in a `<dxf>` an absent `patternType` and an explicit
      // `<b val="0"/>` each mean something a cell's fill and font do not. See the two xforms.
      fill: new FillXform(true),
      font: new FontXform(FontXform.DIFFERENTIAL_OPTIONS),
      numFmt: new NumFmtXform(),
      protection: new ProtectionXform(),
      extLst: new PreservedSubtreeXform()
    };
    this.unknownDepth = 0;
  }

  get tag(): string {
    return "dxf";
  }

  // how do we generate dxfid?

  render(xmlStream: XmlSink, model: DxfModel): void {
    xmlStream.openNode(this.tag);

    if (model.font) {
      this.map.font.render(xmlStream, model.font);
    }
    // `!== undefined`, not truthiness: `General` is built-in id 0, and a truthiness test dropped it.
    if (model.numFmt && model.numFmtId !== undefined) {
      const formatCode = typeof model.numFmt === "string" ? model.numFmt : model.numFmt.formatCode;
      this.map.numFmt.render(xmlStream, { id: model.numFmtId, formatCode });
    }
    if (model.fill) {
      this.map.fill.render(xmlStream, model.fill);
    }
    if (model.alignment) {
      this.map.alignment.render(xmlStream, model.alignment);
    }
    if (model.border) {
      this.map.border.render(xmlStream, model.border);
    }
    if (model.protection) {
      this.map.protection.render(xmlStream, model.protection);
    }
    // Last, as `CT_Dxf` requires. See `dxf-extensions.ts`.
    const extensions = dxfExtensionsOf(model);
    if (extensions !== undefined) {
      xmlStream.writeRaw(extensions);
    }

    xmlStream.closeNode();
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    if (this.unknownDepth > 0) {
      this.unknownDepth++;
      return true;
    }
    switch (node.name) {
      case this.tag:
        // this node is often repeated. Need to reset children
        this.reset();
        this.unknownDepth = 0;
        return true;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        } else {
          this.unknownDepth = 1;
        }
        return true;
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
        this.parser = undefined;
      }
      return true;
    }
    if (this.unknownDepth > 0) {
      this.unknownDepth--;
      return true;
    }
    if (name === this.tag) {
      this.model = {
        alignment: this.map.alignment.model,
        border: this.map.border.model,
        fill: this.map.fill.model,
        font: this.map.font.model,
        numFmt: this.map.numFmt.model,
        protection: this.map.protection.model
      };
      const extensions = this.map.extLst.model as string | undefined;
      if (extensions !== undefined) {
        setDxfExtensions(this.model, extensions);
      }
      return false;
    }

    return true;
  }
}

/**
 * One entry of `<dxfs>`: a `<dxf>`, or an `<mc:AlternateContent>` offering several spellings of one.
 *
 * Markup Compatibility (ECMA-376 Part 3) lets a producer write a `<dxf>` that relies on its own extension inside an
 * `<mc:Choice Requires="…">`, with a plain one in `<mc:Fallback>`. The pair is still **one** entry: every `dxfId`
 * after it counts it once. Hancom Office writes every table-style format this way (`Requires="hs"`).
 *
 * A consumer takes the first `Choice` whose required namespaces it understands, else the `Fallback`. This reader
 * understands no extension namespace inside a `<dxf>`, so it takes the `Fallback` — which is the whole point of
 * one — and when there is none the entry still occupies its index, as an empty format.
 *
 * Before this, the first `<mc:…>` element ended the `<dxfs>` list: the first format was read, every later one was
 * lost, and each `dxfId` above 0 pointed past the end of the table.
 */
class DxfEntryXform extends BaseXform {
  declare private dxf: DxfXform;
  /** Inside an `<mc:AlternateContent>`: depth within it, and whether the current branch is the `Fallback`. */
  declare private alternate?: { depth: number; inFallback: boolean; chosen?: DxfModel };
  declare private delegating: boolean;

  constructor() {
    super();
    this.dxf = new DxfXform();
    this.delegating = false;
  }

  get tag(): string {
    return "dxf";
  }

  render(xmlStream: XmlSink, model: DxfModel): void {
    this.dxf.render(xmlStream, model);
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (this.delegating) {
      this.dxf.parseOpen(node);
      return true;
    }
    const alternate = this.alternate;
    if (alternate === undefined) {
      if (node.name === "mc:AlternateContent") {
        this.alternate = { depth: 1, inFallback: false };
        return true;
      }
      this.delegating = this.dxf.parseOpen(node);
      return this.delegating;
    }
    alternate.depth++;
    if (alternate.depth === 2) {
      alternate.inFallback = node.name === "mc:Fallback";
    } else if (alternate.depth === 3 && alternate.inFallback && node.name === "dxf") {
      this.delegating = this.dxf.parseOpen(node);
    }
    return true;
  }

  parseText(text: string): void {
    if (this.delegating) {
      this.dxf.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    const alternate = this.alternate;
    if (this.delegating) {
      if (this.dxf.parseClose(name)) {
        return true;
      }
      this.delegating = false;
      if (alternate === undefined) {
        this.model = this.dxf.model;
        return false;
      }
      alternate.chosen = this.dxf.model as DxfModel;
      alternate.depth--;
      return true;
    }
    if (alternate === undefined) {
      return false;
    }
    alternate.depth--;
    if (alternate.depth > 0) {
      return true;
    }
    this.model = alternate.chosen ?? {};
    this.alternate = undefined;
    return false;
  }

  reset(): void {
    super.reset();
    this.dxf.reset();
    this.alternate = undefined;
    this.delegating = false;
  }
}

export { DxfEntryXform, DxfXform };
