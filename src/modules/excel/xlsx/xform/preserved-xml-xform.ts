import { XlsxParseError } from "@excel/errors";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { FragmentNamespaces } from "@excel/xlsx/xform/fragment-namespaces";
import type { ParseOpenTag } from "@xml/types";
import { XmlWriter } from "@xml/writer";

/** Preserves an attribute-only OOXML element while rejecting nested content. */
class PreservedLeafXform extends BaseXform {
  declare private readonly elementName: string;
  declare private writer?: XmlWriter;
  declare private namespaces?: FragmentNamespaces;
  declare private readonly enclosing?: () => Readonly<Record<string, string>> | undefined;

  /**
   * @param enclosing - The declarations the element this one is written back into carries, when its
   * xform keeps them — see `FragmentNamespaces`.
   */
  constructor(elementName: string, enclosing?: () => Readonly<Record<string, string>> | undefined) {
    super();
    this.elementName = elementName;
    this.enclosing = enclosing;
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (node.name !== this.elementName || this.writer) {
      throw new XlsxParseError(
        this.elementName,
        `Unexpected xml node in parseOpen: ${JSON.stringify(node)}`
      );
    }
    this.writer = new XmlWriter();
    this.namespaces = new FragmentNamespaces({ enclosing: this.enclosing?.() });
    const tag = this.namespaces.open(node);
    this.writer.openNode(tag.name, tag.attributes);
    return true;
  }

  parseText(text: string): void {
    if (text.trim()) {
      throw new XlsxParseError(this.elementName, "Unexpected text content");
    }
  }

  parseClose(name: string): boolean {
    if (name !== this.elementName || !this.writer) {
      throw new XlsxParseError(this.elementName, `Unexpected xml node in parseClose: ${name}`);
    }
    this.namespaces!.close(name);
    this.writer.closeNode();
    this.model = this.namespaces!.finish(this.writer.xml);
    this.writer = undefined;
    this.namespaces = undefined;
    return false;
  }

  reset(): void {
    super.reset();
    this.writer = undefined;
    this.namespaces = undefined;
  }
}

/** Preserves an OOXML subtree's XML semantics across read/write round trips. */
class PreservedSubtreeXform extends BaseXform {
  declare private depth: number;
  declare private writer?: XmlWriter;
  /** The subtree's text views and the declarations it needs — see `fragment-namespaces.ts`. */
  declare private namespaces?: FragmentNamespaces;
  declare private readonly enclosing?: () => Readonly<Record<string, string>> | undefined;

  /**
   * @param enclosing - The declarations the element this one is written back into carries, when its
   * xform keeps them — see `FragmentNamespaces`.
   */
  constructor(enclosing?: () => Readonly<Record<string, string>> | undefined) {
    super();
    this.depth = 0;
    this.enclosing = enclosing;
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (!this.writer) {
      this.writer = new XmlWriter();
      this.namespaces = new FragmentNamespaces({ enclosing: this.enclosing?.() });
    }
    const tag = this.namespaces!.open(node);
    this.writer.openNode(tag.name, tag.attributes);
    this.depth += 1;
    return true;
  }

  parseText(text: string): void {
    // CDATA arrives here too: `BaseXform.parseCdata` maps it onto `parseText`,
    // so the value is preserved as escaped text rather than a CDATA section.
    this.writer?.writeText(text);
  }

  parseClose(name: string): boolean {
    this.namespaces!.close(name);
    this.writer!.closeNode();
    this.depth -= 1;
    if (this.depth > 0) {
      return true;
    }
    this.model = this.namespaces!.finish(this.writer!.xml);
    this.writer = undefined;
    this.namespaces = undefined;
    return false;
  }

  reset(): void {
    super.reset();
    this.depth = 0;
    this.writer = undefined;
    this.namespaces = undefined;
  }
}

export { PreservedLeafXform, PreservedSubtreeXform };
