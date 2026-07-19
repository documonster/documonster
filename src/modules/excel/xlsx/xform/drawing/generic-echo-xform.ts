import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlAttributes, XmlSink } from "@xml/types";

export interface EchoNode {
  tag: string;
  attrs: XmlAttributes;
  children: EchoNode[];
  /** Ordered element/text content captured from the source subtree. */
  content?: Array<EchoNode | string>;
  /** Legacy single-text representation accepted when rendering older models. */
  text?: string;
}

/**
 * Generic capture-and-replay xform for an XML subtree this library does not
 * model structurally — e.g. `<xdr:grpSp>` grouped shapes, which can nest
 * arbitrary pic / sp / graphicFrame / grpSp children with their own coordinate
 * systems and relationship references.
 *
 * It parses the subtree into a plain tag/attributes/children/text tree and
 * re-emits it on write. The output is not guaranteed byte-identical to the
 * source (attribute insertion order may differ), but every element, attribute
 * and text node is preserved — which is strictly better than the previous
 * behaviour, where an unmapped `<xdr:grpSp>` caused the anchor's dispatch loop
 * to fall through and misinterpret a nested `<xdr:pic>` as the anchor's own
 * picture, silently discarding the rest of the group.
 */
class GenericEchoXform extends BaseXform<EchoNode> {
  private readonly rootTag: string;
  private stack: EchoNode[] = [];

  constructor(rootTag: string) {
    super();
    this.rootTag = rootTag;
  }

  get tag(): string {
    return this.rootTag;
  }

  render(xmlStream: XmlSink, model?: EchoNode): void {
    if (!model) {
      return;
    }
    this.renderNode(xmlStream, model);
  }

  private renderNode(xmlStream: XmlSink, node: EchoNode): void {
    const content =
      node.content ?? (node.text === undefined ? node.children : [node.text, ...node.children]);
    if (content.length === 0) {
      xmlStream.leafNode(node.tag, node.attrs);
      return;
    }
    xmlStream.openNode(node.tag, node.attrs);
    for (const item of content) {
      if (typeof item === "string") {
        xmlStream.writeText(item);
      } else {
        this.renderNode(xmlStream, item);
      }
    }
    xmlStream.closeNode();
  }

  reset(): void {
    super.reset();
    this.stack = [];
  }

  parseOpen(node: ParseOpenTag): boolean {
    // A fresh capture starts whenever the root tag opens at depth 0. Reset
    // first so a reused instance can't leak state from a previous subtree.
    if (node.name === this.rootTag && this.stack.length === 0) {
      this.reset();
    }
    const echoNode: EchoNode = {
      tag: node.name,
      attrs: node.attributes,
      children: [],
      content: []
    };
    if (this.stack.length > 0) {
      const parent = this.stack[this.stack.length - 1];
      parent.children.push(echoNode);
      parent.content!.push(echoNode);
    }
    this.stack.push(echoNode);
    return true;
  }

  parseText(text: string): void {
    if (this.stack.length === 0) {
      return;
    }
    const top = this.stack[this.stack.length - 1];
    const content = top.content!;
    const last = content[content.length - 1];
    if (typeof last === "string") {
      content[content.length - 1] = last + text;
    } else {
      content.push(text);
    }
  }

  parseClose(_name: string): boolean {
    const closed = this.stack.pop();
    if (this.stack.length === 0) {
      // Closed the root of the subtree — publish it and hand control back.
      this.model = closed;
      return false;
    }
    return true;
  }
}

export { GenericEchoXform };
