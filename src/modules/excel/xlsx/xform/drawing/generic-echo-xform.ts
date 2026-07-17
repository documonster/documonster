import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlAttributes, XmlSink } from "@xml/types";

export interface EchoNode {
  tag: string;
  attrs: XmlAttributes;
  children: EchoNode[];
  text?: string;
}

/**
 * Generic capture-and-replay xform for an XML subtree this library doesn't
 * model structurally (e.g. `<xdr:grpSp>` grouped shapes, which can nest
 * arbitrary pic/sp/graphicFrame/grpSp children with their own coordinate
 * systems). Parses into a plain tag/attributes/children tree and re-emits
 * it on write. Not guaranteed byte-identical (e.g. attribute insertion
 * order may differ) but preserves every element, attribute, and text node -
 * unlike silently dropping or misinterpreting the content.
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
    if (node.children.length === 0 && node.text === undefined) {
      xmlStream.leafNode(node.tag, node.attrs);
      return;
    }
    xmlStream.openNode(node.tag, node.attrs);
    if (node.text !== undefined) {
      xmlStream.writeText(node.text);
    }
    for (const child of node.children) {
      this.renderNode(xmlStream, child);
    }
    xmlStream.closeNode();
  }

  reset(): void {
    super.reset();
    this.stack = [];
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (node.name === this.rootTag && this.stack.length === 0) {
      this.reset();
    }
    const echoNode: EchoNode = { tag: node.name, attrs: node.attributes, children: [] };
    if (this.stack.length > 0) {
      this.stack[this.stack.length - 1].children.push(echoNode);
    }
    this.stack.push(echoNode);
    return true;
  }

  parseText(text: string): void {
    if (this.stack.length === 0) {
      return;
    }
    const top = this.stack[this.stack.length - 1];
    top.text = (top.text ?? "") + text;
  }

  parseClose(name: string): boolean {
    const closed = this.stack.pop();
    if (this.stack.length === 0) {
      this.model = closed;
      return false;
    }
    return true;
  }
}

export { GenericEchoXform };
