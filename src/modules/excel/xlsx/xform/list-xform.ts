import { ExcelLimitExceededError } from "@excel/errors";
import type { ExcelLimit } from "@excel/errors";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlAttributes, XmlSink } from "@xml/types";

/**
 * The surface ListXform requires of its child xform. Child renders may take
 * an optional positional index (e.g. vml-shape-xform), which BaseXform's
 * 2-arg render doesn't express, so it's modelled structurally here.
 */
interface ListChildXform<TChild> {
  prepare(model: TChild, options?: unknown): void;
  render(xmlStream: XmlSink, model: TChild, index?: number): void;
  parseOpen(node: ParseOpenTag): boolean | void;
  parseText(text: string): void;
  parseClose(name: string): boolean;
  reconcile(model: TChild, options?: unknown): void;
  reset(): void;
  model?: unknown;
}

interface ListXformOptions {
  tag: string;
  always?: boolean;
  count?: boolean;
  empty?: boolean;
  $count?: string;
  $?: XmlAttributes;
  // The list holds heterogeneous child models; the concrete element type is
  // recovered through the class generic `TChild`. `any` here lets any concrete
  // child xform be supplied without forcing every caller to thread the generic.
  childXform: ListChildXform<any>;
  maxItems?: number;
  /** The read option `maxItems` came from, which a refusal names. */
  limit?: ExcelLimit;
}

class ListXform<TChild = unknown> extends BaseXform<TChild[]> {
  declare protected tag: string;
  declare protected always: boolean;
  declare protected count?: boolean;
  declare protected empty?: boolean;
  declare public $count: string;
  declare public $?: XmlAttributes;
  declare protected childXform: ListChildXform<TChild>;
  declare protected maxItems?: number;
  declare protected limit?: ExcelLimit;
  declare public parser?: ListChildXform<TChild>;
  /** Depth inside an element the child parser declined, whose subtree is skipped. */
  private skipping = 0;

  constructor(options: ListXformOptions) {
    super();

    this.tag = options.tag;
    this.always = !!options.always;
    this.count = options.count;
    this.empty = options.empty;
    this.$count = options.$count ?? "count";
    this.$ = options.$;
    this.childXform = options.childXform as ListChildXform<TChild>;
    this.maxItems = options.maxItems;
    this.limit = options.limit;
  }

  prepare(model: TChild[], options: unknown): void {
    const { childXform } = this;
    if (model) {
      model.forEach((childModel, index) => {
        (options as { index?: number }).index = index;
        childXform.prepare(childModel, options);
      });
    }
  }

  render(xmlStream: XmlSink, model?: TChild[]): void {
    if (this.always || (model && model.length)) {
      xmlStream.openNode(this.tag, this.$);
      if (this.count) {
        xmlStream.addAttribute(this.$count, model?.length ?? 0);
      }

      const { childXform } = this;
      (model ?? []).forEach((childModel, index) => {
        childXform.render(xmlStream, childModel, index);
      });

      xmlStream.closeNode();
    } else if (this.empty) {
      xmlStream.leafNode(this.tag);
    }
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    if (this.skipping > 0) {
      this.skipping++;
      return true;
    }
    if (node.name === this.tag) {
      this.model = [];
      return true;
    }
    if (this.childXform.parseOpen(node)) {
      this.parser = this.childXform;
      return true;
    }
    // **An element the child does not take is skipped with its subtree**, once the list is open. Declining it
    // (`false`) left its closing tag to arrive here with no child parser, where it read as the list's own close —
    // ending `<sheetData>` at the first stray `</c>` and handing the rest of the rows to whatever parsed next.
    if (this.model !== undefined) {
      this.skipping = 1;
      return true;
    }
    return false;
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model!.push(this.parser.model as TChild);
        this.parser = undefined;

        if (this.maxItems && this.model!.length > this.maxItems) {
          // The concrete child may declare `tag` as a private member; probe it
          // structurally for the diagnostic message.
          const childTag = (this.childXform as unknown as { tag?: string }).tag ?? "";
          throw new ExcelLimitExceededError(childTag, this.maxItems, { limit: this.limit });
        }
      }
      return true;
    }
    if (this.skipping > 0) {
      this.skipping--;
      return true;
    }

    return false;
  }

  reconcile(model: TChild[], options: unknown): void {
    if (model) {
      const { childXform } = this;
      model.forEach(childModel => {
        childXform.reconcile(childModel, options);
      });
    }
  }

  reset(): void {
    super.reset();
    this.skipping = 0;
    if (this.childXform) {
      this.childXform.reset();
    }
  }
}

export { ListXform };
