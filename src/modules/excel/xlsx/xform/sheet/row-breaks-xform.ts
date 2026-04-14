import { ListXform } from "@excel/xlsx/xform/list-xform";
import { PageBreaksXform } from "@excel/xlsx/xform/sheet/page-breaks-xform";

/**
 * Xform for row page breaks (rowBreaks element in worksheet XML)
 * Used to define manual page breaks between rows when printing.
 */
class RowBreaksXform extends ListXform {
  constructor() {
    super({
      tag: "rowBreaks",
      count: true,
      childXform: new PageBreaksXform()
    });
  }

  // Override to add manualBreakCount attribute required by Excel
  render(xmlStream: any, model: any): void {
    if (model && model.length) {
      xmlStream.openNode(this.tag, this.$);
      xmlStream.addAttribute(this.$count, model.length);
      xmlStream.addAttribute("manualBreakCount", model.length);

      const { childXform } = this;
      for (const childModel of model) {
        childXform.render(xmlStream, childModel);
      }
      xmlStream.closeNode();
    }
  }
}

export { RowBreaksXform };
