import { ListXform } from "@excel/xlsx/xform/list-xform";
import { PageBreaksXform } from "@excel/xlsx/xform/sheet/page-breaks-xform";
import type { PageBreakModel } from "@excel/xlsx/xform/sheet/page-breaks-xform";
import type { XmlSink } from "@xml/types";

/**
 * Xform for column page breaks (colBreaks element in worksheet XML)
 * Used to define manual page breaks between columns when printing.
 *
 * XML structure:
 * <colBreaks count="3" manualBreakCount="3">
 *   <brk id="3" max="1048575" man="1"/>
 *   <brk id="6" max="1048575" man="1"/>
 * </colBreaks>
 */
class ColBreaksXform extends ListXform<PageBreakModel> {
  constructor() {
    super({
      tag: "colBreaks",
      count: true,
      childXform: new PageBreaksXform()
    });
  }

  // Override to add manualBreakCount attribute required by Excel
  render(xmlStream: XmlSink, model?: PageBreakModel[]): void {
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

export { ColBreaksXform };
