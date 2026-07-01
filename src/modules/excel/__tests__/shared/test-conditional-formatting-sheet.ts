import { addSheetTo } from "@excel/__tests__/shared/add-sheet-to";
import conditionalFormattingJson from "@excel/__tests__/shared/data/conditional-formatting.json" with { type: "json" };
import { fix } from "@excel/__tests__/shared/tools";
import { Workbook, Worksheet } from "@excel/index";
import { expect } from "vitest";

const self = {
  conditionalFormattings: fix(conditionalFormattingJson),
  getConditionalFormatting(type: string) {
    return (self.conditionalFormattings as any)[type] || null;
  },
  addSheet(wb: any) {
    const ws = addSheetTo(wb, "conditional-formatting");
    const { types } = self.conditionalFormattings as any;
    types.forEach((type: string) => {
      const conditionalFormatting = self.getConditionalFormatting(type);
      if (conditionalFormatting) {
        if (typeof (ws as any).addConditionalFormatting === "function") {
          (ws as any).addConditionalFormatting(conditionalFormatting);
        } else {
          Worksheet.addConditionalFormatting(ws, conditionalFormatting);
        }
      }
    });
  },

  checkSheet(wb: any) {
    const ws = Workbook.getWorksheet(wb, "conditional-formatting")!;
    expect(ws).toBeDefined();
    expect(ws.conditionalFormattings).toBeDefined();
    ws.conditionalFormattings?.forEach((item: any) => {
      const type = item.rules && item.rules[0].type;
      const conditionalFormatting = self.getConditionalFormatting(type);
      expect(item).toHaveProperty("ref");
      expect(item).toHaveProperty("rules");
      expect((self.conditionalFormattings as any)[type]).toHaveProperty("ref");
      expect((self.conditionalFormattings as any)[type]).toHaveProperty("rules");
      expect(item.ref).toEqual(conditionalFormatting.ref);
      expect(item.rules.length).toBe(conditionalFormatting.rules.length);
    });
  }
};

const conditionalFormatting = self;
export { conditionalFormatting };
