import { TwoCellAnchorXform } from "@excel/xlsx/xform/drawing/two-cell-anchor-xform";
import { describe, it, expect } from "vitest";

describe("TwoCellAnchorXform", () => {
  describe("reconcile", () => {
    it("handles null picture gracefully", () => {
      const twoCell = new TwoCellAnchorXform();
      const model: any = { picture: null };
      expect(() =>
        twoCell.reconcile(model, {} as Parameters<typeof twoCell.reconcile>[1])
      ).not.toThrow();
      // Model should remain intact with null picture
      expect(model.picture).toBeNull();
    });

    it("handles missing tl anchor gracefully", () => {
      const twoCell = new TwoCellAnchorXform();
      const model: any = { br: { col: 1, row: 1 } };
      expect(() =>
        twoCell.reconcile(model, {} as Parameters<typeof twoCell.reconcile>[1])
      ).not.toThrow();
      // br should be preserved, tl should remain absent
      expect(model.br).toEqual({ col: 1, row: 1 });
      expect(model.tl).toBeUndefined();
    });

    it("handles missing br anchor gracefully", () => {
      const twoCell = new TwoCellAnchorXform();
      const model: any = { tl: { col: 1, row: 1 } };
      expect(() =>
        twoCell.reconcile(model, {} as Parameters<typeof twoCell.reconcile>[1])
      ).not.toThrow();
      // tl should be preserved, br should remain absent
      expect(model.tl).toEqual({ col: 1, row: 1 });
      expect(model.br).toBeUndefined();
    });

    it("returns undefined when rId is missing from rels", () => {
      const twoCell = new TwoCellAnchorXform();
      const model: any = {
        picture: { rId: "rId99" },
        range: { editAs: "oneCell" }
      };
      const options = {
        rels: {
          // rId99 does not exist in rels
          rId1: { Target: "../media/image1.png" }
        },
        media: [],
        mediaIndex: {}
      };
      // Should not throw — previously this would crash with TypeError
      expect(() => twoCell.reconcile(model, options)).not.toThrow();
      // medium should be undefined since the rel was not found
      expect(model.medium).toBeUndefined();
    });
  });
});
