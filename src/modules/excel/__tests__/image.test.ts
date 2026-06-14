/**
 * Excel Image Class Unit Tests
 *
 * Tests for the Image class:
 * - Construction with different model inputs
 * - Model getter for background and image types
 * - Range parsing from string and object inputs
 * - Hyperlinks handling
 */

import { applyImageModel, imageClone, imageCreate, imageModel } from "@excel/image";
import { Workbook } from "@excel/index";
import type { WorkbookData } from "@excel/workbook-core";
import type { WorksheetData } from "@excel/worksheet-core";
import { describe, it, expect, beforeEach } from "vitest";

// =============================================================================
// Test Setup
// =============================================================================

describe("Image", () => {
  let workbook: WorkbookData;
  let worksheet: WorksheetData;

  beforeEach(() => {
    workbook = Workbook.create();
    worksheet = Workbook.addWorksheet(workbook, "Sheet1");
  });

  // ===========================================================================
  // Constructor Tests
  // ===========================================================================

  describe("constructor", () => {
    it("creates image without model", () => {
      const image = imageCreate(worksheet);
      expect(image.worksheet).toBe(worksheet);
      expect(image.type).toBeUndefined();
      expect(image.imageId).toBeUndefined();
      expect(image.range).toBeUndefined();
    });

    it("creates background image with model", () => {
      const image = imageCreate(worksheet, {
        type: "background",
        imageId: "img1"
      });
      expect(image.type).toBe("background");
      expect(image.imageId).toBe("img1");
    });

    it("creates positioned image with string range", () => {
      const image = imageCreate(worksheet, {
        type: "image",
        imageId: "img2",
        range: "B2:D6"
      });
      expect(image.type).toBe("image");
      expect(image.imageId).toBe("img2");
      expect(image.range).toBeDefined();
      expect(image.range!.editAs).toBe("oneCell");
    });

    it("creates positioned image with object range", () => {
      const image = imageCreate(worksheet, {
        type: "image",
        imageId: "img3",
        range: {
          tl: { col: 1, row: 1 },
          br: { col: 3, row: 5 },
          editAs: "twoCell"
        }
      });
      expect(image.type).toBe("image");
      expect(image.imageId).toBe("img3");
      expect(image.range!.editAs).toBe("twoCell");
    });
  });

  // ===========================================================================
  // Model Getter Tests
  // ===========================================================================

  describe("model getter", () => {
    it("returns background model", () => {
      const image = imageCreate(worksheet, {
        type: "background",
        imageId: "bg1"
      });

      const model = imageModel(image);
      expect(model.type).toBe("background");
      expect(model.imageId).toBe("bg1");
    });

    it("returns image model with range", () => {
      const image = imageCreate(worksheet, {
        type: "image",
        imageId: "img1",
        range: "A1:C3"
      });

      const model = imageModel(image);
      expect(model.type).toBe("image");
      expect(model.imageId).toBe("img1");
      if (model.type === "image") {
        expect(model.range).toBeDefined();
        expect(model.range.tl).toBeDefined();
        expect(model.range.br).toBeDefined();
      }
    });

    it("throws error for invalid type", () => {
      const image = imageCreate(worksheet);
      image.type = "invalid";
      image.imageId = "test";

      expect(() => imageModel(image)).toThrow("Invalid Image Type");
    });
  });

  // ===========================================================================
  // Model Setter Tests
  // ===========================================================================

  describe("model setter", () => {
    it("sets background type", () => {
      const image = imageCreate(worksheet);
      applyImageModel(image, {
        type: "background",
        imageId: "bg2"
      });

      expect(image.type).toBe("background");
      expect(image.imageId).toBe("bg2");
    });

    it("sets image type with string range", () => {
      const image = imageCreate(worksheet);
      applyImageModel(image, {
        type: "image",
        imageId: "img4",
        range: "B2:D4"
      });

      expect(image.type).toBe("image");
      expect(image.imageId).toBe("img4");
      expect(image.range).toBeDefined();
    });

    it("sets image type with object range containing native offsets", () => {
      const image = imageCreate(worksheet);
      applyImageModel(image, {
        type: "image",
        imageId: "img5",
        range: {
          tl: { nativeCol: 1, nativeRow: 1, nativeColOff: 100, nativeRowOff: 100 },
          br: { nativeCol: 3, nativeRow: 5, nativeColOff: 0, nativeRowOff: 0 }
        }
      });

      expect(image.range).toBeDefined();
      expect(image.range!.tl.nativeCol).toBe(1);
      expect(image.range!.tl.nativeRow).toBe(1);
    });

    it("sets image with ext dimensions", () => {
      const image = imageCreate(worksheet);
      applyImageModel(image, {
        type: "image",
        imageId: "img6",
        range: {
          tl: { col: 0, row: 0 },
          ext: { width: 100, height: 200 }
        }
      });

      expect(image.range!.ext).toEqual({ width: 100, height: 200 });
    });
  });

  // ===========================================================================
  // Hyperlinks Tests
  // ===========================================================================

  describe("hyperlinks", () => {
    it("sets hyperlinks from model input", () => {
      const image = imageCreate(worksheet, {
        type: "image",
        imageId: "img7",
        range: {
          tl: { col: 0, row: 0 },
          br: { col: 2, row: 2 }
        },
        hyperlinks: {
          hyperlink: "https://example.com",
          tooltip: "Click me"
        }
      });

      expect(image.range!.hyperlinks).toEqual({
        hyperlink: "https://example.com",
        tooltip: "Click me"
      });
    });

    it("includes hyperlinks in model output", () => {
      const image = imageCreate(worksheet, {
        type: "image",
        imageId: "img8",
        range: {
          tl: { col: 0, row: 0 }
        },
        hyperlinks: {
          hyperlink: "https://test.com"
        }
      });

      const model = imageModel(image);
      if (model.type === "image") {
        expect(model.hyperlinks).toEqual({ hyperlink: "https://test.com" });
      }
    });

    it("handles hyperlinks in range object", () => {
      const image = imageCreate(worksheet);
      applyImageModel(image, {
        type: "image",
        imageId: "img9",
        range: {
          tl: { col: 0, row: 0 },
          hyperlinks: {
            hyperlink: "https://range-hyperlink.com"
          }
        }
      });

      expect(image.range!.hyperlinks).toEqual({
        hyperlink: "https://range-hyperlink.com"
      });
    });
  });

  // ===========================================================================
  // editAs Tests
  // ===========================================================================

  describe("editAs property", () => {
    it("defaults to oneCell for string ranges", () => {
      const image = imageCreate(worksheet, {
        type: "image",
        imageId: "img10",
        range: "A1:B2"
      });

      expect(image.range!.editAs).toBe("oneCell");
    });

    it("preserves editAs from object range", () => {
      const image = imageCreate(worksheet, {
        type: "image",
        imageId: "img11",
        range: {
          tl: { col: 0, row: 0 },
          editAs: "absolute"
        }
      });

      expect(image.range!.editAs).toBe("absolute");
    });

    it("includes editAs in model output", () => {
      const image = imageCreate(worksheet, {
        type: "image",
        imageId: "img12",
        range: {
          tl: { col: 0, row: 0 },
          br: { col: 2, row: 2 },
          editAs: "twoCell"
        }
      });

      const model = imageModel(image);
      if (model.type === "image") {
        expect(model.range.editAs).toBe("twoCell");
      }
    });
  });

  // ===========================================================================
  // Round-trip Tests
  // ===========================================================================

  describe("round-trip", () => {
    it("preserves data through model get/set cycle", () => {
      const original = imageCreate(worksheet, {
        type: "image",
        imageId: "round-trip",
        range: {
          tl: { col: 1, row: 2 },
          br: { col: 3, row: 4 },
          editAs: "oneCell"
        },
        hyperlinks: {
          hyperlink: "https://example.com",
          tooltip: "Test"
        }
      });

      const model = imageModel(original);

      const restored = imageCreate(worksheet);
      applyImageModel(restored, model as any);

      expect(restored.type).toBe(original.type);
      expect(restored.imageId).toBe(original.imageId);
      expect(restored.range!.editAs).toBe(original.range!.editAs);
    });
  });

  // ===========================================================================
  // Clone Tests
  // ===========================================================================

  describe("clone", () => {
    it("creates a deep copy on the same worksheet", () => {
      const original = imageCreate(worksheet, {
        type: "image",
        imageId: "clone1",
        range: {
          tl: { col: 1, row: 2 },
          br: { col: 3, row: 4 },
          editAs: "oneCell"
        },
        hyperlinks: {
          hyperlink: "https://example.com",
          tooltip: "Test"
        }
      });

      const cloned = imageClone(original);

      expect(cloned).not.toBe(original);
      expect(cloned.worksheet).toBe(original.worksheet);
      expect(cloned.type).toBe(original.type);
      expect(cloned.imageId).toBe(original.imageId);
      expect(cloned.range!.editAs).toBe(original.range!.editAs);
      expect(cloned.range!.hyperlinks).toEqual(original.range!.hyperlinks);
      // Verify deep copy — mutating clone should not affect original
      expect(cloned.range!.tl).not.toBe(original.range!.tl);
      expect(cloned.range!.br).not.toBe(original.range!.br);
      expect(cloned.range!.hyperlinks).not.toBe(original.range!.hyperlinks);
    });

    it("clones to a different worksheet", () => {
      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
      const original = imageCreate(worksheet, {
        type: "image",
        imageId: "clone2",
        range: {
          tl: { col: 0, row: 0 },
          ext: { width: 200, height: 150 }
        }
      });

      const cloned = imageClone(original, worksheet2);

      expect(cloned.worksheet).toBe(worksheet2);
      expect(cloned.type).toBe("image");
      expect(cloned.imageId).toBe("clone2");
      expect(cloned.range!.ext).toEqual({ width: 200, height: 150 });
      expect(cloned.range!.ext).not.toBe(original.range!.ext);
      expect(cloned.range!.br).toBeUndefined();
    });

    it("clones background image (no range)", () => {
      const original = imageCreate(worksheet, {
        type: "background",
        imageId: "bg-clone"
      });

      const cloned = imageClone(original);

      expect(cloned.type).toBe("background");
      expect(cloned.imageId).toBe("bg-clone");
      expect(cloned.range).toBeUndefined();
    });

    it("clone model round-trips correctly", () => {
      const original = imageCreate(worksheet, {
        type: "image",
        imageId: "rt-clone",
        range: {
          tl: { nativeCol: 2, nativeRow: 3, nativeColOff: 50000, nativeRowOff: 60000 },
          br: { nativeCol: 5, nativeRow: 8, nativeColOff: 0, nativeRowOff: 0 },
          editAs: "twoCell"
        },
        hyperlinks: {
          hyperlink: "https://clone-test.com",
          tooltip: "clone tooltip"
        }
      });

      const cloned = imageClone(original);
      const originalModel = imageModel(original);
      const clonedModel = imageModel(cloned);

      expect(clonedModel).toEqual(originalModel);
    });
  });

  // ===========================================================================
  // Object range with string tl/br: must produce the same 0-based nativeCol/Row
  // as the equivalent string-only range. Regression for off-by-one bug where
  // `{ tl: "A1", br: "C3" }` resolved to nativeCol=1/nativeRow=1.
  // ===========================================================================
  describe("object range with string tl/br", () => {
    it("object {tl: 'B2', br: 'D4'} matches string range 'B2:D4'", () => {
      const stringImage = imageCreate(worksheet, {
        type: "image",
        imageId: "str-range",
        range: "B2:D4"
      });
      const objectImage = imageCreate(worksheet, {
        type: "image",
        imageId: "obj-range",
        range: { tl: "B2", br: "D4" }
      });

      expect(objectImage.range!.tl.nativeCol).toBe(stringImage.range!.tl.nativeCol);
      expect(objectImage.range!.tl.nativeRow).toBe(stringImage.range!.tl.nativeRow);
      expect(objectImage.range!.br!.nativeCol).toBe(stringImage.range!.br!.nativeCol);
      expect(objectImage.range!.br!.nativeRow).toBe(stringImage.range!.br!.nativeRow);
    });

    it("object {tl: 'A1'} resolves to 0-based native col/row", () => {
      const image = imageCreate(worksheet, {
        type: "image",
        imageId: "tl-string",
        range: { tl: "A1", ext: { width: 100, height: 100 } }
      });
      expect(image.range!.tl.nativeCol).toBe(0);
      expect(image.range!.tl.nativeRow).toBe(0);
    });

    it("object {tl: {col, row}} keeps 0-based native col/row", () => {
      const image = imageCreate(worksheet, {
        type: "image",
        imageId: "tl-object",
        range: { tl: { col: 0, row: 0 }, ext: { width: 100, height: 100 } }
      });
      expect(image.range!.tl.nativeCol).toBe(0);
      expect(image.range!.tl.nativeRow).toBe(0);
    });
  });
});
