/**
 * Excel Note Class Unit Tests
 *
 * Tests for the Note class (cell comments):
 * - Construction with string and object inputs
 * - Model getter with default config merging
 * - Model setter with text extraction
 * - Static factory method
 */

import { Note } from "@excel/note";
import { describe, it, expect } from "vitest";

// =============================================================================
// Constructor Tests
// =============================================================================

describe("Note", () => {
  describe("constructor", () => {
    it("creates note with string", () => {
      const note = new Note("Hello World");
      expect(note.note).toBe("Hello World");
    });

    it("creates note with object config", () => {
      const config = {
        texts: [{ text: "First line" }, { text: "Second line" }],
        editAs: "oneCells"
      };
      const note = new Note(config);
      expect(note.note).toEqual(config);
    });

    it("creates note without argument", () => {
      const note = new Note();
      expect(note.note).toBeUndefined();
    });
  });

  // ===========================================================================
  // Model Getter Tests
  // ===========================================================================

  describe("model getter", () => {
    it("converts string to model with texts array", () => {
      const note = new Note("Simple note");
      const model = note.model;

      expect(model.type).toBe("note");
      expect(model.note.texts).toEqual([{ text: "Simple note" }]);
    });

    it("merges default configs for string input", () => {
      const note = new Note("Test");
      const model = note.model;

      // Should have default margins
      expect(model.note.margins).toBeDefined();
      expect(model.note.margins!.insetmode).toBe("auto");
      expect(model.note.margins!.inset).toEqual([0.13, 0.13, 0.25, 0.25]);

      // Should have default protection
      expect(model.note.protection).toBeDefined();
      expect(model.note.protection!.locked).toBe("True");
      expect(model.note.protection!.lockText).toBe("True");

      // Should have default editAs
      expect(model.note.editAs).toBe("absolute");
    });

    it("merges default configs for object input", () => {
      const note = new Note({
        texts: [{ text: "Custom" }]
      });
      const model = note.model;

      // Should merge defaults
      expect(model.note.margins).toBeDefined();
      expect(model.note.protection).toBeDefined();
      expect(model.note.editAs).toBe("absolute");
    });

    it("preserves custom config over defaults", () => {
      const note = new Note({
        texts: [{ text: "Custom" }],
        editAs: "oneCell",
        protection: {
          locked: "False"
        }
      });
      const model = note.model;

      expect(model.note.editAs).toBe("oneCell");
      expect(model.note.protection!.locked).toBe("False");
    });
  });

  // ===========================================================================
  // Model Setter Tests
  // ===========================================================================

  describe("model setter", () => {
    it("extracts simple text from single-text model", () => {
      const note = new Note();
      note.model = {
        type: "note",
        note: {
          texts: [{ text: "Extracted text" }]
        }
      };

      expect(note.note).toBe("Extracted text");
    });

    it("preserves full config for complex texts", () => {
      const note = new Note();
      const complexNote = {
        texts: [{ text: "Bold", font: { bold: true } }, { text: " and normal" }],
        editAs: "absolute"
      };
      note.model = {
        type: "note",
        note: complexNote
      };

      expect(note.note).toEqual(complexNote);
    });

    it("preserves full config for multi-text notes", () => {
      const note = new Note();
      const multiText = {
        texts: [{ text: "Line 1" }, { text: "Line 2" }]
      };
      note.model = {
        type: "note",
        note: multiText
      };

      expect(note.note).toEqual(multiText);
    });
  });

  // ===========================================================================
  // Static fromModel Tests
  // ===========================================================================

  describe("fromModel", () => {
    it("creates Note instance from model", () => {
      const model = {
        type: "note",
        note: {
          texts: [{ text: "From model" }]
        }
      };

      const note = Note.fromModel(model);

      expect(note).toBeInstanceOf(Note);
      expect(note.note).toBe("From model");
    });

    it("handles complex model", () => {
      const model = {
        type: "note",
        note: {
          texts: [{ text: "Complex", font: { size: 14 } }],
          editAs: "oneCell"
        }
      };

      const note = Note.fromModel(model);

      expect(note).toBeInstanceOf(Note);
      expect(note.note).toEqual(model.note);
    });
  });

  // ===========================================================================
  // DEFAULT_CONFIGS Tests
  // ===========================================================================

  describe("DEFAULT_CONFIGS", () => {
    it("has correct default values", () => {
      expect(Note.DEFAULT_CONFIGS).toEqual({
        type: "note",
        note: {
          margins: {
            insetmode: "auto",
            inset: [0.13, 0.13, 0.25, 0.25]
          },
          protection: {
            locked: "True",
            lockText: "True"
          },
          editAs: "absolute"
        }
      });
    });
  });

  // ===========================================================================
  // Round-trip Tests
  // ===========================================================================

  describe("round-trip", () => {
    it("preserves simple string through model cycle", () => {
      const original = new Note("Test note");
      const model = original.model;

      const restored = new Note();
      restored.model = model;

      expect(restored.note).toBe("Test note");
    });

    it("preserves complex config through model cycle", () => {
      const config = {
        texts: [{ text: "Rich text", font: { bold: true, italic: true } }],
        editAs: "oneCell",
        protection: { locked: "False", lockText: "False" }
      };
      const original = new Note(config);
      const model = original.model;

      const restored = new Note();
      restored.model = model;

      // Complex config should be preserved
      expect(typeof restored.note).toBe("object");
      const noteConfig = restored.note as any;
      expect(noteConfig.texts[0].font.bold).toBe(true);
    });
  });

  // ===========================================================================
  // Edge Cases
  // ===========================================================================

  describe("edge cases", () => {
    it("handles empty string", () => {
      const note = new Note("");
      const model = note.model;

      expect(model.note.texts).toEqual([{ text: "" }]);
    });

    it("handles multiline text", () => {
      const note = new Note("Line 1\nLine 2\nLine 3");
      const model = note.model;

      expect(model.note.texts).toEqual([{ text: "Line 1\nLine 2\nLine 3" }]);
    });

    it("handles unicode text", () => {
      const note = new Note("你好世界 🎉");
      const model = note.model;

      expect(model.note.texts).toEqual([{ text: "你好世界 🎉" }]);
    });

    it("handles whitespace-only text", () => {
      const note = new Note("   ");
      const model = note.model;

      expect(model.note.texts).toEqual([{ text: "   " }]);
    });
  });
});
