/**
 * Excel Note Unit Tests
 *
 * Tests for cell-comment (note) helpers (de-classed domain model):
 * - Creation with string and object inputs
 * - Model serialization with default config merging (`noteModel`)
 * - Model deserialization with text extraction (`noteFromModel`)
 * - Structural guard (`isNoteData`)
 */

import { isNoteData, noteCreate, noteFromModel, noteModel } from "@excel/core/note";
import { describe, it, expect } from "vitest";

// =============================================================================
// noteCreate Tests
// =============================================================================

describe("Note", () => {
  describe("noteCreate", () => {
    it("creates note with string", () => {
      const note = noteCreate("Hello World");
      expect(note.note).toBe("Hello World");
    });

    it("creates note with object config", () => {
      const config = {
        texts: [{ text: "First line" }, { text: "Second line" }],
        editAs: "oneCells"
      };
      const note = noteCreate(config);
      expect(note.note).toEqual(config);
    });

    it("creates note without argument", () => {
      const note = noteCreate();
      expect(note.note).toBeUndefined();
    });
  });

  // ===========================================================================
  // noteModel (serialization) Tests
  // ===========================================================================

  describe("noteModel", () => {
    it("converts string to model with texts array", () => {
      const model = noteModel(noteCreate("Simple note"));

      expect(model.type).toBe("note");
      expect(model.note.texts).toEqual([{ text: "Simple note" }]);
    });

    it("merges default configs for string input", () => {
      const model = noteModel(noteCreate("Test"));

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
      const model = noteModel(
        noteCreate({
          texts: [{ text: "Custom" }]
        })
      );

      // Should merge defaults
      expect(model.note.margins).toBeDefined();
      expect(model.note.protection).toBeDefined();
      expect(model.note.editAs).toBe("absolute");
    });

    it("preserves custom config over defaults", () => {
      const model = noteModel(
        noteCreate({
          texts: [{ text: "Custom" }],
          editAs: "oneCell",
          protection: {
            locked: "False"
          }
        })
      );

      expect(model.note.editAs).toBe("oneCell");
      expect(model.note.protection!.locked).toBe("False");
    });

    it("applies the full default config block (margins + protection + editAs)", () => {
      // Replaces the former `Note.DEFAULT_CONFIGS` static-shape assertion:
      // verify the exact defaults are materialized on a plain note.
      const model = noteModel(noteCreate("x"));
      expect(model.note.margins).toEqual({
        insetmode: "auto",
        inset: [0.13, 0.13, 0.25, 0.25]
      });
      expect(model.note.protection).toEqual({
        locked: "True",
        lockText: "True"
      });
      expect(model.note.editAs).toBe("absolute");
      expect(model.type).toBe("note");
    });
  });

  // ===========================================================================
  // noteFromModel (deserialization) Tests
  // ===========================================================================

  describe("noteFromModel", () => {
    it("extracts simple text from single-text model", () => {
      const note = noteFromModel({
        type: "note",
        note: {
          texts: [{ text: "Extracted text" }]
        }
      });

      expect(note.note).toBe("Extracted text");
    });

    it("preserves full config for complex texts", () => {
      const complexNote = {
        texts: [{ text: "Bold", font: { bold: true } }, { text: " and normal" }],
        editAs: "absolute"
      };
      const note = noteFromModel({
        type: "note",
        note: complexNote
      });

      expect(note.note).toEqual(complexNote);
    });

    it("preserves full config for multi-text notes", () => {
      const multiText = {
        texts: [{ text: "Line 1" }, { text: "Line 2" }]
      };
      const note = noteFromModel({
        type: "note",
        note: multiText
      });

      expect(note.note).toEqual(multiText);
    });

    it("creates a note record from model", () => {
      const model = {
        type: "note" as const,
        note: {
          texts: [{ text: "From model" }]
        }
      };

      const note = noteFromModel(model);

      expect(isNoteData(note)).toBe(true);
      expect(note.note).toBe("From model");
    });

    it("handles complex model", () => {
      const model = {
        type: "note" as const,
        note: {
          texts: [{ text: "Complex", font: { size: 14 } }],
          editAs: "oneCell"
        }
      };

      const note = noteFromModel(model);

      expect(isNoteData(note)).toBe(true);
      expect(note.note).toEqual(model.note);
    });
  });

  // ===========================================================================
  // isNoteData guard Tests
  // ===========================================================================

  describe("isNoteData", () => {
    it("recognizes a note record", () => {
      expect(isNoteData(noteCreate("x"))).toBe(true);
      expect(isNoteData(noteCreate())).toBe(true);
    });

    it("rejects a raw NoteConfig (has texts, not author)", () => {
      expect(isNoteData({ texts: [{ text: "x" }] })).toBe(false);
    });

    it("rejects non-objects", () => {
      expect(isNoteData(undefined)).toBe(false);
      expect(isNoteData(null)).toBe(false);
      expect(isNoteData("string")).toBe(false);
    });
  });

  // ===========================================================================
  // Author Round-trip Tests
  // ===========================================================================

  describe("author round-trip", () => {
    it("preserves author through noteModel", () => {
      const note = noteCreate("x", "Alice");
      expect(noteModel(note).author).toBe("Alice");
    });

    it("preserves author through noteFromModel", () => {
      const note = noteFromModel({
        type: "note",
        note: { texts: [{ text: "x" }] },
        author: "Bob"
      });
      expect(note.author).toBe("Bob");
    });

    it("preserves author through full model cycle", () => {
      const original = noteCreate("Hello", "Alice");
      const restored = noteFromModel(noteModel(original));
      expect(restored.author).toBe("Alice");
    });

    it("preserves empty string author", () => {
      const note = noteCreate("x", "");
      const model = noteModel(note);
      expect(model.author).toBe("");

      const restored = noteFromModel(model);
      expect(restored.author).toBe("");
    });

    it("omits author from model when undefined", () => {
      const model = noteModel(noteCreate("x"));
      expect(model.author).toBeUndefined();
      expect("author" in model).toBe(false);
    });
  });

  // ===========================================================================
  // Round-trip Tests
  // ===========================================================================

  describe("round-trip", () => {
    it("preserves simple string through model cycle", () => {
      const model = noteModel(noteCreate("Test note"));
      const restored = noteFromModel(model);

      expect(restored.note).toBe("Test note");
    });

    it("preserves complex config through model cycle", () => {
      const config = {
        texts: [{ text: "Rich text", font: { bold: true, italic: true } }],
        editAs: "oneCell",
        protection: { locked: "False", lockText: "False" }
      };
      const model = noteModel(noteCreate(config));
      const restored = noteFromModel(model);

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
      const model = noteModel(noteCreate(""));
      expect(model.note.texts).toEqual([{ text: "" }]);
    });

    it("handles multiline text", () => {
      const model = noteModel(noteCreate("Line 1\nLine 2\nLine 3"));
      expect(model.note.texts).toEqual([{ text: "Line 1\nLine 2\nLine 3" }]);
    });

    it("handles unicode text", () => {
      const model = noteModel(noteCreate("你好世界 🎉"));
      expect(model.note.texts).toEqual([{ text: "你好世界 🎉" }]);
    });

    it("handles whitespace-only text", () => {
      const model = noteModel(noteCreate("   "));
      expect(model.note.texts).toEqual([{ text: "   " }]);
    });
  });

  // ===========================================================================
  // Comment box size (width / height)
  // ===========================================================================

  describe("comment box size", () => {
    it("preserves width/height through noteModel", () => {
      const model = noteModel(noteCreate({ texts: [{ text: "Sized" }], width: 200, height: 120 }));

      expect(model.note.width).toBe(200);
      expect(model.note.height).toBe(120);
    });

    it("preserves width/height through a full model cycle", () => {
      const model = noteModel(noteCreate({ texts: [{ text: "Sized" }], width: 150.5, height: 90 }));
      const restored = noteFromModel(model);

      const config = restored.note as any;
      expect(config.width).toBe(150.5);
      expect(config.height).toBe(90);
    });

    it("leaves width/height undefined when not configured", () => {
      const model = noteModel(noteCreate("No size"));

      expect(model.note.width).toBeUndefined();
      expect(model.note.height).toBeUndefined();
    });
  });
});
