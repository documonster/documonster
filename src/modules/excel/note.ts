import type { NoteConfig, NoteModel } from "@excel/cell";
import { deepMerge } from "@utils/object";

/**
 * Plain-data representation of a cell comment (note).
 *
 * De-classed domain model: a bare record carrying the authored note payload
 * plus optional author, with flat helpers ({@link noteModel} /
 * {@link noteFromModel}) replacing the former getter/setter + static factory.
 */
export interface NoteData {
  note: string | NoteConfig | undefined;
  author: string | undefined;
}

/** Default geometry/protection applied to every cell comment on serialization. */
const DEFAULT_CONFIGS: NoteModel = {
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
  },
  type: "note"
};

/** Create a note record from an authored payload (string or full config). */
export function noteCreate(note?: string | NoteConfig, author?: string): NoteData {
  return { note, author };
}

/**
 * Structural guard: is `value` a {@link NoteData} record (vs a raw
 * {@link NoteConfig})? A note record always carries the `note`/`author` keys
 * (either may be `undefined`), whereas a `NoteConfig` carries `texts`/geometry.
 */
export function isNoteData(value: unknown): value is NoteData {
  return (
    typeof value === "object" &&
    value !== null &&
    "note" in value &&
    "author" in value &&
    !("texts" in value)
  );
}

/** Serialize a note record to its persisted {@link NoteModel}. */
export function noteModel(n: NoteData): NoteModel {
  let value: NoteModel;
  switch (typeof n.note) {
    case "string":
      value = {
        type: "note",
        note: {
          texts: [{ text: n.note }]
        }
      };
      break;
    default:
      value = {
        type: "note",
        note: n.note ?? {}
      };
      break;
  }
  // Suitable for all cell comments
  const result = deepMerge<NoteModel>({}, DEFAULT_CONFIGS, value);
  if (n.author !== undefined) {
    result.author = n.author;
  }
  return result;
}

/** Build a note record from a persisted {@link NoteModel}. */
export function noteFromModel(model: NoteModel): NoteData {
  const { note } = model;
  const { texts } = note;
  // A single, plain text run with no extra box geometry can be flattened
  // back to a simple string. Custom width/height must keep the full config
  // so the sizing survives the model round-trip.
  const hasCustomSize = note.width !== undefined || note.height !== undefined;
  const payload =
    texts && texts.length === 1 && Object.keys(texts[0]).length === 1 && !hasCustomSize
      ? texts[0].text
      : note;
  return { note: payload, author: model.author };
}
