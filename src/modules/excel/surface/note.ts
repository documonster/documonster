/**
 * `Note` namespace surface — cell note / comment values.
 *
 * `import { Note } from "documonster/excel"` →
 *   `Note.create("text", { author: "Alice" })`, `Note.isNote(v)`, `Note.getModel(n)`,
 *   `Note.fromModel(model)`.
 *
 * Attaching a note to a cell (`Cell.getNote` / `Cell.setNote` /
 * `Cell.setComment`) lives on the `Cell` namespace.
 */
import type { NoteConfig, NoteModel } from "@excel/core/cell";
import type { NoteData } from "@excel/core/note";
import { isNoteData, noteCreate, noteFromModel, noteModel } from "@excel/core/note";
import { checkNote } from "@excel/surface/argument-checks";
import { argumentChecks } from "@excel/utils/arguments";
import { isRecordArgument } from "@utils/object";

export { isNoteData as isNote } from "@excel/core/note";

/**
 * Rebuild a note from the model {@link getModel} returned.
 *
 * @throws {ExcelArgumentError} when `model` is not a note model — an object whose `note` is an object.
 */
export function fromModel(model: NoteModel): NoteData {
  const check = argumentChecks("Note.fromModel");
  check.required("model", model);
  check.record("model", model);
  check.required("model.note", (model as { note?: unknown }).note);
  check.record("model.note", (model as { note?: unknown }).note);
  return noteFromModel(model);
}

/** A copy of the note as its round-trip model, the shape {@link fromModel} takes: editing it changes nothing. */
export function getModel(n: NoteData): NoteModel {
  if (!isNoteData(n)) {
    argumentChecks("Note.getModel").fail("n", "a note (Note.create)", n);
  }
  return noteModel(n);
}

/** Options for {@link create}. */
export interface CreateNoteOptions {
  /** The note's author, shown by Excel above its text. */
  author?: string;
}

/** Create a note from plain text or a full {@link NoteConfig}. */
export function create(note?: string | NoteConfig, options?: CreateNoteOptions): NoteData;
/**
 * @deprecated Pass the author in the options object: `create(note, { author })`. Removed in 1.0.
 */
export function create(note: string | NoteConfig | undefined, author: string | undefined): NoteData;
export function create(
  note?: string | NoteConfig,
  authorOrOptions?: string | CreateNoteOptions
): NoteData {
  const author = isRecordArgument(authorOrOptions) ? authorOrOptions.author : authorOrOptions;
  checkNote("Note.create", "note", note);
  // The deprecated positional form stores its author as given, as it always did.
  if (isRecordArgument(authorOrOptions)) {
    argumentChecks("Note.create").string("options.author", author);
  }
  return noteCreate(note, author);
}

/** A note handle. */
export type { NoteData as Handle } from "@excel/core/note";

// --- deprecated names ---

/** @deprecated Use {@link getModel}. Removed in 1.0. */
export const model = getModel;
