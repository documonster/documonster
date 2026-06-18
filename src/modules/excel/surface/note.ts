/**
 * `Note` namespace surface — cell note / comment values.
 *
 * `import { Note } from "@cj-tech-master/excelts/excel"` →
 *   `Note.create("text", "author")`, `Note.isNote(v)`, `Note.model(n)`,
 *   `Note.fromModel(model)`.
 *
 * Attaching a note to a cell (`Cell.note` / `Cell.setNote` /
 * `Cell.setComment`) lives on the `Cell` namespace.
 */
export {
  noteCreate as create,
  isNoteData as isNote,
  noteModel as model,
  noteFromModel as fromModel
} from "@excel/note";

/** A note handle. */
export type { NoteData as Handle } from "@excel/note";
