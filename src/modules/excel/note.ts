import type { NoteConfig, NoteModel } from "@excel/cell";
import { deepMerge } from "@excel/utils/under-dash";

class Note {
  note: string | NoteConfig | undefined;
  author: string | undefined;

  static readonly DEFAULT_CONFIGS: NoteModel = {
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

  constructor(note?: string | NoteConfig, author?: string) {
    this.note = note;
    this.author = author;
  }

  get model(): NoteModel {
    let value: NoteModel;
    switch (typeof this.note) {
      case "string":
        value = {
          type: "note",
          note: {
            texts: [
              {
                text: this.note
              }
            ]
          }
        };
        break;
      default:
        value = {
          type: "note",
          note: this.note ?? {}
        };
        break;
    }
    // Suitable for all cell comments
    const result = deepMerge<NoteModel>({}, Note.DEFAULT_CONFIGS, value);
    if (this.author !== undefined) {
      result.author = this.author;
    }
    return result;
  }

  set model(value: NoteModel) {
    const { note } = value;
    const { texts } = note;
    // A single, plain text run with no extra box geometry can be flattened
    // back to a simple string. Custom width/height must keep the full config
    // so the sizing survives the model round-trip.
    const hasCustomSize = note.width !== undefined || note.height !== undefined;
    if (texts && texts.length === 1 && Object.keys(texts[0]).length === 1 && !hasCustomSize) {
      this.note = texts[0].text;
    } else {
      this.note = note;
    }
    this.author = value.author;
  }

  static fromModel(model: NoteModel): Note {
    const note = new Note();
    note.model = model;
    return note;
  }
}

export { Note };
