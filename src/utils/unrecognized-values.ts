/**
 * Values a reader found in a file but could not give a typed field, kept so they are not lost.
 *
 * A reader that meets an attribute outside the vocabulary it knows — an enumeration member it has no name for, a
 * number out of the range it expects — has three choices. Putting the text in the typed field makes the field lie
 * about its type. Dropping it loses what the file said without a word, and turns a vocabulary written too narrowly
 * (a Strict-only spelling, a value a newer Office added) into silent data loss that no test notices. So the value is
 * *set aside*: the typed field is left as the reader would leave an absent attribute, and the text is kept on the
 * object that would have held it, under {@link UNRECOGNIZED_KEY}.
 *
 * A writer then puts the text back exactly as read while the field is still as the reader left it, and writes the
 * field — checked as any caller's value is — once something has set it. A set-aside value therefore round-trips
 * unchanged, never blocks a write, and never overrides an edit. Clearing the field again brings it back, since that
 * is the state the reader left; {@link discardUnrecognizedValues} is how a caller removes one. Each module lists what is set aside on the channel
 * it already has for telling a caller about a read (`Validation.validate` for Word, `readWithDiagnostics` for Excel),
 * so a vocabulary that is too narrow shows up as a report rather than as missing data.
 *
 * The entries are ordinary own properties of plain data, so a spread, `structuredClone` or JSON copy keeps them with
 * the object, and replacing the object drops them with it.
 */

/** The property a set-aside value is kept under, on the object whose field could not hold it. */
export const UNRECOGNIZED_KEY = "_unrecognized";

/** One value a reader set aside. */
export interface UnrecognizedValue {
  /** The attribute text as the file had it. */
  readonly value: string;
  /**
   * What the reader put in the field instead, when the field cannot be absent (a tab stop always has a type). The
   * kept text is written while the field still holds this, as it is while an optional field is still `undefined`.
   */
  readonly readAs?: string | number | boolean;
  /**
   * Other attributes of the same element that mean something only with this value — an underline's colour, when
   * its style is the one set aside — kept so they are written back with it.
   */
  readonly with?: Readonly<Record<string, string>>;
}

/** The set-aside values on one object, by the name of the field each would have filled. */
export type UnrecognizedValues = Readonly<Record<string, UnrecognizedValue>>;

function entries(holder: object | undefined): UnrecognizedValues | undefined {
  const values = (holder as { [UNRECOGNIZED_KEY]?: unknown } | undefined)?.[UNRECOGNIZED_KEY];
  return typeof values === "object" && values !== null ? (values as UnrecognizedValues) : undefined;
}

/** Keep `value` on `holder` for `field`, which the reader leaves absent (or at `readAs`). */
export function setAsideValue(
  holder: object,
  field: string,
  value: string,
  options: Omit<UnrecognizedValue, "value"> = {}
): void {
  const target = holder as { [UNRECOGNIZED_KEY]?: Record<string, UnrecognizedValue> };
  const kept: UnrecognizedValue = { value };
  if (options.readAs !== undefined) {
    (kept as { readAs?: unknown }).readAs = options.readAs;
  }
  if (options.with !== undefined) {
    (kept as { with?: unknown }).with = options.with;
  }
  target[UNRECOGNIZED_KEY] = { ...entries(holder), [field]: kept };
}

/**
 * What was set aside for `field`, while the field is still what the reader left there: `undefined` once something has
 * set the field, or when nothing was set aside for it.
 */
export function unrecognizedEntry(
  holder: object | undefined,
  field: string
): UnrecognizedValue | undefined {
  const kept = entries(holder)?.[field];
  if (kept === undefined) {
    return undefined;
  }
  const current = (holder as Record<string, unknown>)[field];
  return current === undefined || current === kept.readAs ? kept : undefined;
}

/** The text to write for `field` in place of the field itself — see {@link unrecognizedEntry}. */
export function unrecognizedValue(holder: object | undefined, field: string): string | undefined {
  return unrecognizedEntry(holder, field)?.value;
}

/** Every set-aside value still in effect on `holder`, by field. */
export function listUnrecognizedValues(
  holder: object | undefined
): { readonly field: string; readonly value: string }[] {
  const all = entries(holder);
  if (all === undefined) {
    return [];
  }
  const found: { field: string; value: string }[] = [];
  for (const field of Object.keys(all)) {
    const value = unrecognizedValue(holder, field);
    if (value !== undefined) {
      found.push({ field, value });
    }
  }
  return found;
}

/**
 * Every set-aside value still in effect anywhere under `root`, with the path to the field (`body[3].properties.alignment`).
 * A `Map`'s entries are visited by key (`headers["default"]…`). Binary data is not descended into, and an object
 * reached twice is visited once.
 */
export function findUnrecognizedValues(
  root: unknown,
  rootPath = ""
): { readonly path: string; readonly value: string }[] {
  const found: { path: string; value: string }[] = [];
  walkHolders(root, rootPath, (holder, path) => {
    for (const { field, value } of listUnrecognizedValues(holder)) {
      found.push({ path: path ? `${path}.${field}` : field, value });
    }
  });
  return found;
}

/** One kept value as {@link discardUnrecognizedValues} offers it to its filter: where it is, and its text. */
export interface KeptValue {
  /** The path to the field, as {@link findUnrecognizedValues} gives it: `body[3].properties.alignment`. */
  readonly path: string;
  /** The text the file had. */
  readonly value: string;
}

/**
 * Drop the values set aside under `root` — every one, or those `filter` returns `true` for — so a write emits each
 * such field as it now is (absent, or what the reader put there) rather than the text the file had. Returns how many
 * of the dropped values were still in effect.
 *
 * This is the one way to remove such a value: assigning the field and clearing it again brings the kept text back,
 * because absent is exactly the state the reader left it in. `filter` also sees a value a field has since taken over
 * from, so dropping it keeps it from coming back if that field is cleared later.
 */
export function discardUnrecognizedValues(
  root: unknown,
  filter?: (value: KeptValue) => boolean,
  rootPath = ""
): number {
  let discarded = 0;
  walkHolders(root, rootPath, (holder, path) => {
    const all = entries(holder);
    if (all === undefined) {
      return;
    }
    const remaining: Record<string, UnrecognizedValue> = {};
    for (const [field, kept] of Object.entries(all)) {
      const fieldPath = path ? `${path}.${field}` : field;
      if (filter === undefined || filter({ path: fieldPath, value: kept.value })) {
        if (unrecognizedValue(holder, field) !== undefined) {
          discarded++;
        }
      } else {
        remaining[field] = kept;
      }
    }
    if (Object.keys(remaining).length === 0) {
      delete (holder as { [UNRECOGNIZED_KEY]?: unknown })[UNRECOGNIZED_KEY];
    } else {
      (holder as { [UNRECOGNIZED_KEY]?: unknown })[UNRECOGNIZED_KEY] = remaining;
    }
  });
  return discarded;
}

/** Visit every plain object under `root` once, with its path; arrays and `Map`s are descended into, bytes are not. */
function walkHolders(
  root: unknown,
  rootPath: string,
  visit: (holder: object, path: string) => void
): void {
  const seen = new Set<object>();
  const step = (node: unknown, path: string): void => {
    if (typeof node !== "object" || node === null || seen.has(node)) {
      return;
    }
    if (ArrayBuffer.isView(node) || node instanceof ArrayBuffer) {
      return;
    }
    seen.add(node);
    if (Array.isArray(node)) {
      node.forEach((item, i) => step(item, `${path}[${i}]`));
      return;
    }
    if (node instanceof Map) {
      for (const [key, value] of node) {
        step(value, `${path}[${JSON.stringify(key)}]`);
      }
      return;
    }
    visit(node, path);
    for (const [key, child] of Object.entries(node)) {
      if (key !== UNRECOGNIZED_KEY) {
        step(child, path ? `${path}.${key}` : key);
      }
    }
  };
  step(root, rootPath);
}
