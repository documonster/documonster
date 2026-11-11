/**
 * The size of a prolog context — `{ root, entities, encoding?, attributes? }` — in the bytes of its JSON form,
 * which is what `maxContextBytes` bounds: it is what is sent to wherever a later range is read.
 * The parser counts the context it gathers as it goes; one a caller hands in is measured whole.
 */

import type { XmlAttributeDeclaration, XmlPrologContext } from "@xml/types";
import { utf8Length } from "@xml/utf8-decoder";

/** `{"root":` `,"entities":{` `}}` */
const FRAME_BYTES = 8 + 13 + 2;
/** `,"attributes":{` `}` */
const ATTRIBUTES_FRAME_BYTES = 15 + 1;
const ENCODING_FIELD_BYTES = ',"encoding":"ascii"'.length;

/** The UTF-8 bytes `JSON.stringify(text)` takes. */
function jsonStringBytes(text: string): number {
  const json = JSON.stringify(text);
  return utf8Length(json, 0, json.length);
}

/** The UTF-8 bytes `JSON.stringify(declaration)` takes. */
function declarationBytes(declaration: XmlAttributeDeclaration): number {
  const json = JSON.stringify(declaration);
  return utf8Length(json, 0, json.length);
}

/** A context's size as its entities and attribute declarations arrive, before its root is known. */
export class ContextSize {
  /** The bytes of the entities' `"name":"value"` entries, each with a separating comma. */
  private entries = 0;
  /** The bytes of the `"element":{…}` entries of `attributes`, each with a separating comma. */
  private attributeEntries = 0;
  /** How many attributes each element has so far: its first opens its `{…}`. */
  private readonly elements = new Map<string, number>();

  /** Count one more entity. */
  add(name: string, value: string): void {
    this.entries += jsonStringBytes(name) + 1 + jsonStringBytes(value) + 1;
  }

  /** Count one more attribute declaration. */
  addAttribute(element: string, name: string, declaration: XmlAttributeDeclaration): void {
    const count = this.elements.get(element) ?? 0;
    if (count === 0) {
      // `"element":{` `}` and the comma before the next element.
      this.attributeEntries += jsonStringBytes(element) + 1 + 2 + 1;
    } else {
      // The comma between two attributes of one element.
      this.attributeEntries++;
    }
    this.elements.set(element, count + 1);
    this.attributeEntries += jsonStringBytes(name) + 1 + declarationBytes(declaration);
  }

  /** Whether the context carries `encoding`: the document declared ASCII. */
  ascii = false;

  /** The size of the whole context with root `root` — or, before it is known, the least it can be. */
  total(root?: string): number {
    const rootBytes = root === undefined ? 3 : jsonStringBytes(root);
    const attributes =
      this.attributeEntries === 0 ? 0 : ATTRIBUTES_FRAME_BYTES + this.attributeEntries - 1;
    const encoding = this.ascii ? ENCODING_FIELD_BYTES : 0;
    return FRAME_BYTES + rootBytes + Math.max(0, this.entries - 1) + attributes + encoding;
  }
}

/**
 * The size of a context a caller hands in, in the bytes of its JSON form — measured from that form:
 * it is checked once, where a parser is made.
 */
export function contextBytes(context: XmlPrologContext): number {
  const json = JSON.stringify(context);
  return utf8Length(json, 0, json.length);
}
