/**
 * DOCX Writers - markup the model does not represent.
 *
 * Writes back what `@word/reader/preserved-xml` kept: the root's namespace
 * declarations and `mc:Ignorable`, and each unmodelled element at its schema
 * position among the elements the writer produces from the model.
 */

import { NS_A, NS_M, NS_W } from "@word/constants";
import { DocxRawXmlPolicyError } from "@word/errors";
import type { PreservedXml } from "@word/types";
import type { XmlSink } from "@xml/types";

type RawXmlPolicy = "preserve" | "strip" | "reject" | undefined;

const NS_SL = "http://schemas.openxmlformats.org/schemaLibrary/2006/main";

/** The URIs of the prefixes schema-order tables are written with. */
const ORDER_PREFIXES: Readonly<Record<string, string>> = {
  w: NS_W,
  m: NS_M,
  sl: NS_SL,
  a: NS_A
};

/**
 * A schema sequence, as qualified names in the canonical prefixes, turned into
 * a lookup from `uri|local` to position.
 */
export function schemaOrder(names: readonly string[]): ReadonlyMap<string, number> {
  const order = new Map<string, number>();
  names.forEach((name, index) => {
    const colon = name.indexOf(":");
    order.set(`${ORDER_PREFIXES[name.slice(0, colon)]}|${name.slice(colon + 1)}`, index);
  });
  return order;
}

/** Position of a canonical qualified name (`w:zoom`) in a sequence. */
function positionOfName(order: ReadonlyMap<string, number>, name: string): number {
  const colon = name.indexOf(":");
  return order.get(`${ORDER_PREFIXES[name.slice(0, colon)]}|${name.slice(colon + 1)}`) ?? Infinity;
}

const START_TAG = /^<([^\s/>]+)/;

/**
 * Position of a preserved element in a sequence, resolved through the
 * namespace declaration it carries, so a source that used another prefix for
 * the same namespace still lands in place. One the sequence does not name — an
 * extension element — goes last.
 */
function positionOfElement(order: ReadonlyMap<string, number>, xml: string): number {
  const name = START_TAG.exec(xml)?.[1] ?? "";
  const colon = name.indexOf(":");
  const prefix = colon === -1 ? "" : name.slice(0, colon);
  const local = name.slice(colon + 1);
  const declaration = prefix === "" ? "xmlns" : `xmlns:${prefix}`;
  const uri = new RegExp(`\\s${declaration}="([^"]*)"`).exec(
    xml.slice(0, xml.indexOf(">") + 1)
  )?.[1];
  return uri === undefined ? Infinity : (order.get(`${uri}|${local}`) ?? Infinity);
}

/** The preserved elements to write under the raw-XML policy (`label` names the field). */
export function preservedElements(
  elements: readonly string[] | undefined,
  rawXmlPolicy: RawXmlPolicy,
  label: string
): readonly string[] {
  if (!elements || elements.length === 0 || rawXmlPolicy === "strip") {
    return [];
  }
  if (rawXmlPolicy === "reject") {
    throw new DocxRawXmlPolicyError(label);
  }
  return elements;
}

/**
 * `attrs` (the writer's own root attributes) plus the source root's namespace
 * declarations it lacks and its `mc:Ignorable`, restricted to prefixes that
 * end up declared — an undeclared ignorable prefix makes the part invalid.
 */
export function withPreservedNamespaces(
  attrs: Record<string, string>,
  preserved: PreservedXml | undefined
): Record<string, string> {
  if (!preserved?.namespaces && !preserved?.ignorable) {
    return attrs;
  }
  const out = { ...attrs };
  for (const [prefix, uri] of Object.entries(preserved.namespaces ?? {})) {
    const name = prefix === "" ? "xmlns" : `xmlns:${prefix}`;
    if (out[name] === undefined) {
      out[name] = uri;
    }
  }
  if (preserved.ignorable && out["xmlns:mc"] !== undefined) {
    const declared = preserved.ignorable
      .split(/\s+/)
      .filter(prefix => prefix !== "" && out[`xmlns:${prefix}`] !== undefined);
    if (declared.length > 0) {
      out["mc:Ignorable"] = declared.join(" ");
    }
  }
  return out;
}

/**
 * Interleave preserved elements with what the writer emits in schema order:
 * `upTo(name)` writes every pending element that precedes `name` in the
 * sequence, `rest()` the remainder. The writer calls `upTo` before each
 * element it writes, which must therefore come in sequence order.
 */
export function createPreservedEmitter(
  xml: XmlSink,
  order: ReadonlyMap<string, number>,
  elements: readonly string[]
): { upTo(name: string): void; rest(): void } {
  // Stable, so elements sharing a position (or none) keep their source order.
  let pending = elements
    .map(text => ({ text, position: positionOfElement(order, text) }))
    .sort((a, b) => a.position - b.position);
  return {
    upTo(name: string): void {
      if (pending.length === 0) {
        return;
      }
      const limit = positionOfName(order, name);
      const remaining: typeof pending = [];
      for (const element of pending) {
        if (element.position < limit) {
          xml.writeRaw(element.text);
        } else {
          remaining.push(element);
        }
      }
      pending = remaining;
    },
    rest(): void {
      for (const element of pending) {
        xml.writeRaw(element.text);
      }
      pending = [];
    }
  };
}
