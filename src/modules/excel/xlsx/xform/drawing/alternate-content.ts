/**
 * The `<mc:AlternateContent>` a drawing anchor is read with and written back in.
 *
 * `mc:Choice/@Requires` lists **prefixes**, and a prefix means nothing without the declaration that binds
 * it: a consumer resolves each to a namespace and takes the Choice only if it understands all of them.
 * So the model keeps the namespaces beside the prefixes, and the writer declares every prefix it names —
 * an undeclared one is an MCE violation, and Excel rejects the drawing. The Fallback is what a consumer
 * that does not understand them shows instead; the source's own is kept and written back.
 */
import { namespacesInForce } from "@excel/xlsx/xform/fragment-namespaces";
import { canonicalNamespaceOf } from "@excel/xlsx/xform/namespace-prefixes";
import type { ParseOpenTag } from "@xml/types";

export const MCE_NAMESPACE = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/** The prefix Excel names in the `Requires` of a ChartEx anchor's Choice. */
export const CHART_EX_REQUIRES = "cx1";

export interface AlternateContentModel {
  /** The Choice's `Requires`: whitespace-separated prefixes. */
  requires: string;
  /** The namespace each prefix in {@link requires} was bound to, where the source declared it. */
  namespaces?: Record<string, string>;
  /**
   * Where the source put the wrapper: around the whole anchor (`"anchor"`, how form controls are
   * written, and the default), or in the anchor's shape slot (`"content"`, how ChartEx is).
   */
  placement?: "anchor" | "content";
  /** The source's `<mc:Fallback>` content as XML — anchors or a shape, by `placement`. */
  fallbackXml?: string;
}

/** Whether `uri` is a ChartEx namespace: `…/2014/chartex`, or one of its later revisions. */
export function isChartExNamespace(uri: string | undefined): boolean {
  return (
    uri !== undefined &&
    /^http:\/\/schemas\.microsoft\.com\/office\/drawing\/\d{4}\/(?:\d+\/\d+\/)?chartex$/.test(uri)
  );
}

/** The model of an `<mc:Choice>` start tag, its prefixes resolved where the parse can tell. */
export function readChoice(node: ParseOpenTag): AlternateContentModel | undefined {
  const requires = node.attributes?.Requires;
  if (!requires) {
    return undefined;
  }
  const inForce = namespacesInForce();
  const namespaces: Record<string, string> = {};
  for (const prefix of requires.split(/\s+/)) {
    const uri =
      prefix === "" ? undefined : (inForce?.get(prefix) ?? node.attributes[`xmlns:${prefix}`]);
    if (uri !== undefined) {
      namespaces[prefix] = uri;
    }
  }
  return Object.keys(namespaces).length > 0 ? { requires, namespaces } : { requires };
}

/** The namespace `prefix` of `model.requires` stands for: as read, or the canonical one. */
function namespaceOf(model: AlternateContentModel, prefix: string): string | undefined {
  return model.namespaces?.[prefix] ?? canonicalNamespaceOf(prefix);
}

/** Whether `model` requires a ChartEx namespace. */
export function requiresChartEx(model: AlternateContentModel | undefined): boolean {
  return (
    model !== undefined &&
    model.requires.split(/\s+/).some(prefix => isChartExNamespace(namespaceOf(model, prefix)))
  );
}

/**
 * The `<mc:Choice>` attributes for `model`: a declaration for every prefix it requires, then `Requires`.
 * `undefined` when a prefix stands for no namespace this writer can name — a Choice requiring it would
 * name an undeclared prefix.
 */
export function choiceAttributes(model: AlternateContentModel): Record<string, string> | undefined {
  const attributes: Record<string, string> = {};
  for (const prefix of model.requires.split(/\s+/)) {
    if (prefix === "") {
      continue;
    }
    const uri = namespaceOf(model, prefix);
    if (uri === undefined) {
      return undefined;
    }
    attributes[`xmlns:${prefix}`] = uri;
  }
  attributes.Requires = model.requires;
  return attributes;
}
