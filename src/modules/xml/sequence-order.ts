/**
 * Child order checking for vocabularies whose content models are `xsd:sequence`s — OOXML, where Office refuses
 * or "repairs" a part with children out of order.
 *
 * The vocabulary is a module's: a table of content models flattened into ordered positions (generated from the
 * schemas by `scripts/extract-ooxml-sequences.ts`), and a tree saying where each model applies. This file is only
 * the walk, so that the workbook and document validators rank children the same way rather than each with a
 * walker of its own.
 *
 * A child is matched to its model by namespace and local name — an element of another namespace that shares a name
 * with a standard one is not that element ({@link OOXML_SEQUENCE_NAMESPACES}). The namespace is the one the model
 * declares the child in: its own (`xdr:spPr` in an `xdr:` model, though its content is `a:CT_ShapeProperties`), or
 * the one a table entry names (`a:graphic` in `wp:CT_Anchor`). An element whose prefix is not declared is matched by
 * local name, as a fragment without its declarations would otherwise match nothing. A child the model does
 * not name is skipped, not reported: an extension element may appear anywhere this check
 * cannot see. `mc:AlternateContent` is unwrapped to its first `mc:Choice`, whose children stand in the wrapper's
 * place (a wrapper nested in that choice is unwrapped the same way). Every other branch — a further `mc:Choice`, and
 * the `mc:Fallback` a consumer that understands none of them reads — is the same place's other alternative, so its
 * children are checked as well: ranked where they stand, among the siblings around the wrapper as a consumer reading
 * that branch sees them, and descended into — never against another alternative's.
 */

import { OOXML_SEQUENCE_NAMESPACES } from "@xml/ooxml-namespaces";
import type { XmlElement } from "@xml/types";

/** Where content models apply: `type` ranks this element's children, `children` descends by local name. */
export interface SequenceNode<Type extends string> {
  readonly type?: Type;
  readonly children?: Readonly<Record<string, SequenceNode<Type>>>;
}

/** A child found after one the model places later. */
export interface SequenceViolation<Type extends string> {
  /** Local names from the root to the container. */
  readonly trail: readonly string[];
  readonly child: string;
  readonly after: string;
  readonly type: Type;
}

/** Content models as ordered positions; the names in one position may appear in any order among themselves. */
export type SequenceTable<Type extends string> = Readonly<
  Record<Type, readonly (readonly string[])[]>
>;

/**
 * A content model's positions without the names it allows at two of them. Such a name cannot be ranked by itself —
 * `CT_Tbl` takes a `bookmarkStart` before its `tblPr` and between its rows; XLSB's `SLICERSEX` and
 * `TABLESLICERSEX` both open with `BrtBeginSlicersEx` — so it is left out: unchecked, never misreported, and every
 * other name keeps its position. Used by the OOXML table generator and the XLSB grammar alike.
 */
export function withoutAmbiguousNames(slots: readonly (readonly string[])[]): string[][] {
  const positions = new Map<string, Set<number>>();
  slots.forEach((slot, i) => {
    for (const name of slot) {
      positions.set(name, (positions.get(name) ?? new Set()).add(i));
    }
  });
  return slots
    .map(slot => slot.filter(name => positions.get(name)!.size === 1))
    .filter(slot => slot.length > 0);
}

/**
 * A content model's positions by child local name, and the prefix of each child declared in a namespace other than
 * its type's (`graphic` → `a` in `wp:CT_Anchor`); every other child is of its type's.
 */
export type ModelRanks = ReadonlyMap<string, number> & {
  readonly declaredIn: ReadonlyMap<string, string>;
};

/** Each content model's {@link ModelRanks}. */
export type SequenceRanks<Type extends string> = ReadonlyMap<Type, ModelRanks>;

/**
 * A table's positions by child local name, built once per table. A name the table prefixes (`a:graphic`) is ranked
 * by its local name and its prefix recorded as where it is declared.
 */
export function sequenceRanks<Type extends string>(
  table: SequenceTable<Type>
): SequenceRanks<Type> {
  return new Map(
    (Object.keys(table) as Type[]).map(type => {
      const positions = new Map<string, number>();
      const declaredIn = new Map<string, string>();
      table[type].forEach((slot, i) => {
        for (const name of slot) {
          const colon = name.indexOf(":");
          const local = name.slice(colon + 1);
          // Ranked by local name: two children of one model sharing it would overwrite each other's position.
          if (positions.has(local)) {
            throw new Error(`${type}: ${local} is named twice`);
          }
          positions.set(local, i);
          if (colon >= 0) {
            declaredIn.set(local, name.slice(0, colon));
          }
        }
      });
      return [type, Object.assign(positions, { declaredIn })];
    })
  );
}

/**
 * Walk `root` by `tree` and report every child out of its model's order. `stop` is asked before each element, so
 * a caller with a problem cap can end the walk.
 */
export function checkSequenceOrder<Type extends string>(
  root: XmlElement,
  tree: SequenceNode<Type>,
  ranks: SequenceRanks<Type>,
  report: (violation: SequenceViolation<Type>) => void,
  stop: () => boolean = () => false
): void {
  const { declared, isMc } = namespaceTests(root, ranks);
  walk(root, tree, [localNameOf(root.name)]);

  function walk(element: XmlElement, node: SequenceNode<Type>, trail: string[]): void {
    forEachAlternative(
      element,
      isMc,
      trail,
      stop,
      (children, at) => {
        if (node.type) {
          rankChildren(children, node.type, at, ranks, report, declared);
        }
      },
      (children, at) => {
        if (!node.children) {
          return;
        }
        for (const child of children) {
          if (stop()) {
            return;
          }
          const local = localNameOf(child.name);
          const next = Object.hasOwn(node.children, local) ? node.children[local] : undefined;
          if (next && declared(child, node.type, next.type)) {
            at.push(local);
            walk(child, next, at);
            at.pop();
          }
        }
      }
    );
  }
}

/** The order a type requires, for a message: positions joined by `, `, the members of one by `|`. */
export function describeSequence<Type extends string>(
  table: SequenceTable<Type>,
  type: Type
): string {
  return table[type].map(slot => slot.join("|")).join(", ");
}

/**
 * {@link checkSequenceOrder} for a vocabulary whose content models are known by element name rather than by path:
 * every element in the tree is visited, and `resolve` names the model of one by its local name and its parent's.
 * A chart is such a vocabulary — `c:dLbls` is `CT_DLbls` wherever it sits — and so its rules are a list
 * (`@xml/ooxml-chart-order`) rather than a tree.
 */
export function checkSequenceOrderByName<Type extends string>(
  root: XmlElement,
  resolve: (local: string, parentLocal: string | undefined) => Type | undefined,
  ranks: SequenceRanks<Type>,
  report: (violation: SequenceViolation<Type>) => void,
  stop: () => boolean = () => false
): void {
  const { declared, isMc } = namespaceTests(root, ranks);
  visit(root, undefined, undefined, []);

  function visit(
    element: XmlElement,
    parentLocal: string | undefined,
    parentType: Type | undefined,
    trail: string[]
  ): void {
    const local = localNameOf(element.name);
    trail.push(local);
    const resolved = resolve(local, parentLocal);
    const type = declared(element, parentType, resolved) ? resolved : undefined;
    forEachAlternative(
      element,
      isMc,
      trail,
      stop,
      (children, at) => {
        if (type !== undefined) {
          rankChildren(children, type, at, ranks, report, declared);
        }
      },
      (children, at) => {
        for (const child of children) {
          if (stop()) {
            return;
          }
          visit(child, local, type, at);
        }
      }
    );
    trail.pop();
  }
}

/** Report each of `children` found after one its model places later. */
function rankChildren<Type extends string>(
  children: readonly XmlElement[],
  type: Type,
  trail: readonly string[],
  ranks: SequenceRanks<Type>,
  report: (violation: SequenceViolation<Type>) => void,
  declared: DeclarationTest
): void {
  const positions = ranks.get(type)!;
  let maxRank = -1;
  let maxName = "";
  for (const child of children) {
    const local = localNameOf(child.name);
    const position = declared(child, type, undefined) ? positions.get(local) : undefined;
    if (position === undefined) {
      continue;
    }
    if (position < maxRank) {
      report({ trail: [...trail], child: local, after: maxName, type });
    } else if (position > maxRank) {
      maxRank = position;
      maxName = local;
    }
  }
}

/**
 * Whether `element` is the one a step names, with `parentType` the model of the element it is in and `type` its own:
 * of the namespace `parentType` declares it in, when that model names it; otherwise — reached through a wildcard, as
 * `w14:checkbox` in `w:sdtPr` and `wps:wsp` in `a:graphicData` are — of its own type's, as its schema declares it.
 * A step with no model on either side, a prefix naming no known namespace, and an element whose prefix is not
 * declared, are not held to one.
 */
type DeclarationTest = (
  element: XmlElement,
  parentType: string | undefined,
  type: string | undefined
) => boolean;

/** The Markup Compatibility namespace, whose `AlternateContent`, `Choice` and `Fallback` the walk steps through. */
const MC_NAMESPACE = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/** Whether `element` is the Markup Compatibility element `local` (`AlternateContent`, `Choice`, `Fallback`). */
type McTest = (element: XmlElement, local: string) => boolean;

/**
 * The {@link DeclarationTest} and {@link McTest} for the elements under `root`. Each element's namespace is resolved
 * once, from the declarations in scope — those on an `mc:Choice` included, which the walk steps through. An element
 * whose prefix is not declared is matched by local name, as everywhere here.
 */
function namespaceTests(
  root: XmlElement,
  ranks: SequenceRanks<string>
): { declared: DeclarationTest; isMc: McTest } {
  const namespaces = new Map<XmlElement, string | undefined>();
  const resolve = (element: XmlElement, scope: ReadonlyMap<string, string>): void => {
    let own = scope;
    for (const [name, value] of Object.entries(element.attributes)) {
      if (name === "xmlns" || name.startsWith("xmlns:")) {
        if (own === scope) {
          own = new Map(scope);
        }
        (own as Map<string, string>).set(name === "xmlns" ? "" : name.slice(6), value);
      }
    }
    const colon = element.name.indexOf(":");
    namespaces.set(element, own.get(colon < 0 ? "" : element.name.slice(0, colon)));
    for (const child of element.children) {
      if (child.type === "element") {
        resolve(child, own);
      }
    }
  };
  resolve(root, new Map());
  const isOf = (element: XmlElement, prefix: string): boolean => {
    const allowed = OOXML_SEQUENCE_NAMESPACES[prefix];
    const namespace = namespaces.get(element);
    return allowed === undefined || namespace === undefined || allowed.includes(namespace);
  };
  const prefixOf = (type: string): string => type.slice(0, type.indexOf(":"));
  return {
    declared: (element, parentType, type) => {
      const local = localNameOf(element.name);
      const parent = parentType === undefined ? undefined : ranks.get(parentType);
      if (parent?.has(local)) {
        return isOf(element, parent.declaredIn.get(local) ?? prefixOf(parentType!));
      }
      return type === undefined || isOf(element, prefixOf(type));
    },
    isMc: (element, local) => {
      if (localNameOf(element.name) !== local) {
        return false;
      }
      const namespace = namespaces.get(element);
      return namespace === undefined || namespace === MC_NAMESPACE;
    }
  };
}

function localNameOf(name: string): string {
  return name.slice(name.lastIndexOf(":") + 1);
}

/**
 * Rank and descend into `element`'s children as {@link elementChildren} gives them, then do the same for every other
 * alternative among them ({@link alternatives}): rank the children as a consumer reading that branch sees them — the
 * siblings around the wrapper, with the branch in its place ({@link childrenWith}) — and descend into the branch's own.
 * Stops between alternatives once `stop` says so.
 */
function forEachAlternative(
  element: XmlElement,
  isMc: McTest,
  trail: string[],
  stop: () => boolean,
  rank: (children: XmlElement[], trail: string[]) => void,
  descend: (children: XmlElement[], trail: string[]) => void
): void {
  const children = [...elementChildren(element, isMc, new Map())];
  rank(children, trail);
  descend(children, trail);
  forEachBranch(element, element, new Map(), isMc, trail, stop, rank, descend);
}

/**
 * The alternatives under `scope`, each ranked among `context`'s children as a consumer reading it sees them: with
 * the branches already taken to reach `scope` (`taken`, wrapper → branch) in their wrappers' places, and the new one
 * in its own. A branch's nested wrappers are alternatives too, ranked in that same whole context — not in the branch
 * alone, where the siblings around its outer wrapper would be out of sight.
 */
function forEachBranch(
  context: XmlElement,
  scope: XmlElement,
  taken: ReadonlyMap<XmlElement, XmlElement>,
  isMc: McTest,
  trail: string[],
  stop: () => boolean,
  rank: (children: XmlElement[], trail: string[]) => void,
  descend: (children: XmlElement[], trail: string[]) => void
): void {
  for (const [via, wrapper, branch] of alternatives(scope, isMc)) {
    if (stop()) {
      return;
    }
    const reading = new Map(taken).set(wrapper, branch);
    trail.push("AlternateContent", via);
    rank([...elementChildren(context, isMc, reading)], trail);
    descend([...elementChildren(branch, isMc, new Map())], trail);
    forEachBranch(context, branch, reading, isMc, trail, stop, rank, descend);
    trail.length -= 2;
  }
}

/**
 * The `mc:` branches other than the first choice, with the wrapper each belongs to: those of the wrappers among
 * `element`'s children, and of the wrappers in their first choices.
 */
function* alternatives(
  element: XmlElement,
  isMc: McTest
): Generator<[string, XmlElement, XmlElement]> {
  for (const wrapper of element.children) {
    if (wrapper.type !== "element" || !isMc(wrapper, "AlternateContent")) {
      continue;
    }
    let first = true;
    for (const branch of wrapper.children) {
      if (branch.type !== "element") {
        continue;
      }
      if (isMc(branch, "Choice") && first) {
        first = false;
        yield* alternatives(branch, isMc);
      } else if (isMc(branch, "Choice") || isMc(branch, "Fallback")) {
        yield [localNameOf(branch.name), wrapper, branch];
      }
    }
  }
}

/**
 * The element children of `element` as a consumer reads them: each `mc:AlternateContent` replaced by the children of
 * the branch `taken` names for it, or else of its first `mc:Choice` — themselves unwrapped the same way.
 */
function* elementChildren(
  element: XmlElement,
  isMc: McTest,
  taken: ReadonlyMap<XmlElement, XmlElement>
): Generator<XmlElement> {
  for (const child of element.children) {
    if (child.type !== "element") {
      continue;
    }
    if (!isMc(child, "AlternateContent")) {
      yield child;
      continue;
    }
    const branch =
      taken.get(child) ??
      child.children.find((b): b is XmlElement => b.type === "element" && isMc(b, "Choice"));
    if (branch) {
      yield* elementChildren(branch, isMc, taken);
    }
  }
}
