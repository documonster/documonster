/**
 * Compare two API records (see `dump.ts`) name by name. Used by `scripts/api-changes.ts`.
 */
import { escapeRegExp } from "../../src/utils/regexp.ts";
import type { ApiEntry, ApiRecord } from "./dump.ts";

export type Row = [key: string, entry: ApiEntry];
export type ChangedRow = [key: string, before: ApiEntry, after: ApiEntry];

export interface ApiDiff {
  removed: Row[];
  added: Row[];
  changed: ChangedRow[];
  deprecated: Row[];
  undeprecated: Row[];
  deprecatedOverload: Row[];
}

const ALIAS = /^type ([\w$]+)(?:<[^=]*>)? = ([\w$.]+)(?:<[^;]*>)?;$/;

/**
 * Compare two releases. `declared` is the same pair spelled with declared type names (see `dumpApi`): a signature that
 * differs only in the published name of a type it mentions is the same signature, so a row is reported as changed only
 * when it differs in both spellings.
 */
type Canon = (s: string | undefined, entry: string) => string | undefined;

const OPEN = "([{<";
const CLOSE = ")]}>";

/**
 * `s` with the members of every union sorted, at every depth: TypeScript prints a union in the order it met its
 * members, so `EventTarget | EventEmitter` in one build is `EventEmitter | EventTarget` in another.
 *
 * Each bracketed group is normalized first and stands in its parent as an opaque placeholder, so the parent's own
 * separators are all at its top level: a union is what follows a parameter's or a property's `name:` (or an arrow's
 * `=>`), up to the next `,` or `;`. A `>` that ends an arrow closes nothing.
 */
export function sortUnions(s: string): string {
  let i = 0;
  const level = (close: string | undefined): string => {
    const held: string[] = [];
    let flat = "";
    while (i < s.length) {
      const c = s[i]!;
      if (close !== undefined && c === close && !(c === ">" && s[i - 1] === "=")) {
        break;
      }
      i++;
      const o = OPEN.indexOf(c);
      if (o >= 0 && !(c === "<" && s[i - 2] === "=")) {
        const inner = level(CLOSE[o]);
        held.push(c + inner + (s[i] ?? ""));
        i++;
        flat += `\u0000${held.length - 1}\u0000`;
      } else if (c === '"' || c === "'" || c === "`") {
        const end = s.indexOf(c, i);
        const stop = end < 0 ? s.length : end + 1;
        held.push(c + s.slice(i, stop));
        i = stop;
        flat += `\u0000${held.length - 1}\u0000`;
      } else {
        flat += c;
      }
    }
    const out = flat
      .split(/([,;])/)
      .map(seg => {
        if (seg === "," || seg === ";") {
          return seg;
        }
        // The union starts after an arrow, else after a `name:` or a type alias's `=`.
        const arrow = seg.lastIndexOf("=> ");
        const colon = seg.indexOf(": ");
        const assign = seg.indexOf(" = ");
        const modifier = /^\s*readonly /.exec(seg)?.[0].length ?? 0;
        const at =
          arrow >= 0
            ? arrow + 3
            : colon >= 0 && !seg.slice(0, colon).includes(" | ")
              ? colon + 2
              : assign >= 0 && !seg.slice(0, assign).includes(" | ")
                ? assign + 3
                : modifier;
        const head = seg.slice(0, at);
        const tail = seg.slice(at);
        const lead = /^\s*/.exec(tail)![0];
        const trail = /\s*$/.exec(tail)![0];
        const members = tail.trim().split(" | ");
        const restored = (m: string): string =>
          m.replace(/\u0000(\d+)\u0000/g, (_, n: string) => held[Number(n)]!);
        members.sort((x, y) =>
          restored(x) < restored(y) ? -1 : restored(x) > restored(y) ? 1 : 0
        );
        return members.length > 1 ? head + lead + members.join(" | ") + trail : seg;
      })
      .join("");
    return out.replace(/\u0000(\d+)\u0000/g, (_, n: string) => held[Number(n)]!);
  };
  return level(undefined);
}

/**
 * Differences that are only printing: a type printed under its new name — read from the `type Old = New;` aliases the
 * new surface keeps as deprecated forwards, per entry, since two entries may each forward an old name to a different
 * type (`TransformResult` is `ArchiveTransformResult` in one and `CsvTransformResult` in another) — a namespace
 * printed under the name it was imported as, and the order TypeScript prints a union of string literals in.
 */
function canonFor(after: ApiRecord): Canon {
  const rename = (from: string, to: string): [RegExp, string] => [
    new RegExp("(?<![\\w$.])" + escapeRegExp(from) + "(?![\\w$])", "g"),
    to
  ];
  // Per entry, then — for a signature in one entry naming a type another entry forwards — every forward whose old
  // name means the same new name wherever it is forwarded.
  const perEntry = new Map<string, Map<string, string>>();
  const anywhere = new Map<string, Set<string>>();
  for (const [k, v] of Object.entries(after)) {
    const m = v.kind === "type" ? ALIAS.exec(v.sig) : null;
    if (m && m[1] !== m[2]) {
      const to = m[2]!.split(".").pop()!;
      const entry = k.split(" :: ")[0]!;
      (perEntry.get(entry) ?? perEntry.set(entry, new Map()).get(entry)!).set(m[1]!, to);
      (anywhere.get(m[1]!) ?? anywhere.set(m[1]!, new Set()).get(m[1]!)!).add(to);
    }
  }
  const renames = new Map<string, [RegExp, string][]>();
  const renamesOf = (entry: string): [RegExp, string][] => {
    let list = renames.get(entry);
    if (!list) {
      const own = perEntry.get(entry) ?? new Map<string, string>();
      list = [...own].map(([from, to]) => rename(from, to));
      for (const [from, tos] of anywhere) {
        if (!own.has(from) && tos.size === 1) {
          list.push(rename(from, [...tos][0]!));
        }
      }
      renames.set(entry, list);
    }
    return list;
  };
  return (s, entry) => {
    if (!s) {
      return s;
    }
    let r = s;
    for (const [re, to] of renamesOf(entry)) {
      r = r.replace(re, to);
    }
    r = r.replace(/\b(?:Stream|DrawNs|MermaidNs)\./g, "");
    return sortUnions(r);
  };
}

/**
 * Compare two releases. `declared` is the same pair spelled with declared type names (see `dumpApi`): a signature that
 * differs only in the published name of a type it mentions is the same signature, so a row is reported as changed only
 * when it differs in both spellings.
 */
export function diffApi(
  before: ApiRecord,
  after: ApiRecord,
  declared?: { before: ApiRecord; after: ApiRecord }
): ApiDiff {
  const canonPublished = canonFor(after);
  const canonDeclared = declared ? canonFor(declared.after) : canonPublished;
  const tparamsOf = (v: ApiEntry): string =>
    v.kind === "interface" ? (/^<.*?>(?= |$)/.exec(v.sig) ?? [""])[0] : (v.tparams ?? "");

  const res: ApiDiff = {
    removed: [],
    added: [],
    changed: [],
    deprecated: [],
    undeprecated: [],
    deprecatedOverload: []
  };
  for (const [k, a] of Object.entries(before)) {
    const b = after[k];
    if (!b) {
      res.removed.push([k, a]);
      continue;
    }
    const entry = k.split(" :: ")[0]!;
    const sameIn = (x: ApiEntry, y: ApiEntry, canon: Canon): boolean => {
      const c = (v: string | undefined): string | undefined => canon(v, entry);
      if (x.kind === y.kind && (x.sig === y.sig || c(x.sig) === c(y.sig))) {
        return true;
      }
      // Overloads that differ only in parameter names accept the same calls.
      if (
        x.kind === y.kind &&
        x.shape !== undefined &&
        y.shape !== undefined &&
        c(x.shape) === c(y.shape)
      ) {
        return true;
      }
      // An interface that became `type X = Y` is compared field by field — each field is its own row.
      if (y.kind === "type" && ALIAS.test(y.sig) && x.kind === "interface") {
        return c(tparamsOf(x)) === c(tparamsOf(y));
      }
      // A type alias whose text changed but which resolves to the same type is the same — provided its type
      // parameters (their constraints and defaults) are too: `Box<T extends string | number>` → `Box<T extends
      // string>` resolves to the same `{ value: T }` and still breaks `Box<number>`.
      return (
        x.kind === "type" &&
        y.kind === "type" &&
        c(x.tparams) === c(y.tparams) &&
        (x.resolved === y.resolved || c(x.resolved) === c(y.resolved))
      );
    };
    const da = declared?.before[k];
    const db = declared?.after[k];
    if (!sameIn(a, b, canonPublished) && !(da && db && sameIn(da, db, canonDeclared))) {
      res.changed.push([k, a, b]);
    }
    if (!a.deprecated && b.deprecated) {
      res.deprecated.push([k, b]);
    }
    if (a.deprecated && !b.deprecated) {
      res.undeprecated.push([k, b]);
    }
    if (!a.partial && b.partial && !b.deprecated) {
      res.deprecatedOverload.push([k, b]);
    }
  }
  for (const [k, b] of Object.entries(after)) {
    if (!(k in before)) {
      res.added.push([k, b]);
      if (b.partial && !b.deprecated) {
        res.deprecatedOverload.push([k, b]);
      }
    }
  }
  // A browser row that says exactly what its Node row says is noise.
  const dedupe = <T extends [string, ...unknown[]]>(list: T[]): T[] =>
    list.filter(([k, ...v]) => {
      if (!k.includes(" [browser]")) {
        return true;
      }
      const other = list.find(([k2]) => k2 === k.replace(" [browser]", ""));
      return !other || JSON.stringify(other.slice(1)) !== JSON.stringify(v);
    });
  return {
    removed: dedupe(res.removed),
    added: dedupe(res.added),
    changed: dedupe(res.changed),
    deprecated: dedupe(res.deprecated),
    undeprecated: dedupe(res.undeprecated),
    deprecatedOverload: dedupe(res.deprecatedOverload)
  };
}
