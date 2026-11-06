/**
 * What the migration-guide checks share (`scripts/verify-excelts-guide.ts`, `scripts/verify-exceljs-guide.ts`): reading
 * which names a guide mentions, which replacements its tables give, and checking those replacements against the
 * current build.
 */
import { escapeRegExp } from "../../src/utils/regexp.ts";
import type { ApiRecord } from "./dump.ts";

/** Every identifier written inside a code span or fence of `text`. */
export function mentionedNames(text: string): Set<string> {
  const names = new Set<string>();
  for (const m of text.matchAll(/```[\s\S]*?```|`[^`\n]+`/g)) {
    for (const id of m[0].matchAll(/[A-Za-z_$][\w$]*/g)) {
      names.add(id[0]);
    }
  }
  return names;
}

/** All code-span text of `text`, joined — for names that are not identifiers (`[Symbol.dispose]`). */
export function codeSpanText(text: string): string {
  return [...text.matchAll(/`[^`\n]+`/g)].map(m => m[0]).join(" ");
}

/** The `Namespace.member` references in the right-hand cells of the tables of `text`, with their line numbers. */
export function replacements(text: string): { ref: string; line: number }[] {
  const out: { ref: string; line: number }[] = [];
  text.split("\n").forEach((row, i) => {
    if (!row.startsWith("|") || /^\|\s*-/.test(row)) {
      return;
    }
    for (const cell of row.split(/(?<!\\)\|/).slice(2, -1)) {
      for (const span of cell.matchAll(/`([^`]+)`/g)) {
        for (const m of span[1]!.matchAll(/(?<![\w$.#])([A-Z][\w$]*(?:\.[A-Za-z_$][\w$]*)+)/g)) {
          out.push({ ref: m[1]!, line: i + 1 });
        }
      }
    }
  });
  return out;
}

/**
 * The problems with the replacements `text` gives: a `Namespace.member` of a namespace the current build publishes
 * that does not exist, or is deprecated — which would send a reader to a name they then have to migrate again.
 */
export function checkReplacements(text: string, now: ApiRecord): string[] {
  const live = new Map<string, boolean>();
  for (const [k, v] of Object.entries(now)) {
    const n = k.split(" :: ")[1]!.replace(" (type)", "").replace(/\?$/, "");
    live.set(n, (live.get(n) ?? false) || !v.deprecated);
  }
  const namespaces = new Set([...live.keys()].filter(n => !/[.#]/.test(n)));
  // A constant object's members are not recorded one by one; its printed type names them (`{ readonly Ref: "#REF!"; … }`).
  const constants = new Map<string, string>();
  for (const [k, v] of Object.entries(now)) {
    if (v.kind === "const" && !v.deprecated) {
      constants.set(k.split(" :: ")[1]!, v.sig);
    }
  }
  const constantMember = (ref: string): boolean => {
    const at = ref.lastIndexOf(".");
    const sig = constants.get(ref.slice(0, at));
    return (
      sig !== undefined &&
      new RegExp(`(?:^|[{;,]\\s*)(?:readonly )?${ref.slice(at + 1)}\\??:`).test(sig)
    );
  };
  const problems: string[] = [];
  for (const { ref, line } of replacements(text)) {
    if (!namespaces.has(ref.split(".")[0]!) || live.get(ref) === true || constantMember(ref)) {
      continue;
    }
    problems.push(
      `guide line ${line}: \`${ref}\` ${live.has(ref) ? "is deprecated" : "does not exist"}`
    );
  }
  return problems;
}

/** Table text as `pnpm format` leaves it equal: without padding, and with a separator's dashes counted as one. */
export function normalized(s: string): string {
  return s.replace(/\s+/g, "").replace(/-+/g, "-");
}

/**
 * The names a guide writes an object's members under (`ws.getCell`), by the class or interface the object was. A member
 * counts as described only in a block (a paragraph, a table row, a code block) that names its owner — by one of these,
 * or by the owner's own name — so `getImage` described for a workbook does not account for a reader's `getImage`.
 */
export const OWNER_ALIASES: Readonly<Record<string, readonly string[]>> = {
  Workbook: ["wb", "workbook"],
  Worksheet: ["ws", "worksheet", "sheet"],
  Row: ["row"],
  Column: ["col", "column"],
  Cell: ["cell"],
  Range: ["range", "r"],
  Image: ["image"],
  Table: ["table", "t"],
  Note: ["note"],
  DataValidations: ["dv", "dataValidations"],
  FormCheckbox: ["checkbox", "cb"],
  DefinedNames: ["dn", "definedNames"],
  Chartsheet: ["cs", "chartsheet"],
  Anchor: ["anchor"],
  Chart: ["chart"],
  PivotTable: ["pt", "pivot"],
  WorkbookWriter: ["writer", "wb"],
  WorksheetWriter: ["writer", "ws"],
  WorkbookReader: ["reader"],
  WorksheetReader: ["reader", "worksheet"],
  Xlsx: ["xlsx"],
  Csv: ["csv"]
};

/** The paragraphs, table rows and code blocks of `text`. */
function blocks(text: string): string[] {
  const out: string[] = [];
  let current: string[] = [];
  let fence = false;
  const flush = (): void => {
    if (current.length) {
      out.push(current.join("\n"));
      current = [];
    }
  };
  for (const line of text.split("\n")) {
    if (line.startsWith("```")) {
      current.push(line);
      if (fence) {
        flush();
      }
      fence = !fence;
      continue;
    }
    if (fence) {
      current.push(line);
    } else if (line.startsWith("|")) {
      flush();
      out.push(line);
    } else if (line.trim() === "") {
      flush();
    } else {
      current.push(line);
    }
  }
  flush();
  return out;
}

/**
 * Whether `text` describes `member` of `owner`: a block that has the member in a code span or block, and names the
 * owner there too — `ws.`, `Worksheet`, or the owner's other names. An owner with no aliases here is matched by its own
 * name only, which every block about it carries.
 */
export function describesMember(text: string, owner: string, member: string): boolean {
  const names = [owner.split(".").pop()!, ...(OWNER_ALIASES[owner.split(".").pop()!] ?? [])];
  const ownerRe = new RegExp(`(?<![\\w$])(?:${names.map(escapeRegExp).join("|")})(?![\\w$])`);
  const symbol = member.startsWith("[") ? member.slice(1, -1) : undefined;
  return (blockCache.get(text) ?? blockCache.set(text, blocks(text)).get(text)!).some(block => {
    const code = [...block.matchAll(/```[\s\S]*?```|`[^`\n]+`/g)].map(m => m[0]).join(" ");
    const has = symbol
      ? code.includes(symbol)
      : new RegExp(`(?<![\\w$])${escapeRegExp(member)}(?![\\w$])`).test(code);
    return has && ownerRe.test(block);
  });
}
const blockCache = new Map<string, string[]>();
