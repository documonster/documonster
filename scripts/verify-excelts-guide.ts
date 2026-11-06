/**
 * MIGRATION_EXCELTS_TO_DOCUMONSTER.md has to account for every name excelts published.
 *
 * The guide is written to be applied mechanically, so a name it does not mention is a name a migration stalls on —
 * and nothing else in this repository knows what excelts published. This reads the declarations of the last excelts
 * release (`@cj-tech-master/excelts@10.2.0`, fetched from npm) and of the current build, and sorts every export of
 * every subpath (node and browser) and every member of every exported class into one of:
 *
 * - **unchanged** — the same name in the corresponding `documonster` entry, not deprecated;
 * - **indexed** — the same name inside one of that entry's namespaces (`WorksheetWriter` → `Stream.WorksheetWriter`),
 *   or a name `documonster` still exports but has deprecated, whose deprecation names the replacement. These are
 *   written into the guide's generated name index (§12), between its `excelts-index` markers;
 * - **described** — anything else, which the hand-written sections of the guide must mention in a code span.
 *
 * It fails when a name is none of these, when the generated index is out of date, and when a replacement the guide
 * gives — a `Namespace.member` in a table's right-hand column — does not exist or is deprecated, which would send a
 * reader to a name they then have to migrate again.
 *
 * "Mentioned" is deliberately loose — it proves the guide addresses a name, not that it addresses it correctly — and
 * the replacement check is what keeps the answers honest.
 *
 * It needs `dist/types` (`pnpm build:esm`) and network access, so it is a command to run when the guide or the API
 * changes rather than part of `pnpm check`.
 *
 * Usage:
 *   node scripts/verify-excelts-guide.ts           # check
 *   node scripts/verify-excelts-guide.ts --write   # regenerate the name index, then check
 */
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { diffApi } from "./api-changes/diff.ts";
import { dumpApi } from "./api-changes/dump.ts";
import type { ApiEntry, ApiRecord } from "./api-changes/dump.ts";
import {
  checkReplacements,
  describesMember,
  mentionedNames,
  normalized
} from "./api-changes/guide-check.ts";
import { fetchPackage } from "./api-changes/package.ts";
import { sigCell } from "./api-changes/sig-cell.ts";

const ROOT = path.resolve(import.meta.dirname, "..");
const GUIDE = path.join(ROOT, "MIGRATION_EXCELTS_TO_DOCUMONSTER.md");
const TYPE_CHANGES = path.join(ROOT, "docs/api-changes-from-excelts-10.2.0.md");
const EXCELTS = "@cj-tech-master/excelts@10.2.0";
const TYPE_ROOTS = [path.join(ROOT, "node_modules/@types")];
const START = "<!-- excelts-index:start -->";
const END = "<!-- excelts-index:end -->";

/**
 * Classes that are still classes under another name, as the guide's sections say (§5.1, §6, §8.2, §8.4, §8.5). Their
 * members are compared with the new class's rather than required in the guide one by one.
 */
const CLASS_RENAMES: Record<string, string> = {
  CsvParserStream: "documonster/csv :: Csv.ParserStream",
  CsvFormatterStream: "documonster/csv :: Csv.FormatterStream",
  PdfDocumentBuilder: "documonster/pdf :: Pdf.Builder",
  PdfPageBuilder: "documonster/pdf :: Pdf.PageBuilder",
  PdfEditor: "documonster/pdf :: Pdf.Editor",
  PdfEditorPage: "documonster/pdf :: Pdf.EditorPage",
  XmlWriter: "documonster/xml :: Xml.Writer",
  XmlStreamWriter: "documonster/xml :: Xml.StreamWriter",
  StreamingDocxWriter: "documonster/word :: Stream.DocumentWriter",
  StringChunk: "documonster/stream :: StreamDataChunk",
  ByteChunk: "documonster/stream :: StreamDataChunk",
  JsonDataSource: "documonster/word :: DocxDataSource",
  XmlDataSource: "documonster/word :: DocxDataSource",
  CsvDataSource: "documonster/word :: DocxDataSource",
  CompositeDataSource: "documonster/word :: DocxDataSource"
};

/** The `documonster` entry an excelts entry became: the root and `./browser` are `documonster/excel`. */
function entryFor(oldEntry: string): string {
  const browser = oldEntry.endsWith(" [browser]") || oldEntry === "x/browser";
  let sub = oldEntry.replace(" [browser]", "").slice("x".length);
  if (sub === "" || sub === "/browser") {
    sub = "/excel";
  }
  return `documonster${sub}${browser ? " [browser]" : ""}`;
}

function both(
  root: string,
  label: string,
  names: "published" | "declared" = "published"
): ApiRecord {
  return {
    ...dumpApi(root, "node", label, TYPE_ROOTS, names),
    ...dumpApi(root, "browser", label, TYPE_ROOTS, names)
  };
}

/** The excelts record keyed by the `documonster` entries its entries became, so the two can be diffed. */
function rekey(old: ApiRecord): ApiRecord {
  const out: ApiRecord = {};
  for (const [k, v] of Object.entries(old)) {
    const [e, n] = k.split(" :: ") as [string, string];
    out[`${entryFor(e)} :: ${n}`] = v;
  }
  return out;
}

/** The names excelts and `documonster` both publish whose type differs, as Markdown. */
function typeChanges(
  old: ApiRecord,
  now: ApiRecord,
  oldDeclared: ApiRecord,
  nowDeclared: ApiRecord
): string {
  const d = diffApi(rekey(old), now, { before: rekey(oldDeclared), after: nowDeclared });
  const byEntry = new Map<string, string[]>();
  // A change that only narrows a byte result from `Uint8Array<ArrayBufferLike>` to `Uint8Array<ArrayBuffer>` asks
  // nothing of a caller, so it is listed by name after the changes that do, rather than drowning them.
  const bytesOnly = new Map<string, string[]>();
  for (const [k, a, b] of d.changed) {
    const [e, n] = k.split(" :: ") as [string, string];
    if (/(?:\.|#)(isError|captureStackTrace|prepareStackTrace|stackTraceLimit)$|__/.test(n)) {
      continue;
    }
    if (b.sig.replaceAll("Uint8Array<ArrayBuffer>", "Uint8Array<ArrayBufferLike>") === a.sig) {
      (bytesOnly.get(e) ?? bytesOnly.set(e, []).get(e)!).push(`\`${n}\``);
      continue;
    }
    (byEntry.get(e) ?? byEntry.set(e, []).get(e)!).push(
      `| \`${n}\` | ${sigCell(a.sig, b.sig)} | ${sigCell(b.sig, a.sig)} |`
    );
  }
  let md = `# Type changes from excelts 10.2.0

Every name \`@cj-tech-master/excelts\` 10.2.0 and \`documonster\` both publish whose type differs, generated by
\`pnpm verify:excelts-guide --write\` from the two packages' declarations. A name that moved or was renamed is not
here — [the migration guide](../MIGRATION_EXCELTS_TO_DOCUMONSTER.md) says where it went; this file is for the names
that stayed and changed underneath. Differences that are only printing (a union in another order, a type printed
under another of its names) are left out, as in [the release diff](api-changes-since-v0.21.1.md).
`;
  for (const [e, rows] of [...byEntry].sort(([a], [b]) => a.localeCompare(b))) {
    md += `\n## \`${e}\`\n\n| Name | excelts 10.2.0 | documonster |\n| ---- | -------------- | ----------- |\n${rows.join("\n")}\n`;
  }
  if (bytesOnly.size) {
    md += `
## Byte results narrowed to \`Uint8Array<ArrayBuffer>\`

These names changed only in that a \`Uint8Array<ArrayBufferLike>\` they returned or yielded is now a
\`Uint8Array<ArrayBuffer>\` — a plain array over a plain \`ArrayBuffer\`, never a \`Buffer\` or a view of a
\`SharedArrayBuffer\`. It is assignable everywhere the old type was, so code that compiled keeps compiling.
`;
    for (const [e, names] of [...bytesOnly].sort(([a], [b]) => a.localeCompare(b))) {
      md += `\n- \`${e}\`: ${names.join(", ")}\n`;
    }
  }
  return md;
}

// Only `|` is escaped, and that is the whole job: GitHub's table splitter treats every `\|` as an escaped pipe — however
// many backslashes precede it — and strips exactly that one backslash before the cell is parsed as Markdown. So the cell
// renders as it would outside a table, a backslash and `\|` included. Escaping backslashes as well would double every
// one shown inside a code span. Checked against GitHub's renderer (`gh api markdown -f mode=gfm`).
const esc = (s: string): string => s.replace(/\|/g, "\\|").replace(/\s+/g, " ");

interface IndexRow {
  oldEntry: string;
  name: string;
  now: string;
  note: string;
}

/**
 * Add to `record` what each entry of the package at `root` exports at run time without declaring it — a name its
 * declarations leave out is still one a caller can import, and so one the guide has to account for.
 */
async function addRuntimeExports(root: string, record: ApiRecord): Promise<void> {
  const pkg = JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8"));
  for (const [sub, spec] of Object.entries<{ import?: { default?: string } }>(pkg.exports)) {
    const file = spec?.import?.default;
    if (!file?.endsWith(".js")) {
      continue;
    }
    const mod: Record<string, unknown> = await import(pathToFileURL(path.join(root, file)).href);
    const entry = sub === "." ? "x" : `x${sub.slice(1)}`;
    for (const name of Object.keys(mod)) {
      if (!record[`${entry} :: ${name}`] && !record[`${entry} :: ${name} (type)`]) {
        record[`${entry} :: ${name}`] = { kind: "undeclared export", sig: "", deprecated: false };
      }
    }
  }
}

async function main(): Promise<void> {
  if (!fs.existsSync(path.join(ROOT, "dist/types"))) {
    console.error("✗ verify:excelts-guide — dist/types is missing. Run pnpm build:esm.");
    process.exit(2);
  }
  const work = fs.mkdtempSync(path.join(os.tmpdir(), "documonster-excelts-guide-"));
  try {
    console.log(`verify:excelts-guide — fetching ${EXCELTS}`);
    const oldRoot = fetchPackage(EXCELTS, work);
    const old = both(oldRoot, "x");
    await addRuntimeExports(oldRoot, old);
    const now = both(ROOT, "documonster");
    const changes = typeChanges(
      old,
      now,
      both(oldRoot, "x", "declared"),
      both(ROOT, "documonster", "declared")
    );

    const byEntry = new Map<string, string[]>();
    for (const k of Object.keys(now)) {
      const [e, n] = k.split(" :: ") as [string, string];
      (byEntry.get(e) ?? byEntry.set(e, []).get(e)!).push(n.replace(" (type)", ""));
    }
    // An entry has a browser record only when its browser declarations differ; otherwise it is the Node one.
    const own = (e: string): string => (byEntry.has(e) ? e : e.replace(" [browser]", ""));
    const get = (e: string, n: string): ApiEntry | undefined =>
      now[`${own(e)} :: ${n}`] ?? now[`${own(e)} :: ${n} (type)`] ?? now[`${own(e)} :: ${n}?`];
    /** `n` itself, or the same name inside a namespace of `e` — live ones first. */
    const locate = (e: string, n: string): { name: string; entry: ApiEntry } | undefined => {
      const same = get(e, n);
      if (same && !same.deprecated) {
        return { name: n, entry: same };
      }
      const inside = (byEntry.get(own(e)) ?? []).filter(
        x => !x.includes("#") && x.endsWith(`.${n}`)
      );
      const live = inside.find(x => !get(e, x)!.deprecated);
      const pick = live ?? (same ? n : inside[0]);
      return pick ? { name: pick, entry: get(e, pick)! } : undefined;
    };
    /** The replacement a deprecation names (`Use \`openZip\``), qualified as the deprecated name is. */
    const replacementOf = (e: string, deprecatedName: string, text: string): string | undefined => {
      const use = /Use `([^`]+)`/.exec(text)?.[1];
      if (!use) {
        return undefined;
      }
      const ns = deprecatedName.includes(".")
        ? deprecatedName.slice(0, deprecatedName.lastIndexOf("."))
        : "";
      // A deprecation names its replacement relative to the namespace it sits in (`Use \`Crypto.decryptAes\``
      // inside `Archive`), so try it qualified by each enclosing namespace, innermost first.
      const scopes = ns.split(".").map((_, i, parts) => parts.slice(0, parts.length - i).join("."));
      const bare = use.replace(/^new /, "");
      const isNew = bare !== use ? "new " : "";
      for (const candidate of [
        ...scopes.filter(Boolean).map(scope => `${isNew}${scope}.${bare}`),
        use
      ]) {
        const target = candidate.replace(/^new /, "").replace(/\(.*$/, "");
        const v = get(e, target);
        if (v && !v.deprecated) {
          return candidate;
        }
      }
      return undefined;
    };

    const rows = new Map<string, IndexRow>();
    const undescribed: { oldEntry: string; name: string; kind: string }[] = [];
    const classes = new Map<string, { entry: string; name: string } | undefined>();

    for (const [k, v] of Object.entries(old)) {
      const [oe, n0] = k.split(" :: ") as [string, string];
      const n = n0.replace(" (type)", "");
      if (/[.#]/.test(n)) {
        continue;
      }
      const e = entryFor(oe);
      const found = locate(e, n);
      if (v.kind === "class") {
        const renamed = CLASS_RENAMES[n];
        const [re, rn] = renamed ? (renamed.split(" :: ") as [string, string]) : [];
        const target =
          renamed && re && rn
            ? { entry: e.endsWith(" [browser]") ? `${re} [browser]` : re, name: rn }
            : found
              ? { entry: e, name: found.name }
              : undefined;
        classes.set(`${oe} :: ${n}`, target);
      }
      if (found && found.name === n && !found.entry.deprecated) {
        continue; // unchanged
      }
      const key = `${oe.replace(" [browser]", "")} :: ${n}`;
      if (found && !found.entry.deprecated) {
        rows.set(key, { oldEntry: oe, name: n, now: found.name, note: "" });
        continue;
      }
      if (found && found.entry.deprecated) {
        const text = String(found.entry.deprecated);
        rows.set(key, {
          oldEntry: oe,
          name: n,
          now: replacementOf(e, found.name, text) ?? "",
          note: `\`${found.name}\` is deprecated: ${text}`
        });
        continue;
      }
      undescribed.push({ oldEntry: oe, name: n, kind: v.kind });
    }

    const memberGaps: { oldEntry: string; cls: string; member: string }[] = [];
    for (const k of Object.keys(old)) {
      const m = /^(.+?) :: ([\w$]+)#([^?]+)\??$/.exec(k);
      if (!m) {
        continue;
      }
      const [, oe, cls, member] = m as unknown as [string, string, string, string];
      if (!classes.has(`${oe} :: ${cls}`)) {
        continue;
      }
      const target = classes.get(`${oe} :: ${cls}`);
      if (target) {
        const sep = get(target.entry, target.name)?.kind === "interface" ? "." : "#";
        const v = get(target.entry, `${target.name}${sep}${member}`);
        if (v && !v.deprecated) {
          continue;
        }
        if (v && v.deprecated) {
          const text = String(v.deprecated);
          const use = /Use `([^`]+)`/.exec(text)?.[1];
          rows.set(`${oe.replace(" [browser]", "")} :: ${cls}#${member}`, {
            oldEntry: oe,
            name: `${cls}#${member}`,
            now: use ? `${target.name}#${use.replace(/^[\w$]+\./, "")}` : "",
            note: `deprecated: ${text}`
          });
          continue;
        }
      }
      // Error classes carry `Error`'s own members, and `__…` fields are type brands no code reads.
      if (/Error$/.test(cls) || member.startsWith("__")) {
        continue;
      }
      memberGaps.push({ oldEntry: oe, cls, member });
    }

    // The generated index.
    const byOldEntry = new Map<string, IndexRow[]>();
    for (const row of rows.values()) {
      const e = row.oldEntry.replace(" [browser]", "");
      (byOldEntry.get(e) ?? byOldEntry.set(e, []).get(e)!).push(row);
    }
    let index = `${START}\n\n`;
    // `./browser` was the root entry's browser build, and both became `documonster/excel`: list only its differences.
    const rootRows = new Set(
      (byOldEntry.get("x") ?? []).map(r => `${r.name}\u0000${r.now}\u0000${r.note}`)
    );
    const browserRows = (byOldEntry.get("x/browser") ?? []).filter(
      r => !rootRows.has(`${r.name}\u0000${r.now}\u0000${r.note}`)
    );
    if (byOldEntry.has("x/browser")) {
      byOldEntry.set("x/browser", browserRows);
    }
    for (const [e, list] of [...byOldEntry].sort(([a], [b]) => a.localeCompare(b))) {
      const spec = `@cj-tech-master/excelts${e.slice(1)}`;
      if (e === "x/browser") {
        index += `**\`${spec}\`** → \`${entryFor(e)}\` — every row of \`@cj-tech-master/excelts\` above applies; only the differences are listed\n\n`;
        if (!list.length) {
          continue;
        }
        index += `| excelts | documonster | Note |\n| --- | --- | --- |\n`;
      } else {
        index += `**\`${spec}\`** → \`${entryFor(e)}\`\n\n| excelts | documonster | Note |\n| --- | --- | --- |\n`;
      }
      for (const r of list.sort((a, b) => a.name.localeCompare(b.name))) {
        index += `| \`${r.name}\` | ${r.now ? `\`${r.now}\`` : "—"} | ${esc(r.note)} |\n`;
      }
      index += "\n";
    }
    index += END;

    let text = fs.readFileSync(GUIDE, "utf8");
    const a = text.indexOf(START);
    const b = text.indexOf(END);
    if (a < 0 || b < a) {
      console.error(`✗ verify:excelts-guide — the guide has no ${START} … ${END} markers`);
      process.exit(1);
    }
    const current = text.slice(a, b + END.length);
    const problems: string[] = [];
    if (process.argv.includes("--write")) {
      text = text.slice(0, a) + index + text.slice(b + END.length);
      fs.writeFileSync(GUIDE, text);
      fs.writeFileSync(TYPE_CHANGES, changes);
      console.log(
        `verify:excelts-guide — wrote the name index and ${path.relative(ROOT, TYPE_CHANGES)} (run pnpm format)`
      );
    } else {
      if (normalized(current) !== normalized(index)) {
        problems.push("the generated name index (§12) is out of date — run with --write");
      }
      const written = fs.existsSync(TYPE_CHANGES) ? fs.readFileSync(TYPE_CHANGES, "utf8") : "";
      if (normalized(written) !== normalized(changes)) {
        problems.push(`${path.relative(ROOT, TYPE_CHANGES)} is out of date — run with --write`);
      }
    }

    const handWritten = text.slice(0, text.indexOf(START)) + text.slice(text.indexOf(END));
    const mentioned = mentionedNames(handWritten);
    for (const { oldEntry, name, kind } of undescribed) {
      if (!mentioned.has(name)) {
        problems.push(
          `${oldEntry.replace(/^x/, "excelts")}: \`${name}\` (${kind}) is not described`
        );
      }
    }
    // A member counts as described where the guide names it beside its owner (`ws.`, `Worksheet`, …).
    for (const { oldEntry, cls, member } of memberGaps) {
      const named = describesMember(handWritten, cls, member);
      if (!named) {
        problems.push(
          `${oldEntry.replace(/^x/, "excelts")}: \`${cls}#${member}\` is not described`
        );
      }
    }

    problems.push(...checkReplacements(handWritten, now));

    if (problems.length) {
      console.error(`✗ verify:excelts-guide — ${problems.length} problem(s):`);
      for (const p of [...new Set(problems)]) {
        console.error(`  ${p}`);
      }
      process.exit(1);
    }
    console.log(
      `✓ verify:excelts-guide — every name ${EXCELTS} publishes is accounted for, and every replacement exists.`
    );
  } finally {
    fs.rmSync(work, { recursive: true, force: true });
  }
}

await main();
