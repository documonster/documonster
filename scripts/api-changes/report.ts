/**
 * Render an API diff as Markdown: the exhaustive per-name list (`docs/api-changes-since-v<base>.md`), and the
 * deprecation appendix and summary table `MIGRATION.md` carries. Used by `scripts/api-changes.ts`.
 */
import type { ApiDiff } from "./diff.ts";
import type { ApiEntry, ApiRecord } from "./dump.ts";
import { sigCell as cell } from "./sig-cell.ts";
import type { ToolCatalog } from "./tools.ts";

const NOISE = /(?:\.|#)(isError|captureStackTrace|prepareStackTrace|stackTraceLimit)$|__@/;
const INTERNAL_NOTE = /^Internal|not part of the (public )?API|Internal state/;

const noisy = (k: string): boolean => NOISE.test(k);
const esc = (s: string): string => s.replace(/\|/g, "\\|");
const code = (s: string): string => "`" + esc(s.replace(/`/g, "'")) + "`";
// The namespaces are published as `Draw` / `Mermaid`; the declarations sometimes print the import alias.
const shown = (s: string): string =>
  s.replace(/\bDrawNs\./g, "Draw.").replace(/\bMermaidNs\./g, "Mermaid.");
const sigCell = (s: string, other?: string): string => cell(shown(s), other && shown(other));
const isInternal = (name: string, e: ApiEntry): boolean =>
  /[.#]_/.test(name) || INTERNAL_NOTE.test(String(e.deprecated));
const split = (k: string): { e: string; n: string; top: string } => {
  const [e, n] = k.split(" :: ") as [string, string];
  return { e, n, top: n.split(/[.#]/)[0]!.replace(" (type)", "") };
};
const entryOrder = (a: string, b: string): number =>
  a.replace(" [browser]", "~").localeCompare(b.replace(" [browser]", "~"));

type Section =
  | "removed"
  | "optional"
  | "changed"
  | "deprecated"
  | "overload"
  | "addedMember"
  | "addedNew";
const SECTIONS: [Section, string, string][] = [
  ["removed", "Removed", "| Name | Kind | Before |\n| ---- | ---- | ------ |"],
  [
    "optional",
    "Became optional (absent instead of `null` / a sentinel)",
    "| Name | Before | Now |\n| ---- | ------ | --- |"
  ],
  [
    "changed",
    "Changed signature or type",
    "| Name | Kind | Before | Now |\n| ---- | ---- | ------ | --- |"
  ],
  [
    "deprecated",
    "Newly deprecated (still works; removed in 1.0)",
    "| Name | Kind | Replacement |\n| ---- | ---- | ----------- |"
  ],
  [
    "overload",
    "Newly deprecated call forms (the name stays; the form described is removed in 1.0)",
    "| Name | Replacement |\n| ---- | ----------- |"
  ],
  [
    "addedMember",
    "Added to exports that already existed",
    "| Name | Kind | Signature |\n| ---- | ---- | --------- |"
  ],
  [
    "addedNew",
    "Added exports, with every member",
    "| Name | Kind | Signature |\n| ---- | ---- | --------- |"
  ]
];

export interface ReportInput {
  base: string;
  before: ApiRecord;
  after: ApiRecord;
  diff: ApiDiff;
  mcpDiff: ApiDiff;
  toolsBefore: ToolCatalog;
  toolsAfter: ToolCatalog;
}

function group(
  input: ReportInput
): Record<string, Partial<Record<Section | "internal", string[]>>> {
  const { before, after, diff } = input;
  const renamedFrom: Record<string, Record<string, Set<string>>> = {};
  for (const [k, v] of Object.entries(after)) {
    const m =
      v.kind === "type" ? /^type ([\w$]+)(?:<[^=]*>)? = ([\w$.]+)(?:<[^;]*>)?;$/.exec(v.sig) : null;
    const to = m?.[2]?.split(".").pop();
    if (m && to && m[1] !== to && v.deprecated) {
      ((renamedFrom[split(k).e] ??= {})[to] ??= new Set()).add(m[1]!);
    }
  }
  const out: Record<string, Partial<Record<Section | "internal", string[]>>> = {};
  const put = (e: string, sec: Section | "internal", row: string): void => {
    ((out[e] ??= {})[sec] ??= []).push(row);
  };
  for (const [k, a] of diff.removed) {
    if (noisy(k)) {
      continue;
    }
    const { e, n } = split(k);
    const optional = after[k + "?"];
    if (!n.endsWith("?") && optional) {
      put(
        e,
        "optional",
        `| ${code(n)} | ${sigCell(a.sig, optional.sig)} | ${sigCell(optional.sig, a.sig)} |`
      );
    } else {
      put(e, "removed", `| ${code(n)} | ${a.kind} | ${sigCell(a.sig)} |`);
    }
  }
  for (const [k, a, b] of diff.changed) {
    if (!noisy(k)) {
      const { e, n } = split(k);
      put(
        e,
        "changed",
        `| ${code(n)} | ${a.kind === b.kind ? b.kind : `${a.kind} → ${b.kind}`} | ${sigCell(a.sig, b.sig)} | ${sigCell(b.sig, a.sig)} |`
      );
    }
  }
  for (const [k, b] of diff.deprecated) {
    if (noisy(k)) {
      continue;
    }
    const { e, n } = split(k);
    if (isInternal(n, b)) {
      put(e, "internal", n);
    } else {
      put(e, "deprecated", `| ${code(n)} | ${b.kind} | ${esc(String(b.deprecated))} |`);
    }
  }
  for (const [k, b] of diff.deprecatedOverload) {
    if (!noisy(k) && k in before) {
      const { e, n } = split(k);
      put(e, "overload", `| ${code(n)} | ${esc(b.partial ?? "")} |`);
    }
  }
  for (const [k, b] of diff.added) {
    if (noisy(k)) {
      continue;
    }
    const { e, n, top } = split(k);
    if (n.endsWith("?") && before[k.slice(0, -1)]) {
      continue; // reported as "became optional"
    }
    if (n.endsWith(" (type)") && before[k.replace(" (type)", "")]?.kind === "class") {
      continue; // a class now published as a value and a type of one name
    }
    const isTop = !/[.#]/.test(n.replace(" (type)", ""));
    const isNewTop = !(`${e} :: ${top}` in before) && !(`${e} :: ${top} (type)` in before);
    let note = "";
    const was = isTop ? renamedFrom[e]?.[n.replace(" (type)", "")] : undefined;
    if (was) {
      note += " — renamed from " + [...was].map(code).join(", ");
    }
    if (b.deprecated) {
      note += " — deprecated: " + esc(String(b.deprecated));
    }
    put(
      e,
      isNewTop ? "addedNew" : "addedMember",
      `| ${code(n)} | ${b.kind}${note} | ${sigCell(b.sig)} |`
    );
  }
  return out;
}

/** The heading ID GitHub gives an entry's `## \`entry\`` heading. */
function entryAnchor(entry: string): string {
  return entry
    .toLowerCase()
    .replace(/[^\p{L}\p{M}\p{Nd}\p{Pc}\- ]/gu, "")
    .replace(/ /g, "-");
}

/**
 * The summary rows: one per entry, counting the rows of each section. Each entry links to its section of the full
 * report; `report` is that file's path relative to where the table is written (empty when it is the report itself).
 */
export function summaryTable(input: ReportInput, report = ""): string {
  const link = (e: string): string => `[${code(e)}](${report}#${entryAnchor(e)})`;
  const groups = group(input);
  const head =
    "| Entry | Removed | Became optional | Changed | Deprecated | Deprecated call forms | Added to existing | Added (new exports and their members) |\n" +
    "| ----- | ------: | --------------: | ------: | ---------: | --------------------: | ----------------: | ------------------------------------: |\n";
  const rows = Object.keys(groups)
    .sort(entryOrder)
    .map(e => {
      const s = groups[e]!;
      const n = (k: Section): number => s[k]?.length ?? 0;
      return `| ${link(e)} | ${n("removed")} | ${n("optional")} | ${n("changed")} | ${n("deprecated")} | ${n("overload")} | ${n("addedMember")} | ${n("addedNew")} |`;
    });
  const mcp = input.mcpDiff;
  rows.push(
    `| ${link("@documonster/mcp")} | ${mcp.removed.length} | 0 | ${mcp.changed.length} | ${mcp.deprecated.length} | ${mcp.deprecatedOverload.length} | 0 | ${mcp.added.length} |`
  );
  return head + rows.join("\n") + "\n";
}

function flatten(
  o: unknown,
  prefix = "",
  out: Record<string, unknown> = {}
): Record<string, unknown> {
  if (o && typeof o === "object") {
    for (const [k, v] of Object.entries(o)) {
      flatten(v, `${prefix}/${k}`, out);
    }
  } else {
    out[prefix] = o;
  }
  return out;
}

/** The exhaustive per-name list. */
export function fullReport(input: ReportInput): string {
  const { base, mcpDiff, toolsBefore, toolsAfter } = input;
  const groups = group(input);
  let md = `# Public API changes since v${base}

The complete list of what changed in the public API between the published \`documonster@${base}\` /
\`@documonster/mcp@${base}\` and the current tree. [MIGRATION.md](../MIGRATION.md#unreleased) says what to do about it;
this file says, name by name, what changed.

## How to use this list

- **Looking for an old name?** Search this page for it. A deprecated name's row gives its replacement; a name under
  "Changed signature or type" shows the old and new signatures side by side.
- **Looking for one module?** Follow its link in the summary below. Each entry has the same sections, in this order:
  removed, became optional, changed, deprecated, deprecated call forms, added to existing, added.
- **Want to know what to do, not what changed?** Read [MIGRATION.md](../MIGRATION.md#the-most-common-changes-before-and-after)
  instead; this list does not describe behaviour that changed without a change of type.

## Summary

${summaryTable(input)}
## How this list is made

It is generated by \`pnpm api:changes\` (\`scripts/api-changes.ts\`), not written. The declarations of every \`exports\`
subpath are read with the TypeScript compiler — the published package for ${base}, a fresh build for the current tree
— once with the default conditions and once with \`browser\`. Every public name is recorded: each namespace member
(recursively), each class's constructor, instance and static members, each overload of a function, each field of an
interface or object type, each enum member and constant. Private and protected members are not public and are
skipped. The two records are then compared name by name.

- **Signatures are shown as the compiler prints them**, one overload per line. A line longer than 160 characters is
  clipped to the stretch that differs from the signature it is compared with, marked \`…\` where cut, and at most four
  overloads are shown per cell (the rest are counted). The whole signature is one hover away in your editor, or in
  the package's \`.d.ts\` files.
- A change is reported only when the printed type differs after removing differences that are only printing:
  a type printed under its new name (when the old name is now a deprecated alias of it), an internal import alias
  (\`…Impl\`), and the order of a union of string literals.
- An interface that became an alias of a renamed interface (\`type Old = New\`) is compared field by field, so a
  field that changed shows up as its own row.
- \`[browser]\` sections list only what differs from the Node entry of the same subpath.
- \`Name#member\` is an instance member; \`Name.member\` is a namespace member, a static member or a field.
- Members that every \`Error\` has (\`captureStackTrace\`, \`stackTraceLimit\`, …) are left out.
- Behaviour that changed without a change of type is not visible here; MIGRATION.md describes it.
`;
  for (const e of Object.keys(groups).sort(entryOrder)) {
    const s = groups[e]!;
    md += `\n## ${code(e)}\n`;
    for (const [key, title, head] of SECTIONS) {
      if (s[key]) {
        md += `\n### ${title}\n\n${head}\n${s[key]!.join("\n")}\n`;
      }
    }
    if (s.internal) {
      const byParent: Record<string, string[]> = {};
      for (const n of s.internal) {
        const p = n.slice(0, Math.max(n.lastIndexOf("."), n.lastIndexOf("#")));
        (byParent[p] ??= []).push(n.slice(p.length + 1));
      }
      md +=
        "\n### Internal fields now marked deprecated\n\nPublished by accident; stripped from the types in 1.0, with no replacement. Use the namespace functions instead.\n\n";
      for (const [p, l] of Object.entries(byParent)) {
        md += `- ${code(p)}: ${l.map(code).join(", ")}\n`;
      }
    }
  }

  md += "\n## `@documonster/mcp`\n\n### TypeScript entry\n\n";
  const mcpRows = [
    ...mcpDiff.removed.map(
      ([k, a]) => `| ${code(k.split(" :: ")[1]!)} | removed (${a.kind}) | ${sigCell(a.sig)} |`
    ),
    ...mcpDiff.changed.map(
      ([k, , b]) => `| ${code(k.split(" :: ")[1]!)} | changed (${b.kind}) | ${sigCell(b.sig)} |`
    ),
    ...mcpDiff.added.map(
      ([k, b]) => `| ${code(k.split(" :: ")[1]!)} | added (${b.kind}) | ${sigCell(b.sig)} |`
    )
  ];
  md += mcpRows.length
    ? `| Name | Change | Signature |\n| ---- | ------ | --------- |\n${mcpRows.join("\n")}\n`
    : "No change.\n";

  const names = Object.keys(toolsAfter.tools);
  const added = names.filter(n => !(n in toolsBefore.tools));
  const removed = Object.keys(toolsBefore.tools).filter(n => !(n in toolsAfter.tools));
  md += `\n### Tools\n\n${names.length} tools in the groups ${toolsAfter.groups.map(code).join(", ")}.`;
  md += added.length ? ` Added: ${added.map(code).join(", ")}.` : " No tool was added.";
  md += removed.length ? ` Removed: ${removed.map(code).join(", ")}.` : " No tool was removed.";
  md += " Input schemas were compared as JSON Schema.\n\n";
  const toolRows: string[] = [];
  for (const name of names.filter(n => n in toolsBefore.tools)) {
    const a = toolsBefore.tools[name]!;
    const b = toolsAfter.tools[name]!;
    const fa = flatten(a.schema);
    const fb = flatten(b.schema);
    const props = new Map<string, string>();
    for (const k of new Set([...Object.keys(fa), ...Object.keys(fb)])) {
      if (fa[k] !== fb[k]) {
        const p = k.replace(/\/(description|type|enum\/\d+|default)$/, "");
        props.set(
          p,
          !Object.keys(fa).some(x => x.startsWith(p + "/") || x === p)
            ? "added"
            : !Object.keys(fb).some(x => x.startsWith(p + "/") || x === p)
              ? "removed"
              : "changed"
        );
      }
    }
    const notes = [...props].map(([p, how]) => `schema ${code(p)}: ${how}`);
    if (a.title !== b.title) {
      notes.push("title changed");
    }
    if (JSON.stringify(a.group) !== JSON.stringify(b.group)) {
      notes.push(`group ${code(JSON.stringify(a.group))} → ${code(JSON.stringify(b.group))}`);
    }
    if (JSON.stringify(a.annotations) !== JSON.stringify(b.annotations)) {
      notes.push("annotations changed");
    }
    if (a.description !== b.description) {
      notes.push("description changed");
    }
    if (notes.length) {
      toolRows.push(`| ${code(name)} | ${notes.join("; ")} |`);
    }
  }
  md += toolRows.length
    ? `| Tool | Change |\n| ---- | ------ |\n${toolRows.join("\n")}\n\nEvery other tool is identical. `
    : "Every tool is identical. ";
  md += "Behaviour changes are described in [MIGRATION.md](../MIGRATION.md#documonstermcp).\n";
  return md;
}

/** The "Every deprecation in this release" appendix of MIGRATION.md, from `#### Deprecated names` on. */
export function deprecationAppendix(input: ReportInput): string {
  const { before, diff } = input;
  const names: Record<string, Map<string, string>> = {};
  const forms: Record<string, Map<string, string>> = {};
  const internals: Record<string, Set<string>> = {};
  for (const [k, b] of diff.deprecated) {
    if (noisy(k)) {
      continue;
    }
    const { e } = split(k);
    const n = split(k).n.replace(" (type)", "");
    if (isInternal(n, b)) {
      const p = n.slice(0, Math.max(n.lastIndexOf("."), n.lastIndexOf("#")));
      (internals[`${e} :: ${p}`] ??= new Set()).add(n.slice(p.length + 1));
    } else {
      (names[e] ??= new Map()).set(n, `| ${code(n)} | ${esc(String(b.deprecated))} |`);
    }
  }
  for (const [k, b] of diff.deprecatedOverload) {
    if (!noisy(k) && k in before) {
      const { e, n } = split(k);
      (forms[e] ??= new Map()).set(n, `| ${code(n)} | ${esc(b.partial ?? "")} |`);
    }
  }
  let md = "#### Deprecated names\n\n";
  for (const e of Object.keys(names).sort(entryOrder)) {
    md += `**${code(e)}**\n\n| Deprecated | Replacement |\n| ---------- | ----------- |\n${[...names[e]!.values()].join("\n")}\n\n`;
  }
  md +=
    "#### Deprecated call forms\n\nThe name stays; the call form described is removed in 1.0.\n\n";
  for (const e of Object.keys(forms).sort(entryOrder)) {
    md += `**${code(e)}**\n\n| Name | Use instead |\n| ---- | ----------- |\n${[...forms[e]!.values()].join("\n")}\n\n`;
  }
  md +=
    "#### Internal fields marked deprecated\n\nThese were published by accident. They are stripped from the types in 1.0 and have no replacement — read and\nchange the object through its namespace instead.\n\n";
  for (const [p, l] of Object.entries(internals)) {
    const [e, n] = p.split(" :: ") as [string, string];
    md += `- ${code(n)} (${code(e)}): ${[...l].map(code).join(", ")}\n`;
  }
  return md;
}
