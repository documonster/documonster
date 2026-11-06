/**
 * A part means the same whatever prefixes it binds its namespaces to. So every fixture, rewritten with
 * every prefix renamed — `p` to `zp`, the default namespace to a `zdflt` prefix — must read into the model
 * the original does. A namespace the reader resolves by its prefix rather than its URI shows up here as
 * a field that is lost or comes out different.
 *
 * One difference is by design and is undone before comparing: a namespace the reader does not know keeps
 * the prefix the document gave it, so a model that records such a prefix (a root's preserved
 * declarations, its `mc:Ignorable`) records `zp`. Spelling is all that undoing reverts — a field the
 * reader failed to fill stays missing.
 */
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";

import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { expandedNames, reprefix } from "@excel/__tests__/helpers/expanded-names";
import { Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

const DATA = join(__dirname, "../../__tests__/data");

const renamed = (prefix: string) => `z${prefix || "dflt"}`;

/**
 * `bytes` with every XML part rewritten, its entries kept in their order in the central directory: the
 * order they are read in, which decides the order of the workbook's media. Returns the prefixes the
 * parts declared, too.
 */
async function rewritten(bytes: Uint8Array): Promise<{ bytes: Uint8Array; prefixes: Set<string> }> {
  const prefixes = new Set<string>();
  const files = [...(await extractAll(bytes)).values()].filter(file => !file.path.endsWith("/"));
  const out = await createZip(
    files.map(({ path: name, data }) => ({
      name,
      data: /\.(xml|rels|vml)$/.test(name)
        ? new TextEncoder().encode(
            reprefix(new TextDecoder().decode(data), prefix => {
              prefixes.add(prefix);
              return renamed(prefix);
            })
          )
        : data
    })),
    { noSort: true }
  );
  return { bytes: out, prefixes };
}

/**
 * XML a model keeps as text — a part written back verbatim (a theme) or a fragment — compared by what it
 * means: its names resolved to namespaces. A fragment that relies on its context for a declaration
 * cannot be resolved alone and is compared as text.
 */
function meaning(value: string): unknown {
  if (!value.startsWith("<")) {
    return value;
  }
  try {
    return { xml: expandedNames(value) };
  } catch {
    return value;
  }
}

/** The workbook's model as plain data. `prefixes`: the renamed ones to spell back. */
async function model(bytes: Uint8Array, prefixes?: ReadonlySet<string>): Promise<unknown> {
  const back = new Map([...(prefixes ?? [])].map(prefix => [renamed(prefix), prefix]));
  // A prefix is spelt back where text names one: before a colon (`zx:name`, `</zx:name>` — the end tag was missed
  // until a kept fragment first ended in a prefixed one), after `xmlns:`, and as an
  // entry of a prefix list (an `Ignorable`, a `Requires`) — not in a value that merely looks like one.
  const unrename = (text: string, list = false): string =>
    back.size === 0
      ? text
      : text.replace(
          list
            ? /(^|\s)(z[\w.-]+)(?=\s|$)/g
            : /(^|xmlns:|[\s"<]|<\/)(z[\w.-]+)(?=:)|(xmlns:)(z[\w.-]+)$/g,
          (
            whole,
            before: string | undefined,
            name: string | undefined,
            xmlns?: string,
            declared?: string
          ) => {
            const prefix = back.get((name ?? declared)!);
            return prefix === undefined ? whole : (before ?? xmlns) + prefix;
          }
        );
  const wb = Workbook.create();
  await Workbook.read(wb, bytes);
  return JSON.parse(
    JSON.stringify(Workbook.getModel(wb), (key, value) => {
      if (typeof value === "string") {
        const xml = meaning(value);
        return typeof xml === "string" ? unrename(xml, /ignorable|requires/i.test(key)) : xml;
      }
      if (value instanceof Uint8Array) {
        return meaning(
          new TextDecoder().decode(value).replace(/^\uFEFF?(?:<\?xml[^>]*\?>)?\s*/, "")
        );
      }
      if (value instanceof Map) {
        value = Object.fromEntries(value);
      }
      if (value !== null && typeof value === "object" && !Array.isArray(value)) {
        return Object.fromEntries(Object.entries(value).map(([k, v]) => [unrename(k), v]));
      }
      return value;
    })
  );
}

describe("every fixture, every prefix renamed, reads into the same model", () => {
  for (const file of readdirSync(DATA).filter(name => name.endsWith(".xlsx"))) {
    it(file, async () => {
      const original = new Uint8Array(readFileSync(join(DATA, file)));
      const { bytes, prefixes } = await rewritten(original);
      expect(await model(bytes, prefixes)).toEqual(await model(original));
    });
  }
});
