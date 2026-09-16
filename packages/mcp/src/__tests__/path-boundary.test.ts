/**
 * Path-boundary regression suite.
 *
 * `sandbox.test.ts` proves `resolveInRoot` is correct. This file proves every
 * argument that can name a host path actually goes through it — a different
 * failure, and the one that would matter. A tool that forgets the sandbox has a
 * working happy path and passes every behavioural test it owns; nothing but a
 * check like this notices.
 *
 * The hard part is not testing a path, it is *knowing where the paths are*, and
 * three earlier versions of this file got that wrong in instructive ways.
 *
 * The first listed the paths by hand and matched them by argument name. It
 * missed three real carriers nested in option arrays — `pdf_edit`'s diagram op,
 * `sheet_edit`'s image op, `sheet_write`'s per-sheet images.
 *
 * The second derived them from Zod's private `_zod.def`. That removed the
 * hand-written list but read a library's internals, where a renamed field
 * degrades quietly instead of loudly: the object case read `def.shape ?? {}`, so
 * had `shape` been renamed every field under it would have left the inventory
 * without an error. It also decided which nodes could hold a string from a
 * hardcoded list of Zod *type names* — the same guess-the-name mistake one layer
 * down — and silently dropped `template_fill`'s unconstrained `data` record.
 *
 * The third derived the inventory from Zod's public `toJSONSchema`, which fixed
 * both, but collapsed every branch of a discriminated union onto one accessor.
 * That made a quarter of the claims below vacuous: a value injected into
 * `ops[].source` landed on an `op: "rotate"` object, which does not read
 * `source`, so the assertion passed without exercising anything. Accessors are
 * therefore qualified by discriminator — `ops[op=diagram].source` — and each
 * branch is built from its own valid base in {@link BRANCH_BASES}, so an
 * injection reaches the code that actually reads the field.
 *
 * Every leaf must then be classified, and there is no third option:
 *
 * - listed in {@link NON_PATH_LEAVES}, a claim that the field names no file, or
 * - exercised by a case in {@link CASES}, asserting the call is refused.
 *
 * That is what makes the guard independent of naming: an argument called `dir`,
 * `file` or `location` fails this file until someone decides which side it is
 * on, where a check recognising only `path`/`from`/`out` waves it through.
 *
 * **What the non-path claims do and do not establish.** Each one is fed a
 * hostile path inside an otherwise valid call for its own branch, and must
 * neither return content from outside the root nor be refused as `outside_root`
 * — the latter would mean the field is resolved as a path, so it is a carrier
 * whatever it is called. What this cannot see is a tool that opens the file and
 * discards it without reporting anything, which would need a spy on `fs` and an
 * allow-list for the reads the server makes legitimately (its own output root,
 * staging directories, embedded fonts). That allow-list would be another
 * hand-maintained security claim, so the boundary is drawn here deliberately.
 */

import { mkdtemp, realpath, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { ArchiveFile } from "documonster/archive";
import { Workbook } from "documonster/excel";
import { Pdf } from "documonster/pdf";
import { Build, Document, Io } from "documonster/word";
import { z } from "zod";

import { resolveConfig, type ServerConfig } from "../config.js";
import { formatToolError, type ToolErrorCode } from "../errors.js";
import { ALL_TOOLS } from "../tools/index.js";
import type { AnyToolDefinition } from "../tools/types.js";

// ---------------------------------------------------------------------------
// Schema introspection
// ---------------------------------------------------------------------------

/** JSON Schema keywords that describe *values*, and so have to be walked. */
const STRUCTURAL_KEYWORDS = new Set([
  "type",
  "properties",
  "items",
  "additionalProperties",
  "anyOf",
  "oneOf",
  "allOf",
  "enum",
  "const"
]);

/** Keywords that constrain or describe without nesting a further schema. */
const ANNOTATION_KEYWORDS = new Set([
  "$id",
  "$schema",
  "default",
  "deprecated",
  "description",
  "examples",
  "exclusiveMaximum",
  "exclusiveMinimum",
  "format",
  "id",
  "maxItems",
  "maxLength",
  "maxProperties",
  "maximum",
  "minItems",
  "minLength",
  "minProperties",
  "minimum",
  "multipleOf",
  "nullable",
  "pattern",
  // Constrains a record's *keys*, not its values, so it nests nothing to walk.
  "propertyNames",
  "readOnly",
  "required",
  "title",
  "uniqueItems",
  "writeOnly"
]);

/**
 * The property every branch of a union pins to a literal, or undefined when the
 * union is not discriminated (`pages`, for instance, is `string | number[]`).
 */
function discriminatorOf(branches: readonly unknown[]): string | undefined {
  const keys = branches.map(branch => {
    const properties = (branch as { properties?: Record<string, unknown> }).properties ?? {};
    return Object.entries(properties).find(
      ([, schema]) => typeof (schema as { const?: unknown }).const === "string"
    )?.[0];
  });
  const first = keys[0];
  return first !== undefined && keys.every(key => key === first) ? first : undefined;
}

/** `ops[]` + `op=diagram` reads best as `ops[op=diagram]`. */
function qualifyBranch(accessor: string, key: string, value: string): string {
  return accessor.endsWith("[]")
    ? `${accessor.slice(0, -2)}[${key}=${value}]`
    : `${accessor}[${key}=${value}]`;
}

/**
 * Every free-form string reachable in a JSON Schema, as an accessor.
 *
 * Accessors read `ops[op=append].path`, `sheets[].images[].from`,
 * `images{}.from` — `[]` for an array element, `{}` for a record value,
 * `[key=value]` for one branch of a discriminated union. They are the
 * identifiers the classification below is keyed by, so they must stay stable.
 *
 * Anything that would *shrink* the inventory is loud rather than quiet, because
 * an inventory that lost an entry still passes every assertion in this file:
 *
 * - An unrecognised keyword throws. That covers `$ref`, `not`, `if`/`then` and
 *   anything a future Zod release starts emitting.
 * - An `object` with neither `properties` nor `additionalProperties`, and an
 *   `array` with no `items`, throw too. Both are legal JSON Schema meaning
 *   "unconstrained", and both used to yield nothing at all.
 * - A composition keyword beside a sibling `type` or `properties` throws. Those
 *   constraints are conjunctive, and returning early on the first one skips the
 *   rest.
 * - A schema with no `type` and no composition is unconstrained: it accepts a
 *   string, therefore possibly a path, so it is returned as a leaf and has to be
 *   classified. `template_fill`'s `data` record is one.
 *
 * The single exclusion is a closed value set (`enum`/`const`), and it is
 * structural rather than nominal: the permitted values are fixed in the schema,
 * so a caller cannot put a path in one. That is what removes `mode`, `action`,
 * `theme` and every `op` discriminator without anyone listing them.
 */
function stringLeaves(node: unknown, accessor: string, depth = 0): string[] {
  if (depth > 14) {
    throw new Error(`schema nested deeper than 14 levels at ${accessor}`);
  }
  if (node === null || typeof node !== "object" || Array.isArray(node)) {
    throw new Error(`not a JSON Schema object at ${accessor}`);
  }
  const schema = node as Record<string, unknown>;

  const unhandled = Object.keys(schema).filter(
    key => !STRUCTURAL_KEYWORDS.has(key) && !ANNOTATION_KEYWORDS.has(key)
  );
  if (unhandled.length > 0) {
    throw new Error(
      `stringLeaves does not handle the JSON Schema keyword(s) ${unhandled.join(", ")} ` +
        `at ${accessor}. Add them to the walk, or to ANNOTATION_KEYWORDS if they nest ` +
        `nothing — do not leave them skipped, or paths inside them become invisible here.`
    );
  }

  if ("enum" in schema || "const" in schema) {
    return [];
  }

  const composition = (["anyOf", "oneOf", "allOf"] as const).filter(keyword =>
    Array.isArray(schema[keyword])
  );
  if (composition.length > 0) {
    if (composition.length > 1 || "type" in schema || "properties" in schema) {
      throw new Error(
        `stringLeaves found ${composition.join("/")} beside a sibling constraint at ` +
          `${accessor}. JSON Schema applies those together; walking only the first ` +
          `would skip the rest.`
      );
    }
    const branches = schema[composition[0]!] as readonly unknown[];
    const discriminator = discriminatorOf(branches);
    return branches.flatMap(branch => {
      const value =
        discriminator === undefined
          ? undefined
          : ((branch as { properties?: Record<string, { const?: unknown }> }).properties ?? {})[
              discriminator
            ]?.const;
      return stringLeaves(
        branch,
        typeof value === "string" ? qualifyBranch(accessor, discriminator!, value) : accessor,
        depth + 1
      );
    });
  }

  const types = Array.isArray(schema.type)
    ? (schema.type as string[])
    : typeof schema.type === "string"
      ? [schema.type]
      : [];
  if (types.length === 0) {
    return [accessor];
  }

  const found: string[] = [];
  for (const type of types) {
    if (type === "string") {
      found.push(accessor);
      continue;
    }
    if (type === "object") {
      const properties = schema.properties as Record<string, unknown> | undefined;
      const values = schema.additionalProperties;
      const hasValues = values !== null && typeof values === "object";
      if (properties === undefined && !hasValues) {
        throw new Error(
          `unconstrained object at ${accessor}: neither properties nor ` +
            `additionalProperties. It accepts arbitrary strings, so it cannot be skipped.`
        );
      }
      for (const [key, value] of Object.entries(properties ?? {})) {
        found.push(...stringLeaves(value, accessor === "" ? key : `${accessor}.${key}`, depth + 1));
      }
      if (hasValues) {
        found.push(...stringLeaves(values, `${accessor}{}`, depth + 1));
      }
      continue;
    }
    if (type === "array") {
      const items = schema.items;
      if (items === undefined || items === null) {
        throw new Error(
          `unconstrained array at ${accessor}: no items. Its elements could be ` +
            `strings, so it cannot be skipped.`
        );
      }
      if (Array.isArray(items)) {
        for (const entry of items) {
          found.push(...stringLeaves(entry, `${accessor}[]`, depth + 1));
        }
      } else {
        found.push(...stringLeaves(items, `${accessor}[]`, depth + 1));
      }
    }
  }
  return found;
}

/**
 * Every string leaf in every registered tool, as `tool::accessor`.
 *
 * The schema comes from Zod's public `toJSONSchema` rather than from
 * `_zod.def`, so this survives a change to Zod's internal representation.
 * `io: "input"` asks for the shape a *caller* sends.
 */
function inventory(): string[] {
  const leaves = ALL_TOOLS.flatMap(tool => {
    const schema = z.toJSONSchema(z.object(tool.inputSchema), {
      io: "input",
      unrepresentable: "any"
    });
    return stringLeaves(schema, "").map(accessor => `${tool.name}::${accessor}`);
  });
  // Deduplicated: an accessor is a position in the request, and a non-discriminated
  // union repeats one — `pages` is `string | number[]`, and only the string arm
  // produces a leaf. Discriminated branches keep their own identity above, so
  // this no longer merges two different code paths into one entry.
  return [...new Set(leaves)].toSorted();
}

/**
 * Every string argument that does **not** name a host path.
 *
 * Each entry is a claim, checked behaviourally by "injecting a path into a
 * declared non-path argument" below, within the limits set out in this file's
 * header. The ones worth reading twice are the `source` fields (Mermaid diagram
 * *text*, deliberately not a file) and the cell and page addresses (`ref`,
 * `range`, `at`, `anchor`, `pages`), which are coordinates inside a document
 * rather than on disk.
 */
const NON_PATH_LEAVES: ReadonlySet<string> = new Set([
  // archive_read
  "archive_read::entries[]",
  "archive_read::password",
  // diagram_inspect
  "diagram_inspect::source",
  // diagram_render
  "diagram_render::background",
  "diagram_render::fontFamily",
  "diagram_render::source",
  "diagram_render::themeOverrides.background",
  "diagram_render::themeOverrides.edge",
  "diagram_render::themeOverrides.edgeLabelBackground",
  "diagram_render::themeOverrides.edgeText",
  "diagram_render::themeOverrides.groupFill",
  "diagram_render::themeOverrides.groupStroke",
  "diagram_render::themeOverrides.nodeFill",
  "diagram_render::themeOverrides.nodeStroke",
  "diagram_render::themeOverrides.nodeText",
  "diagram_render::themeOverrides.paletteText",
  "diagram_render::themeOverrides.palette[]",
  "diagram_render::themeOverrides.title",
  // doc_convert
  "doc_convert::sheet",
  // doc_edit
  "doc_edit::find",
  "doc_edit::replace",
  // doc_read
  "doc_read::pages",
  // doc_search
  "doc_search::format.color",
  "doc_search::format.font",
  "doc_search::format.highlight",
  "doc_search::format.paragraphStyle",
  "doc_search::text",
  // doc_write
  "doc_write::markdown",
  // form_fill
  "form_fill::values{}",
  // formula_evaluate
  "formula_evaluate::cell",
  "formula_evaluate::contextFormulas{}",
  "formula_evaluate::context{}",
  "formula_evaluate::formula",
  // pdf_edit
  "pdf_edit::ops[op=append].pages",
  "pdf_edit::ops[op=delete_pages].pages",
  "pdf_edit::ops[op=diagram].background",
  "pdf_edit::ops[op=diagram].pages",
  "pdf_edit::ops[op=diagram].source",
  "pdf_edit::ops[op=keep_pages].pages",
  "pdf_edit::ops[op=page_numbers].format",
  "pdf_edit::ops[op=rotate].pages",
  "pdf_edit::ops[op=stamp].color",
  "pdf_edit::ops[op=stamp].pages",
  "pdf_edit::ops[op=stamp].text",
  "pdf_edit::ops[op=watermark].color",
  "pdf_edit::ops[op=watermark].pages",
  "pdf_edit::ops[op=watermark].text",
  // sheet_edit
  "sheet_edit::ops[op=add_chart].chart.anchor",
  "sheet_edit::ops[op=add_chart].chart.categories",
  "sheet_edit::ops[op=add_chart].chart.seriesNames[]",
  "sheet_edit::ops[op=add_chart].chart.title",
  "sheet_edit::ops[op=add_chart].chart.values",
  "sheet_edit::ops[op=add_chart].chart.values[]",
  "sheet_edit::ops[op=add_image].altText",
  "sheet_edit::ops[op=add_image].at",
  "sheet_edit::ops[op=add_image].background",
  "sheet_edit::ops[op=add_image].source",
  "sheet_edit::ops[op=add_sheet].name",
  "sheet_edit::ops[op=add_sheet].rows[][]",
  "sheet_edit::ops[op=clear].range",
  "sheet_edit::ops[op=insert_rows].rows[][]",
  "sheet_edit::ops[op=set_cell].ref",
  "sheet_edit::ops[op=set_cell].value",
  "sheet_edit::ops[op=set_formula].formula",
  "sheet_edit::ops[op=set_formula].range",
  "sheet_edit::ops[op=set_range].range",
  "sheet_edit::ops[op=set_range].rows[][]",
  "sheet_edit::ops[op=set_style].fillColor",
  "sheet_edit::ops[op=set_style].fontColor",
  "sheet_edit::ops[op=set_style].numFmt",
  "sheet_edit::ops[op=set_style].range",
  "sheet_edit::sheet",
  // sheet_read
  "sheet_read::range",
  "sheet_read::sheet",
  // sheet_write
  "sheet_write::sheets[].cells{}",
  "sheet_write::sheets[].charts[].anchor",
  "sheet_write::sheets[].charts[].categories",
  "sheet_write::sheets[].charts[].seriesNames[]",
  "sheet_write::sheets[].charts[].title",
  "sheet_write::sheets[].charts[].values",
  "sheet_write::sheets[].charts[].values[]",
  "sheet_write::sheets[].csvDelimiter",
  "sheet_write::sheets[].formulas{}",
  "sheet_write::sheets[].generate.columns[].name",
  "sheet_write::sheets[].generate.columns[].values[]",
  "sheet_write::sheets[].images[].altText",
  "sheet_write::sheets[].images[].at",
  "sheet_write::sheets[].images[].background",
  "sheet_write::sheets[].images[].source",
  "sheet_write::sheets[].merges[]",
  "sheet_write::sheets[].name",
  "sheet_write::sheets[].rows[][]",
  "sheet_write::sheets[].styles[].range",
  "sheet_write::sheets[].styles[].style.fillColor",
  "sheet_write::sheets[].styles[].style.fontColor",
  "sheet_write::sheets[].styles[].style.numFmt",
  // template_fill
  "template_fill::data{}",
  "template_fill::images{}.altText",
  "template_fill::images{}.background",
  "template_fill::images{}.source"
]);

// ---------------------------------------------------------------------------
// Fixture
// ---------------------------------------------------------------------------

interface Fixture {
  readonly config: ServerConfig;
  readonly root: string;
  /** A directory outside both roots, holding a file the model must never read. */
  readonly outside: string;
  /** Absolute path to that file. */
  readonly secret: string;
}

async function makeFixture(): Promise<Fixture> {
  const root = await realpath(await mkdtemp(path.join(tmpdir(), "documonster-mcp-boundary-")));
  const outside = await realpath(await mkdtemp(path.join(tmpdir(), "documonster-mcp-outside-")));
  const secret = path.join(outside, "secret.md");
  await writeFile(secret, "# TOP SECRET\n", "utf8");
  await writeFile(path.join(outside, "secret.csv"), "a,b\nTOP,SECRET\n", "utf8");

  // Valid in-root inputs, so a tool reaches its path check instead of failing
  // earlier on the file's type or on having nothing to do. Both happened while
  // writing this file: `doc_edit` returns a no-op before resolving `out` when
  // the search matches nothing, and `template_fill` refuses a template with no
  // placeholders before resolving its images — in each case the path check is
  // never reached and the call looks confined without having been tested.
  await writeFile(path.join(root, "a.csv"), "a,b\n1,2\n", "utf8");
  const doc = Document.create();
  Document.addParagraph(doc, "alpha beta gamma");
  Document.addParagraph(doc, "{{%logo}}");
  await Io.writeFile(Document.build(doc), path.join(root, "real.docx"));
  const builder = new Pdf.Builder();
  builder.addPage().drawText("hello", { x: 60, y: 700 });
  await writeFile(path.join(root, "real.pdf"), await builder.build());
  const zip = new ArchiveFile();
  zip.addText("payload\n", "inner.txt");
  await zip.writeToFile(path.join(root, "real.zip"));
  const wb = Workbook.create();
  Workbook.addWorksheet(wb, "Data");
  await Workbook.writeFile(wb, path.join(root, "real.xlsx"));
  // `doc_review` reports "no tracked changes" and returns before resolving
  // `out`, so its destination is only reachable through a document that really
  // has revisions.
  const tracked = Document.create();
  Document.addContent(tracked, {
    type: "paragraph",
    children: [
      Build.text("Payment due in "),
      Build.deletedRun(Build.text("30"), { author: "Alice", id: 1 }),
      Build.insertedRun(Build.text("14"), { author: "Alice", id: 2 })
    ]
  });
  await Io.writeFile(Document.build(tracked), path.join(root, "tracked.docx"));

  // The secure default: a private output root, disjoint from the input root.
  return { config: resolveConfig(["--root", root]), root, outside, secret };
}

/**
 * One fixture for the whole file, created once and removed afterwards.
 *
 * It used to be built per test, which meant five input roots, five private
 * output roots and five sets of DOCX/PDF/ZIP/XLSX per run, none of them ever
 * deleted — and the injection test writes into its output root a hundred more
 * times. The two structural assertions need no files at all and no longer ask
 * for any.
 */
let shared: Fixture;

beforeAll(async () => {
  shared = await makeFixture();
}, 60_000);

afterAll(async () => {
  await Promise.all([
    rm(shared.root, { recursive: true, force: true }),
    rm(shared.outside, { recursive: true, force: true }),
    rm(shared.config.outputRoot, { recursive: true, force: true })
  ]);
});

// ---------------------------------------------------------------------------
// Cases
// ---------------------------------------------------------------------------

interface BoundaryCase {
  /** The leaf this exercises, as `tool::accessor`. Must match the inventory. */
  readonly carrier: string;
  readonly args: Record<string, unknown>;
  /** Defaults to `outside_root`. */
  readonly code?: ToolErrorCode;
  /**
   * Substring the refusal must contain — always stated, never inferred.
   * Guessing it from the arguments picked the *innocent* argument for
   * `doc_convert`, and so asserted that a refusal naming the destination named
   * the source instead.
   */
  readonly mentions: string;
}

const SECRET_MENTION = "secret.md";

function casesFor(fx: Fixture): readonly BoundaryCase[] {
  const outside = fx.secret;
  const outsideCsv = path.join(fx.outside, "secret.csv");
  return [
    // --- reads -------------------------------------------------------------
    {
      carrier: "doc_inspect::path",
      args: { path: outside },
      mentions: SECRET_MENTION
    },
    {
      carrier: "doc_read::path",
      args: { path: outside },
      mentions: SECRET_MENTION
    },
    {
      carrier: "sheet_read::path",
      args: { path: outsideCsv },
      mentions: "secret.csv"
    },
    {
      carrier: "doc_search::path",
      args: { path: outside, text: "SECRET" },
      mentions: SECRET_MENTION
    },
    {
      carrier: "template_inspect::path",
      args: { path: outside },
      mentions: SECRET_MENTION
    },
    {
      carrier: "diagram_inspect::from",
      args: { from: outside },
      mentions: SECRET_MENTION
    },

    // --- edits: source and destination ------------------------------------
    {
      carrier: "doc_edit::path",
      args: { path: outside, find: "a", replace: "b" },
      mentions: SECRET_MENTION
    },
    {
      carrier: "doc_edit::out",
      args: { path: "real.docx", find: "alpha", replace: "b", out: "../escape.docx" },
      mentions: "../escape.docx"
    },
    {
      carrier: "doc_paginate::path",
      args: { path: outside },
      mentions: SECRET_MENTION
    },
    {
      carrier: "doc_paginate::out",
      args: { path: "real.docx", updateFields: true, out: "../escape.docx" },
      mentions: "../escape.docx"
    },
    {
      carrier: "doc_review::path",
      args: { path: outside },
      mentions: SECRET_MENTION
    },
    {
      carrier: "doc_review::against",
      args: { path: "real.docx", against: outside },
      mentions: SECRET_MENTION
    },
    {
      carrier: "doc_review::out",
      args: { path: "tracked.docx", apply: "accept-all", out: "../escape.docx" },
      mentions: "../escape.docx"
    },
    {
      carrier: "form_fill::path",
      args: { path: outside },
      mentions: SECRET_MENTION
    },
    {
      carrier: "form_fill::out",
      args: { path: "real.docx", values: { a: "b" }, out: "../escape.docx" },
      mentions: "../escape.docx"
    },
    {
      carrier: "sheet_edit::path",
      args: { path: outside, ops: [{ op: "set_cell", ref: "A1", value: 1 }] },
      mentions: SECRET_MENTION
    },
    {
      carrier: "sheet_edit::out",
      args: { path: "real.xlsx", out: "../escape.xlsx", ops: [] },
      mentions: "../escape.xlsx"
    },
    {
      // Nested, and untested until the schema walk found it: an image op reads
      // a file of its own.
      carrier: "sheet_edit::ops[op=add_image].from",
      args: {
        path: "real.xlsx",
        out: "img.xlsx",
        ops: [{ op: "add_image", at: "A1", from: outside }]
      },
      mentions: SECRET_MENTION
    },
    {
      carrier: "pdf_edit::path",
      args: { path: outside, ops: [{ op: "rotate", degrees: 90 }] },
      mentions: SECRET_MENTION
    },
    {
      carrier: "pdf_edit::out",
      args: { path: "real.pdf", ops: [{ op: "rotate", degrees: 90 }], out: "../escape.pdf" },
      mentions: "../escape.pdf"
    },
    {
      // The outer path is a genuine in-root PDF, so the handler reaches the op
      // loop and only the nested path is hostile.
      carrier: "pdf_edit::ops[op=append].path",
      args: { path: "real.pdf", ops: [{ op: "append", path: outside }], out: "merged.pdf" },
      mentions: SECRET_MENTION
    },
    {
      // Nested, and untested until the schema walk found it: a diagram op reads
      // its source from a file.
      carrier: "pdf_edit::ops[op=diagram].from",
      args: { path: "real.pdf", ops: [{ op: "diagram", from: outside }], out: "diag.pdf" },
      mentions: SECRET_MENTION
    },

    // --- writes ------------------------------------------------------------
    {
      carrier: "doc_write::path",
      args: { path: "../escape.docx", markdown: "# x" },
      mentions: "../escape.docx"
    },
    {
      carrier: "doc_convert::from",
      args: { from: outside, to: "out.md" },
      mentions: SECRET_MENTION
    },
    {
      carrier: "doc_convert::to",
      args: { from: "a.csv", to: "../escape.xlsx" },
      mentions: "../escape.xlsx"
    },
    {
      carrier: "sheet_write::path",
      args: { path: "../escape.xlsx", sheets: [] },
      mentions: "../escape.xlsx"
    },
    {
      carrier: "sheet_write::sheets[].fromCsv",
      args: { path: "ok.xlsx", sheets: [{ name: "S", fromCsv: outsideCsv }] },
      mentions: "secret.csv"
    },
    {
      // Nested two levels down, and untested until the schema walk found it.
      carrier: "sheet_write::sheets[].images[].from",
      args: {
        path: "img2.xlsx",
        sheets: [{ name: "S", images: [{ at: "A1", from: outside }] }]
      },
      mentions: SECRET_MENTION
    },
    {
      carrier: "diagram_render::to",
      args: { source: "flowchart LR\n A-->B", to: "../escape.svg" },
      mentions: "../escape.svg"
    },
    {
      carrier: "diagram_render::from",
      args: { from: outside, to: "d.svg" },
      mentions: SECRET_MENTION
    },
    {
      carrier: "template_fill::template",
      args: { template: outside, out: "t.docx", data: {} },
      mentions: SECRET_MENTION
    },
    {
      carrier: "template_fill::out",
      args: { template: "real.docx", out: "../escape.docx", data: {} },
      mentions: "../escape.docx"
    },
    {
      carrier: "template_fill::images{}.from",
      args: {
        template: "real.docx",
        out: "t2.docx",
        data: {},
        images: { logo: { from: outside } }
      },
      mentions: SECRET_MENTION
    },

    // --- archives ----------------------------------------------------------
    {
      carrier: "archive_read::path",
      args: { path: outside },
      mentions: SECRET_MENTION
    },
    {
      carrier: "archive_read::out",
      args: { path: "real.zip", action: "extract", out: "../escape" },
      mentions: "../escape"
    },
    {
      carrier: "archive_write::out",
      args: { out: "../escape.zip", entries: [] },
      mentions: "../escape.zip"
    },
    {
      carrier: "archive_write::entries[].path",
      args: { out: "z.zip", entries: [{ path: outside }] },
      mentions: SECRET_MENTION
    },
    {
      // Not a host path but a path all the same: the name this server writes
      // *into* the archive. Unchecked, this tool would produce a Zip Slip
      // payload for whoever opens the result.
      carrier: "archive_write::entries[].as",
      args: { out: "z2.zip", entries: [{ path: "a.csv", as: "../../evil.txt" }] },
      code: "invalid_input",
      mentions: "../../evil.txt"
    }
  ];
}

// ---------------------------------------------------------------------------
// Assertions
// ---------------------------------------------------------------------------

interface DiscriminatedBranch {
  readonly tool: string;
  readonly key: string;
  readonly value: string;
}

/** Every discriminated union branch in every tool, read from the schemas. */
function discriminatedBranches(): DiscriminatedBranch[] {
  const found: DiscriminatedBranch[] = [];
  const walk = (node: unknown, tool: string, depth = 0): void => {
    if (node === null || typeof node !== "object" || depth > 14) {
      return;
    }
    const schema = node as Record<string, unknown>;
    for (const keyword of ["anyOf", "oneOf", "allOf"] as const) {
      const branches = schema[keyword];
      if (!Array.isArray(branches)) {
        continue;
      }
      const key = discriminatorOf(branches);
      for (const branch of branches) {
        if (key !== undefined) {
          const value = (branch as { properties?: Record<string, { const?: unknown }> })
            .properties?.[key]?.const;
          if (typeof value === "string") {
            found.push({ tool, key, value });
          }
        }
        walk(branch, tool, depth + 1);
      }
    }
    for (const child of Object.values(schema.properties ?? {})) {
      walk(child, tool, depth + 1);
    }
    walk(schema.items, tool, depth + 1);
    walk(schema.additionalProperties, tool, depth + 1);
  };
  for (const tool of ALL_TOOLS) {
    walk(
      z.toJSONSchema(z.object(tool.inputSchema), { io: "input", unrepresentable: "any" }),
      tool.name
    );
  }
  return found;
}

function toolOf(carrier: string): AnyToolDefinition {
  const name = carrier.split("::")[0] ?? "";
  const tool = ALL_TOOLS.find(candidate => candidate.name === name);
  if (tool === undefined) {
    throw new Error(`carrier "${carrier}" names no registered tool`);
  }
  return tool;
}

function textOf(result: { content: readonly unknown[] }): string {
  return result.content
    .filter((block): block is { type: "text"; text: string } => {
      const candidate = block as { type?: string; text?: unknown };
      return candidate.type === "text" && typeof candidate.text === "string";
    })
    .map(block => block.text)
    .join("\n");
}

/**
 * The text a model would see for a hostile call.
 *
 * A handler signals a sandbox refusal by *throwing* `McpToolError`; the server
 * layer renders it. The rendering is done here with that same production
 * function, rather than asserting on a shape the model never receives.
 */
async function replyFor(
  tool: AnyToolDefinition,
  args: Record<string, unknown>,
  config: ServerConfig
): Promise<{ text: string; refused: boolean }> {
  try {
    const result = await tool.handler(args, { config });
    return { text: textOf(result), refused: result.isError === true };
  } catch (cause) {
    return { text: formatToolError(cause), refused: true };
  }
}

// ---------------------------------------------------------------------------
// Verifying the "not a path" claims
// ---------------------------------------------------------------------------

/**
 * A benign, valid call per tool, used as the carrier for an injected value.
 *
 * It has to be a call that really does the work: injecting into a request the
 * tool rejects up front proves nothing, which is the same trap the three
 * early-return cases above fell into.
 */
function baseArgsFor(tool: string): Record<string, unknown> | undefined {
  switch (tool) {
    case "doc_inspect":
      return { path: "a.csv" };
    case "doc_read":
      return { path: "a.csv" };
    case "doc_search":
      return { path: "real.docx", text: "alpha" };
    case "doc_edit":
      return { path: "real.docx", find: "alpha", replace: "beta", out: "e.docx", overwrite: true };
    case "doc_write":
      return { path: "w.docx", markdown: "# x", overwrite: true };
    case "doc_convert":
      return { from: "a.csv", to: "c.xlsx", overwrite: true };
    case "doc_paginate":
      return { path: "real.docx" };
    case "doc_review":
      return { path: "tracked.docx" };
    case "sheet_read":
      return { path: "real.xlsx" };
    case "sheet_write":
      return { path: "w.xlsx", sheets: [{ name: "S", rows: [["a"]] }], overwrite: true };
    case "sheet_edit":
      return { path: "real.xlsx", out: "e.xlsx", overwrite: true, ops: [] };
    case "pdf_edit":
      return { path: "real.pdf", out: "e.pdf", overwrite: true, ops: [] };
    case "template_inspect":
      return { path: "real.docx" };
    case "template_fill":
      return {
        template: "real.docx",
        out: "t.docx",
        overwrite: true,
        data: {},
        images: { logo: { source: "flowchart LR\n A-->B" } }
      };
    case "form_fill":
      return { path: "real.docx" };
    case "formula_evaluate":
      return { formula: "1+1" };
    case "diagram_inspect":
      return { source: "flowchart LR\n A-->B" };
    case "diagram_render":
      return { source: "flowchart LR\n A-->B", to: "d.svg", overwrite: true };
    case "archive_read":
      return { path: "real.zip" };
    case "archive_write":
      return { out: "w.zip", entries: [{ path: "a.csv" }], overwrite: true };
    default:
      return undefined;
  }
}

/**
 * A minimal *valid* object for one branch of a discriminated union, keyed
 * `tool::<discriminator value>`.
 *
 * This is what makes an injection into a union member mean anything. With a
 * single fixed op per tool, a value put in `ops[].source` arrived on an
 * `op: "rotate"` object — which reads `degrees` and `pages` and nothing else —
 * so the assertion passed while exercising no code at all. That was true of 22
 * of the claims in `NON_PATH_LEAVES`. Each branch now carries the siblings it
 * needs to reach its own handler.
 */
const BRANCH_BASES: Record<string, Record<string, unknown>> = {
  "pdf_edit::watermark": { op: "watermark", text: "DRAFT" },
  "pdf_edit::page_numbers": { op: "page_numbers" },
  "pdf_edit::stamp": { op: "stamp", text: "S", x: 20, y: 20 },
  "pdf_edit::diagram": { op: "diagram", source: "flowchart LR\n A-->B" },
  "pdf_edit::delete_pages": { op: "delete_pages", pages: "1" },
  "pdf_edit::keep_pages": { op: "keep_pages", pages: "1" },
  "pdf_edit::rotate": { op: "rotate", degrees: 90 },
  "pdf_edit::append": { op: "append", path: "real.pdf" },
  "sheet_edit::set_cell": { op: "set_cell", ref: "A1", value: 1 },
  "sheet_edit::set_range": { op: "set_range", range: "A1:B1", rows: [["a", "b"]] },
  "sheet_edit::set_formula": { op: "set_formula", range: "A1", formula: "1+1" },
  "sheet_edit::clear": { op: "clear", range: "A1" },
  "sheet_edit::insert_rows": { op: "insert_rows", at: 1, rows: [["a"]] },
  "sheet_edit::delete_rows": { op: "delete_rows", at: 1, count: 1 },
  "sheet_edit::set_style": { op: "set_style", range: "A1", bold: true },
  "sheet_edit::add_sheet": { op: "add_sheet", name: "Added" },
  "sheet_edit::add_image": { op: "add_image", at: "A1", source: "flowchart LR\n A-->B" },
  "sheet_edit::add_chart": {
    op: "add_chart",
    chart: { type: "column", categories: "A1:A2", values: "B1:B2" }
  }
};

interface AccessorWrapper {
  readonly kind: "array" | "record";
  /** For a discriminated branch, the property and literal that select it. */
  readonly discriminator?: { readonly key: string; readonly value: string };
}

interface AccessorStep {
  readonly name: string;
  /** Containers to descend after the named property, outermost first. */
  readonly wrappers: readonly AccessorWrapper[];
}

/**
 * Arguments a schema offers as "this one or that one, not both".
 *
 * Without this an injection can be refused before the field is ever looked at:
 * `resolveDiagramSource` rejects `source` and `from` together
 * (`tools/diagram.ts`), so injecting `from` into a branch base that already
 * carries `source` produced an `invalid_input` about the pair and never reached
 * the path resolution the assertion is about. The exclusive partner is removed
 * from the same object as the injected leaf.
 */
const EXCLUSIVE_SIBLINGS: Record<string, readonly string[]> = {
  source: ["from"],
  from: ["source"]
};

const WRAPPER_RE = /\[\]|\{\}|\[([^\]=]+)=([^\]]+)\]/g;

function parseAccessor(accessor: string): AccessorStep[] {
  return accessor.split(".").map(raw => {
    const boundary = raw.search(/[[{]/);
    const name = boundary === -1 ? raw : raw.slice(0, boundary);
    const wrappers: AccessorWrapper[] = [];
    for (const match of raw.slice(name.length).matchAll(WRAPPER_RE)) {
      if (match[0] === "[]") {
        wrappers.push({ kind: "array" });
      } else if (match[0] === "{}") {
        wrappers.push({ kind: "record" });
      } else {
        wrappers.push({
          kind: "array",
          discriminator: { key: match[1] ?? "", value: match[2] ?? "" }
        });
      }
    }
    return { name, wrappers };
  });
}

/** Remove the arguments `name` may not appear alongside. */
function dropExclusiveSiblings(holder: Record<string | number, unknown>, name: string): void {
  for (const sibling of EXCLUSIVE_SIBLINGS[name] ?? []) {
    delete holder[sibling];
  }
}

/**
 * Set the leaf an accessor names, creating containers on the way.
 *
 * Existing containers are reused — element `0` of an array, key `k` of a record
 * — so an injection lands *inside* the base call's own sheet and the surrounding
 * request stays valid. A `[key=value]` wrapper seeds the element from
 * {@link BRANCH_BASES} first, so the handler selects that branch and reads the
 * injected field.
 */
function injectAt(
  tool: string,
  base: Record<string, unknown>,
  accessor: string,
  value: string
): Record<string, unknown> {
  const clone = structuredClone(base);
  const steps = parseAccessor(accessor);
  let cursor: Record<string | number, unknown> = clone;

  for (const [index, step] of steps.entries()) {
    const lastStep = index === steps.length - 1;
    if (lastStep && step.wrappers.length === 0) {
      dropExclusiveSiblings(cursor, step.name);
      cursor[step.name] = value;
      return clone;
    }
    if (cursor[step.name] === null || typeof cursor[step.name] !== "object") {
      cursor[step.name] = step.wrappers[0]?.kind === "record" ? {} : [];
    }
    cursor = cursor[step.name] as Record<string | number, unknown>;

    for (const [depth, wrapper] of step.wrappers.entries()) {
      const key = wrapper.kind === "array" ? 0 : "k";
      const lastWrapper = depth === step.wrappers.length - 1;

      if (wrapper.discriminator !== undefined) {
        const branchKey = `${tool}::${wrapper.discriminator.value}`;
        const branch = BRANCH_BASES[branchKey];
        if (branch === undefined) {
          throw new Error(`no BRANCH_BASES entry for ${branchKey}`);
        }
        cursor[key] = structuredClone(branch);
      } else if (cursor[key] === null || typeof cursor[key] !== "object") {
        cursor[key] = step.wrappers[depth + 1]?.kind === "record" ? {} : [];
      }

      if (lastStep && lastWrapper) {
        // The leaf *is* the container element (`rows[][]`, `entries[]`).
        cursor[key] = value;
        return clone;
      }
      cursor = cursor[key] as Record<string | number, unknown>;
    }
  }
  return clone;
}

describe("path boundary", () => {
  it("classifies every string argument in every tool", async () => {
    // The drift guard, and the reason this file walks the schemas at all: a new
    // string argument is either declared not to be a path or boundary-tested.
    // Nothing can be added silently, whatever it is called.
    const carriers = new Set(casesFor(shared).map(testCase => testCase.carrier));
    const unclassified = inventory().filter(
      leaf => !NON_PATH_LEAVES.has(leaf) && !carriers.has(leaf)
    );
    expect(
      unclassified,
      "add a boundary case, or list in NON_PATH_LEAVES if it names no file"
    ).toEqual([]);
  });

  it("keeps both classifications anchored to the real schemas", async () => {
    // A stale entry on either side is a silent hole: a case for a leaf that no
    // longer exists stops testing anything, and a NON_PATH_LEAVES entry for a
    // renamed field excuses nothing while looking like it does.
    const known = new Set(inventory());
    expect([...NON_PATH_LEAVES].filter(leaf => !known.has(leaf))).toEqual([]);
    expect(
      casesFor(shared)
        .map(c => c.carrier)
        .filter(leaf => !known.has(leaf))
    ).toEqual([]);
  });

  it("gives every discriminated union branch a base that selects it", () => {
    // Without this the branch bases are unverified data, and a wrong one makes
    // an injection silently vacuous rather than failing: swapping the diagram
    // base's `op` for `rotate` left every claim about the diagram branch
    // passing while exercising nothing.
    const branches = discriminatedBranches();
    expect(branches.length).toBeGreaterThan(0);

    const missing: string[] = [];
    const wrong: string[] = [];
    for (const branch of branches) {
      const key = `${branch.tool}::${branch.value}`;
      const base = BRANCH_BASES[key];
      if (base === undefined) {
        missing.push(key);
        continue;
      }
      if (base[branch.key] !== branch.value) {
        wrong.push(`${key}: base has ${branch.key}=${JSON.stringify(base[branch.key])}`);
      }
    }
    expect(missing, "add a BRANCH_BASES entry").toEqual([]);
    expect(wrong, "the base must select the branch it is keyed by").toEqual([]);

    const known = new Set(branches.map(branch => `${branch.tool}::${branch.value}`));
    expect([...Object.keys(BRANCH_BASES)].filter(key => !known.has(key))).toEqual([]);
  });

  it("confines every hostile path", async () => {
    const fx = shared;
    const failures: string[] = [];

    for (const testCase of casesFor(fx)) {
      const expectedCode = testCase.code ?? "outside_root";
      const { text, refused } = await replyFor(toolOf(testCase.carrier), testCase.args, fx.config);

      if (!refused) {
        failures.push(`${testCase.carrier}: succeeded, expected ${expectedCode}`);
        continue;
      }
      if (!text.includes(`[${expectedCode}]`)) {
        failures.push(`${testCase.carrier}: expected [${expectedCode}], got ${text}`);
        continue;
      }
      // Matching the code alone is not enough: a tool that refused the call for
      // an unrelated reason would look confined. This half of the assertion
      // caught three cases in this very table that were rejected before their
      // path was ever examined.
      if (!text.includes(testCase.mentions)) {
        failures.push(
          `${testCase.carrier}: rejected, but the message never names ${testCase.mentions}: ${text}`
        );
      }
    }

    expect(failures).toEqual([]);
  });

  it("never lets content from outside the root reach the model", async () => {
    // The property that actually matters, asserted on the bytes rather than on
    // an error code: no reply may contain the secret file's contents.
    const fx = shared;
    for (const testCase of casesFor(fx)) {
      const { text } = await replyFor(toolOf(testCase.carrier), testCase.args, fx.config);
      expect(text, testCase.carrier).not.toContain("TOP SECRET");
    }
  });
  describe("stringLeaves", () => {
    // Asserted on synthetic schemas because no current tool produces these
    // shapes — which is exactly why a regression in the guards would otherwise
    // be invisible until a tool did.
    it("returns free-form strings and skips closed value sets", () => {
      expect(stringLeaves({ type: "string" }, "a")).toEqual(["a"]);
      expect(stringLeaves({ type: "string", enum: ["x", "y"] }, "a")).toEqual([]);
      expect(stringLeaves({ type: "string", const: "x" }, "a")).toEqual([]);
      expect(stringLeaves({ type: "number" }, "a")).toEqual([]);
    });

    it("treats an unconstrained schema as a possible string", () => {
      expect(stringLeaves({}, "a")).toEqual(["a"]);
      expect(stringLeaves({ description: "anything" }, "a")).toEqual(["a"]);
    });

    it("qualifies a discriminated union and leaves a plain one alone", () => {
      const discriminated = {
        oneOf: [
          {
            type: "object",
            properties: { op: { type: "string", const: "one" }, v: { type: "string" } }
          },
          {
            type: "object",
            properties: { op: { type: "string", const: "two" }, w: { type: "string" } }
          }
        ]
      };
      expect(stringLeaves(discriminated, "ops[]")).toEqual(["ops[op=one].v", "ops[op=two].w"]);
      expect(
        stringLeaves(
          { anyOf: [{ type: "string" }, { type: "array", items: { type: "number" } }] },
          "p"
        )
      ).toEqual(["p"]);
    });

    it("throws rather than shrinking the inventory", () => {
      // Each of these used to return nothing, which is indistinguishable from
      // "this schema has no paths in it".
      expect(() => stringLeaves({ $ref: "#/$defs/X" }, "a")).toThrow(/does not handle/);
      expect(() => stringLeaves({ type: "object" }, "a")).toThrow(/unconstrained object/);
      expect(() => stringLeaves({ type: "array" }, "a")).toThrow(/unconstrained array/);
      expect(() =>
        stringLeaves({ type: "object", properties: {}, anyOf: [{ type: "string" }] }, "a")
      ).toThrow(/beside a sibling constraint/);
      expect(() =>
        stringLeaves({ anyOf: [{ type: "string" }], oneOf: [{ type: "string" }] }, "a")
      ).toThrow(/beside a sibling constraint/);
      expect(() => stringLeaves(null, "a")).toThrow(/not a JSON Schema object/);
    });
  });

  it("injecting a path into a declared non-path argument neither reads nor resolves it", async () => {
    // This is what turns NON_PATH_LEAVES from a list of assertions into a list
    // of checked facts. Every claimed non-path argument is handed the absolute
    // path of a file outside the root, inside an otherwise valid call.
    //
    // Two things must hold. Nothing from that file may come back — the property
    // the sandbox exists for. And the call must not be refused as
    // `outside_root`: only a *resolved path* produces that code, so seeing it
    // would mean the field is a carrier and belongs in CASES. Any other error
    // is fine and expected — a path is not a cell reference or a formula.
    const fx = shared;
    const leaks: string[] = [];
    const misclassified: string[] = [];
    let exercised = 0;

    for (const leaf of inventory()) {
      if (!NON_PATH_LEAVES.has(leaf)) {
        continue;
      }
      const [toolName = "", accessor = ""] = leaf.split("::");
      const base = baseArgsFor(toolName);
      if (base === undefined) {
        throw new Error(`no base call defined for ${toolName}; add one to baseArgsFor`);
      }
      const tool = ALL_TOOLS.find(candidate => candidate.name === toolName);
      if (tool === undefined) {
        throw new Error(`${leaf} names no registered tool`);
      }

      const { text } = await replyFor(
        tool,
        injectAt(toolName, base, accessor, fx.secret),
        fx.config
      );
      exercised += 1;
      if (text.includes("TOP SECRET")) {
        leaks.push(`${leaf}: returned the contents of a file outside the root`);
      }
      if (text.includes("[outside_root]")) {
        misclassified.push(`${leaf}: resolved as a path — move it to CASES`);
      }
    }

    expect(leaks).toEqual([]);
    expect(misclassified).toEqual([]);
    // Guards against the loop silently doing nothing.
    expect(exercised).toBe(NON_PATH_LEAVES.size);
  });
});
