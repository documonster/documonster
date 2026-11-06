/**
 * A wrong argument to a public function is the module's own error, never a `TypeError` from
 * inside.
 *
 * Every module states this ("Arguments are checked at the boundary"), and each has tests for
 * the checks it knows it has. What those cannot show is the check nobody wrote: one sweep with
 * deliberately wrong arguments over every namespace member of every entry found about seventy
 * members that let a raw `TypeError` escape — `Workbook.toBuffer()`, `Font.createHyphenator({})`,
 * the whole of `Fs` — and two that accepted `{}` as a worksheet and wrote onto it. This is that
 * sweep, kept, so the next member added without a check fails here rather than in a caller.
 *
 * It calls each function — and constructs each class — with a handful of wrong argument lists
 * (each call gets fresh objects, so one member writing onto a shared `{}` cannot change what the
 * next one sees), and requires any throw or rejection to be reported properly: one of the
 * library's errors, of the entry's own module (or of the module a bridge hands its input to
 * unchanged), and, for an argument error, naming the argument in `field`. A call that accepts
 * the arguments is not a failure: many members have no required argument.
 *
 * What it does not reach, and the module's own tests must: the methods of an instance (it would
 * need a valid one of every class), a wrong argument after a valid handle in every position, and
 * whether a refused call left its target unchanged.
 *
 * A member listed in `EXEMPT` keeps its behaviour for a stated reason. The list is checked
 * too: an entry naming a member that no longer exists, or that no longer leaks, fails, so the
 * list cannot outlive what it excuses.
 */
import * as archive from "@archive/index";
import * as csv from "@csv/index";
import * as draw from "@draw/index";
import * as excelCsv from "@excel/bridge/csv.node";
import * as excelFormula from "@excel/bridge/formula";
import * as excelMarkdown from "@excel/bridge/markdown.node";
import * as excelChart from "@excel/chart/index";
import * as excel from "@excel/index";
import * as formula from "@formula/index";
import * as markdown from "@markdown/index";
import * as mermaid from "@mermaid/index";
import * as pdf from "@pdf/index";
import * as stream from "@stream/index";
import { BaseError } from "@utils/errors";
import * as wordCrypto from "@word/crypto";
import * as wordExcel from "@word/excel";
import * as wordHtml from "@word/html";
import * as word from "@word/index";
import * as wordMarkdown from "@word/markdown";
import * as xml from "@xml/index";
import { describe, expect, it } from "vitest";

type ErrorGuard = (error: unknown) => boolean;

const isAbort: ErrorGuard = error => (error as { name?: unknown } | null)?.name === "AbortError";

/**
 * Each entry, and the errors it may throw: its own module's, plus those of the module a bridge
 * hands its input to unchanged (text the CSV parser cannot read stays a `CsvError` through
 * `ExcelCsv`). Anything else — another module's argument error escaping through a bridge — is
 * the failure `WordExcel.toDocx` had, when a wrong workbook came out as `ExcelArgumentError`.
 */
const ENTRIES: Record<
  string,
  { readonly entry: Record<string, unknown>; readonly owns: ErrorGuard }
> = {
  "documonster/excel": { entry: excel, owns: excel.isExcelError },
  "documonster/chart": { entry: excelChart, owns: e => excel.isExcelError(e) || xml.isXmlError(e) },
  "documonster/excel/csv": {
    entry: excelCsv,
    owns: e => excel.isExcelError(e) || csv.isCsvError(e)
  },
  "documonster/excel/markdown": {
    entry: excelMarkdown,
    owns: e => excel.isExcelError(e) || markdown.isMarkdownError(e)
  },
  "documonster/excel/formula": { entry: excelFormula, owns: excel.isExcelError },
  "documonster/word": { entry: word, owns: word.isDocxError },
  "documonster/word/html": { entry: wordHtml, owns: word.isDocxError },
  "documonster/word/markdown": { entry: wordMarkdown, owns: word.isDocxError },
  "documonster/word/excel": { entry: wordExcel, owns: word.isDocxError },
  "documonster/word/crypto": { entry: wordCrypto, owns: word.isDocxError },
  "documonster/pdf": { entry: pdf, owns: pdf.isPdfError },
  "documonster/draw": { entry: draw, owns: draw.isDrawError },
  "documonster/mermaid": { entry: mermaid, owns: mermaid.isMermaidError },
  "documonster/csv": { entry: csv, owns: csv.isCsvError },
  "documonster/markdown": { entry: markdown, owns: markdown.isMarkdownError },
  "documonster/xml": { entry: xml, owns: xml.isXmlError },
  "documonster/formula": { entry: formula, owns: formula.isFormulaError },
  "documonster/archive": { entry: archive, owns: e => archive.isArchiveError(e) || isAbort(e) },
  "documonster/stream": { entry: stream, owns: e => stream.isStreamError(e) || isAbort(e) }
};

/**
 * Members whose leak is deliberate. A deprecated forward keeps the behaviour it was published
 * with (`AGENTS.md`, "Renaming a public name"), including a `TypeError` for a wrong argument;
 * its replacement is checked.
 */
const DEPRECATED =
  "a deprecated forward: it keeps its published behaviour, and its replacement checks";
const EXEMPT: Record<string, string> = {
  "Build.drawingShape": DEPRECATED,
  "Query.collectTables": DEPRECATED,
  "Query.resolveTableStyle": DEPRECATED,
  "Template.fillTemplate": DEPRECATED,
  "Template.fillTemplateEnhanced": DEPRECATED,
  "Template.fillTemplateFromSource": DEPRECATED,
  "Template.listTemplateTags": DEPRECATED,
  "Csv.rowHashArrayMapByHeaders": DEPRECATED,
  "Csv.rowHashArrayToHeaders": DEPRECATED,
  "Csv.rowHashArrayToValues": DEPRECATED,
  "Xml.attr": DEPRECATED,
  "Xml.findChildren": DEPRECATED,
  "Xml.textContent": DEPRECATED,
  "Xml.walk": DEPRECATED,
  "Archive.editZip": DEPRECATED,
  "Archive.editZipUrl": DEPRECATED,
  "Archive.unzip": DEPRECATED,
  "Archive.zip": DEPRECATED,
  "Archive.zipCryptoDecrypt": DEPRECATED,
  "Archive.zipCryptoInitKeys": DEPRECATED,
  // A path that names no file is the file system's refusal, and `Fs.*` pass that on as it is.
  "Fs.readFileBytes": "the file system's refusal of a missing file, which Fs passes on",
  "Fs.readFileBytesSync": "the file system's refusal of a missing file, which Fs passes on",
  "Fs.readFileText": "the file system's refusal of a missing file, which Fs passes on",
  "Fs.readFileTextSync": "the file system's refusal of a missing file, which Fs passes on",
  // `node:stream` semantics: a readable side with no `read` errors when it is read.
  "Streams.createDuplex": "node:stream parity — `_read() is not implemented`",
  "Streams.createReadable": "node:stream parity — `_read() is not implemented`",
  // Content that is not JSON is a `SyntaxError`, as `Response.json()` and `JSON.parse` give.
  "Streams.json": "content that is not JSON is a SyntaxError, as Response.json() gives",
  // A class whose constructor is private in its declaration is opened with its static opener
  // (`Pdf.Editor.read`, `Archive.RemoteZipReader.open`, `Archive.ZipEditor.open`); calling it with
  // `new` compiles in no consumer's code.
  "Pdf.Editor": "a private constructor; the class is opened with Pdf.Editor.read",
  "Archive.RemoteZipReader": "a private constructor; the class is opened with RemoteZipReader.open",
  "Archive.ZipEditor": "a private constructor; the class is opened with ZipEditor.open",
  "Archive.UnzipEntry": DEPRECATED
};

/**
 * Members not called at all: each touches the file system or the network with whatever it is
 * given, so a sweep that passes `"x"` would create files and directories. Their argument checks
 * are covered by the module's own tests.
 */
const NOT_CALLED =
  /^(write|ensure|empty|copy|move|remove|rm|mkdir|createTemp|createWriteStream|streamTo|download|fetch|open[A-Z]?.*Url|readUrl)/;

/** A public member, called with whatever the sweep hands it. */
type AnyFunction = (...args: unknown[]) => unknown;

/** Wrong argument lists, built fresh for every call. */
function wrongArguments(): unknown[][] {
  return [[], [null], [{}], [42], ["x"], [{}, {}], [{}, null], [{}, 42], [{}, "x", 42]];
}

function isNamespace(value: unknown): value is Record<string, unknown> {
  return (
    typeof value === "object" &&
    value !== null &&
    Object.prototype.toString.call(value) === "[object Module]"
  );
}

/** An object the sweep walks as it walks an entry's `export * as` namespace. */
function makeNamespace(members: Record<string, unknown>): Record<string, unknown> {
  return Object.defineProperty({ ...members }, Symbol.toStringTag, { value: "Module" });
}

/** A class is constructed rather than called; both are swept, each in its own form. */
function isClass(fn: AnyFunction): boolean {
  return /^class\b/.test(Function.prototype.toString.call(fn)) || /^[A-Z]/.test(fn.name);
}

/**
 * What is wrong with how `error` reports a wrong argument, or `undefined` when nothing is: it
 * must be the library's, it must be the entry's own module's, and an argument error must say
 * which argument (`field`).
 */
function faultOf(error: unknown, owns: ErrorGuard): string | undefined {
  if (isAbort(error)) {
    return undefined;
  }
  if (!(error instanceof BaseError)) {
    return "not one of the library's errors";
  }
  if (!owns(error)) {
    return "another module's error";
  }
  const field = (error as { field?: unknown }).field;
  if (/ArgumentError$/.test(error.name) && (typeof field !== "string" || field === "")) {
    return "an argument error naming no field";
  }
  return undefined;
}

/** How a call's outcome is known: settled now, or still pending, with its outcome when it comes. */
interface Late {
  /** Set once the pending call settles; `error` is absent when it fulfilled. */
  outcome?: { readonly error?: unknown };
  readonly done: Promise<void>;
}

/** Keep watching `promise` past the quick look, so a rejection that arrives later is still seen. */
function watch(promise: Promise<unknown>): Late {
  const late: Late = {
    done: promise.then(
      () => {
        late.outcome = {};
      },
      (error: unknown) => {
        late.outcome = { error };
      }
    )
  };
  return late;
}

/**
 * Settle what a call returned, as far as a wrong argument would show: a rejection, or a throw on
 * first pull — of an async iterable, or of a synchronous iterator such as a generator's, whose body
 * (and so any check in it) runs only when it is pulled. A pulled iterator is closed again.
 *
 * A promise still pending after a moment is returned rather than waited for: many members wait for
 * input that never comes, and waiting for each would stall the sweep. But a pending promise is not
 * known to be fine — a file system's refusal (`Fs.readFileBytes("x")` is ENOENT) comes back through
 * the I/O thread pool, which on a loaded CI runner took longer than the moment, so the refusal was
 * missed and the check below reported `Fs.readFileBytes`'s exemption as stale. The caller waits for
 * all of an entry's pending calls together, once — see {@link sweep}.
 */
async function settle(result: unknown): Promise<Late | undefined> {
  const quickly = async (promise: Promise<unknown>): Promise<Late | undefined> => {
    const late = watch(promise);
    await Promise.race([late.done, new Promise(resolve => setTimeout(resolve, 10))]);
    if (late.outcome === undefined) {
      return late;
    }
    if (late.outcome.error !== undefined) {
      throw late.outcome.error;
    }
    return undefined;
  };
  if (result instanceof Promise) {
    return quickly(result);
  }
  if (
    result &&
    typeof (result as { [Symbol.asyncIterator]?: unknown })[Symbol.asyncIterator] === "function"
  ) {
    return quickly((result as AsyncIterable<unknown>)[Symbol.asyncIterator]().next());
  }
  if (result && typeof (result as { next?: unknown }).next === "function") {
    const it = result as Iterator<unknown>;
    try {
      it.next();
    } finally {
      it.return?.();
    }
  }
  return undefined;
}

type Leak = { args: unknown[]; error: unknown; fault: string };

/**
 * The first wrong argument list a member reports badly, with what was wrong — or, when no call
 * failed badly at once, the calls still pending, for {@link sweep} to look at once they settle.
 */
async function leakOf(
  fn: AnyFunction,
  construct: boolean,
  owns: ErrorGuard
): Promise<{ leak?: Leak; pending: { args: unknown[]; late: Late }[] }> {
  const pending: { args: unknown[]; late: Late }[] = [];
  for (const args of wrongArguments()) {
    try {
      const Ctor = fn as unknown as new (...a: unknown[]) => unknown;
      const late = await settle(construct ? new Ctor(...args) : fn(...args));
      if (late !== undefined) {
        pending.push({ args, late });
      }
    } catch (error) {
      const fault = faultOf(error, owns);
      if (fault !== undefined) {
        return { leak: { args, error, fault }, pending };
      }
    }
  }
  return { pending };
}

/**
 * How long an entry's pending calls are given, together, to settle. Every call that settles at all
 * does so in milliseconds on an idle machine; the bound only keeps the calls that wait for input
 * forever from stalling the sweep, and it is paid once per entry, not once per call.
 */
const PENDING_BOUND_MS = 3_000;

/** Each member of `entry` with the first wrong argument list it reports badly, if any. */
async function sweep(
  entry: Record<string, unknown>,
  owns: ErrorGuard
): Promise<{ name: string; construct: boolean; leak?: Leak; undecided?: boolean }[]> {
  const results: {
    name: string;
    construct: boolean;
    leak?: Leak;
    pending: { args: unknown[]; late: Late }[];
  }[] = [];
  for (const [name, fn, construct] of membersOf(entry)) {
    results.push({ name, construct, ...(await leakOf(fn, construct, owns)) });
  }
  const waiting = results.flatMap(result => (result.leak ? [] : result.pending));
  if (waiting.length > 0) {
    let timer: ReturnType<typeof setTimeout> | undefined;
    await Promise.race([
      Promise.all(waiting.map(({ late }) => late.done)),
      new Promise(resolve => {
        timer = setTimeout(resolve, PENDING_BOUND_MS);
      })
    ]);
    clearTimeout(timer);
  }
  return results.map(({ name, construct, leak, pending }) => {
    if (leak !== undefined) {
      return { name, construct, leak };
    }
    for (const { args, late } of pending) {
      const error = late.outcome?.error;
      const fault = error === undefined ? undefined : faultOf(error, owns);
      if (fault !== undefined) {
        return { name, construct, leak: { args, error, fault } };
      }
    }
    // A call still pending at the bound has not said how it fails, so it decides nothing.
    return { name, construct, undecided: pending.some(({ late }) => late.outcome === undefined) };
  });
}

/** Every function and class reachable through the namespaces an entry publishes, by dotted name. */
function membersOf(entry: Record<string, unknown>): [string, AnyFunction, boolean][] {
  const found: [string, AnyFunction, boolean][] = [];
  const walk = (prefix: string, ns: Record<string, unknown>): void => {
    for (const [key, value] of Object.entries(ns)) {
      const name = `${prefix}.${key}`;
      if (isNamespace(value)) {
        walk(name, value);
      } else if (typeof value === "function" && !NOT_CALLED.test(key)) {
        const fn = value as AnyFunction;
        found.push([name, fn, isClass(fn)]);
      }
    }
  };
  for (const [key, value] of Object.entries(entry)) {
    if (isNamespace(value)) {
      walk(key, value);
    }
  }
  return found;
}

const leaksByName = new Map<string, string>();
const seen = new Set<string>();
/** Members with a call still pending at the bound: whether they need their exemption is unknown. */
const undecided = new Set<string>();

describe("a wrong argument to a public function is the module's own error", () => {
  it.each(Object.entries(ENTRIES))(
    "%s",
    async (_name, { entry, owns }) => {
      const leaks: string[] = [];
      for (const { name, construct, leak, undecided: open } of await sweep(entry, owns)) {
        seen.add(name);
        if (open) {
          undecided.add(name);
        }
        if (!leak) {
          continue;
        }
        const error = leak.error as { name?: string; message?: string };
        const call = `${construct ? "new " : ""}${name}(${leak.args.map(a => JSON.stringify(a)).join(", ")})`;
        const described = `${call} → ${leak.fault}: ${error?.name}: ${error?.message}`;
        leaksByName.set(name, described);
        if (!Object.hasOwn(EXEMPT, name)) {
          leaks.push(described);
        }
      }
      expect(leaks).toEqual([]);
    },
    120_000
  );

  it("reaches a check a generator runs only when it is first pulled", async () => {
    function* checksLate(value?: unknown): Generator<number> {
      if (typeof value !== "number") {
        throw new TypeError("not checked");
      }
      yield value;
    }
    expect((await leakOf(checksLate, false, () => true)).leak?.fault).toBe(
      "not one of the library's errors"
    );
  });

  it("sees a rejection that arrives after the first look", async () => {
    // A file system's refusal comes back through the I/O thread pool, later than an argument check.
    const rejectsLate = (): Promise<never> =>
      new Promise((_, reject) => setTimeout(() => reject(new TypeError("late")), 50));
    const results = await sweep({ Late: makeNamespace({ rejectsLate }) }, () => true);
    expect(results.find(r => r.name === "Late.rejectsLate")?.leak?.fault).toBe(
      "not one of the library's errors"
    );
  });

  it("every exemption still names a member, and that member still needs it", () => {
    // An exemption is stale only when its member was called and every call settled without the
    // leak. A refusal that comes back through the I/O pool (`Fs.readFileBytes("x")`) can outlast
    // the bound on a loaded runner, and that is not evidence the exemption went unused.
    const stale = Object.keys(EXEMPT).filter(
      name => !seen.has(name) || (!leaksByName.has(name) && !undecided.has(name))
    );
    expect(stale).toEqual([]);
  });
});
