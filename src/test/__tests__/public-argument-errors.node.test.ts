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
  "Archive.randomBytes": DEPRECATED,
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

/** Settle what a call returned, as far as a wrong argument would show: a rejection, or a throw on first pull. */
async function settle(result: unknown): Promise<void> {
  if (result instanceof Promise) {
    // A promise still pending after a moment is not a leak: an argument check rejects at once.
    await Promise.race([result, new Promise(resolve => setTimeout(resolve, 10))]);
  } else if (
    result &&
    typeof (result as { [Symbol.asyncIterator]?: unknown })[Symbol.asyncIterator] === "function"
  ) {
    const it = (result as AsyncIterable<unknown>)[Symbol.asyncIterator]();
    await Promise.race([it.next(), new Promise(resolve => setTimeout(resolve, 10))]);
  }
}

/** The first wrong argument list a member reports badly, with what was wrong. */
async function leakOf(
  fn: AnyFunction,
  construct: boolean,
  owns: ErrorGuard
): Promise<{ args: unknown[]; error: unknown; fault: string } | undefined> {
  for (const args of wrongArguments()) {
    try {
      const Ctor = fn as unknown as new (...a: unknown[]) => unknown;
      await settle(construct ? new Ctor(...args) : fn(...args));
    } catch (error) {
      const fault = faultOf(error, owns);
      if (fault !== undefined) {
        return { args, error, fault };
      }
    }
  }
  return undefined;
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

describe("a wrong argument to a public function is the module's own error", () => {
  it.each(Object.entries(ENTRIES))(
    "%s",
    async (_name, { entry, owns }) => {
      const leaks: string[] = [];
      for (const [name, fn, construct] of membersOf(entry)) {
        seen.add(name);
        const leak = await leakOf(fn, construct, owns);
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

  it("every exemption still names a member, and that member still needs it", () => {
    const stale = Object.keys(EXEMPT).filter(name => !seen.has(name) || !leaksByName.has(name));
    expect(stale).toEqual([]);
  });
});
