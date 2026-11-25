/**
 * Opens every Word document and workbook this repository produces in the real Microsoft Word and Excel, and asks
 * whether they open without a dialog.
 *
 * **Why, beside `verify:libreoffice`.** LibreOffice is the cheap independent reader; Office is the one that counts, and
 * the two disagree in the direction that matters. Word scatter and surface charts and an unqualified sparkline range all
 * opened in LibreOffice and in this library's reader, passed 44,000 tests — and made Office refuse the file. Only
 * opening them in Office found that.
 *
 * **How a file is judged.** The application is told to open the file over AppleScript, and the check waits for that
 * command to complete. A modal dialog raised while the file opens — "found unreadable content … recover?", "Word
 * experienced an error" — holds the command until someone answers it, so:
 *
 * - the command completing, the file among the open documents, and the application still answering a moment later,
 *   is a pass;
 * - the command failing, or completing with the file not open, is a failure, and the run goes on;
 * - the command not completing in time (a minute, more for a large file), or the application not answering
 *   afterwards, is a failure with a dialog up — and only killing the application clears that.
 *
 * Encrypted files, and files whose package has Office fetch content from outside it as they open (an external image,
 * object or template), open to a prompt by design. That prompt cannot be told from a repair prompt here, so they are
 * listed and not opened.
 *
 * **It does not touch a document of yours.** It refuses to start while Word or Excel is running with a document open,
 * opens each file as a copy it made in a directory it created, recognises that copy by its full path — never a name,
 * which another document may share — and closes that one document, unsaved. A document at any other path is never
 * closed: the file itself may be open because someone opened it, during the run as well. When a dialog is up it stops and says which file raised it: clearing it means killing
 * the application, which takes every document it holds, and a document you opened during the run cannot be seen while
 * the application is stuck. `--kill-on-dialog` gives that leave, for a run on a machine nobody is using Office on:
 * the application is killed, the file tried once more, and a copy Office restores afterwards is closed only if its
 * path is this run's copy.
 *
 * Only the Automation permission to control Word and Excel is needed (macOS asks once); window inspection, which would
 * need Accessibility, is not used.
 *
 * **What a pass is worth differs by application, and was measured.** Excel refuses a file whose elements are out of
 * schema order — the `<sheetPr>` order of issue #269 made it show its repair prompt — so an Excel pass is strong
 * evidence. Word repairs order errors and unknown elements in silence and refuses only damage it cannot (malformed XML,
 * a chart it cannot draw), so a Word pass is weak evidence and a Word failure a strong one. The schema-order checks in
 * the test suite cover what Word lets through.
 *
 * ## Usage
 *
 * ```
 * pnpm verify:examples
 * pnpm verify:office --dir tmp/word-examples              # skips loudly when Office is absent
 * pnpm verify:office --dir tmp/word-examples --require    # absence is a failure
 * pnpm verify:office --dir tmp --filter chart             # only paths containing "chart"
 * pnpm verify:office --dir tmp/word-examples --kill-on-dialog   # a long run, unattended, on an idle machine
 * ```
 *
 * A long run needs `--kill-on-dialog`: Office stops answering once after about 62 opens in a row, whatever the file.
 *
 * The directory is named on every run, not defaulted: `tmp/` also holds deliberately broken files, references and
 * leftovers of earlier runs, and a check over whatever happens to be there reports on that rather than on the writer.
 *
 * It is not part of CI: it needs macOS with Office installed, and takes several seconds a file.
 */
import { execFileSync, spawn } from "node:child_process";
import { realpathSync, rmSync } from "node:fs";
import {
  access,
  copyFile,
  mkdir,
  mkdtemp,
  readFile,
  readdir,
  realpath,
  stat
} from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, join } from "node:path";
import process from "node:process";
import { setTimeout as sleep } from "node:timers/promises";

interface App {
  readonly name: string;
  readonly bundle: string;
  /** The AppleScript class of its documents. */
  readonly document: "document" | "workbook";
}

const WORD: App = {
  name: "Microsoft Word",
  bundle: "/Applications/Microsoft Word.app",
  document: "document"
};
const EXCEL: App = {
  name: "Microsoft Excel",
  bundle: "/Applications/Microsoft Excel.app",
  document: "workbook"
};

/** The compound-file signature an encrypted OOXML package is wrapped in. */
const COMPOUND_FILE_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

function appFor(file: string): App {
  return /\.xls[xmb]$/.test(file) ? EXCEL : WORD;
}

function osa(script: string, timeout: number): string | undefined {
  try {
    return execFileSync("osascript", ["-e", script], { timeout, encoding: "utf8" }).trim();
  } catch {
    return undefined;
  }
}

function isRunning(app: App): boolean {
  try {
    execFileSync("pgrep", ["-x", app.name]);
    return true;
  } catch {
    return false;
  }
}

/** Raised when going on could touch a document that is not the file under test, or cannot go on at all. */
class Refusal extends Error {}

/**
 * The full paths of the documents the application has open, or `undefined` when it does not answer — which, once it
 * is running, means a modal dialog is up. A document is recognised by its path, never by its name: a document of
 * someone's may share a name with a file under test, and a copy Office recovered lives somewhere else. Word gives an
 * HFS path and Excel a POSIX one, so both are made POSIX; Word's nameless entry after a restart (`missing value`) is
 * not a document.
 */
function documentPaths(app: App, timeout = 6000): string[] | undefined {
  return documents(app, timeout)?.map(document => document.path);
}

/**
 * The documents the application has open, each with its index — what `${app.document} N` addresses — and its path
 * made canonical, or `undefined` when the application does not answer. One list for finding a document and for
 * closing it, so the two compare paths the same way.
 */
function documents(app: App, timeout = 6000): { index: number; path: string }[] | undefined {
  const out = osa(
    `tell application "${app.name}"
       set listing to ""
       -- By index: neither application lets a script iterate over "every document" directly.
       repeat with i from 1 to (count ${app.document}s)
         -- One entry that cannot be asked (Word's nameless one after a restart) must not end the list.
         try
           set p to (full name of ${app.document} i) as text
           if p is not "missing value" then
             if p does not start with "/" then set p to POSIX path of p
             set listing to listing & i & (character id 9) & p & (character id 10)
           end if
         end try
       end repeat
       return listing
     end tell`,
    timeout
  );
  if (out === undefined) {
    return undefined;
  }
  return out
    .split("\n")
    .filter(line => line.includes("\t"))
    .map(line => {
      const tab = line.indexOf("\t");
      return { index: Number(line.slice(0, tab)), path: canonical(line.slice(tab + 1)) };
    });
}

/**
 * A path with its links resolved, as the copies under test are named: Word reports a temporary file as
 * `/private/var/…` and Excel the same file as `/var/…`, and the two must compare equal.
 */
function canonical(path: string): string {
  try {
    return realpathSync(path);
  } catch {
    return path;
  }
}

/**
 * Refuses to start while the application is running with a document open: that may be someone's unsaved work. A
 * document Office shows when this check starts it — Excel's blank `Book1` — is not the file under test, which is
 * recognised by its path, and is left alone.
 */
function checkIdle(app: App): void {
  if (!isRunning(app)) {
    return;
  }
  const paths = documentPaths(app, 10000);
  if (paths === undefined || paths.length > 0) {
    throw new Refusal(
      `${app.name} is running${paths === undefined ? " and does not answer" : ` with ${paths.length} document(s) open`}. Quit it and run again: this check closes nothing it did not open.`
    );
  }
}

/** Waits for the application to answer, launching it if need be, and returns the documents it has open. */
async function launched(app: App): Promise<string[]> {
  for (let i = 0; i < 60; i++) {
    const paths = documentPaths(app, 5000);
    if (paths !== undefined) {
      return paths;
    }
    await sleep(1000);
  }
  throw new Refusal(`${app.name} did not answer within a minute.`);
}

/**
 * Kills the application and waits until its process is gone. Only with `--kill-on-dialog`: a dialog keeps the
 * application from answering, so it cannot be asked to close anything, and killing it takes every document it holds —
 * nothing of yours only if you opened nothing in it during the run, which this check cannot see while it is stuck.
 * Going on while it is still exiting sent the next file's open to the dying process, so the exit is waited for.
 */
async function kill(app: App): Promise<void> {
  try {
    execFileSync("pkill", ["-9", "-x", app.name]);
  } catch {
    // Not running.
  }
  for (let i = 0; i < 30 && isRunning(app); i++) {
    await sleep(500);
  }
  if (isRunning(app)) {
    throw new Refusal(`${app.name} could not be stopped.`);
  }
}

interface Outcome {
  readonly file: string;
  readonly ok: boolean;
  readonly detail: string;
  /** The application did not answer afterwards: a dialog is up, and nothing can be opened until it is gone. */
  readonly stuck: boolean;
}

/**
 * Runs `script` and resolves with its exit code and output once it ends, or `"timed out"` — with the process killed —
 * when it has not ended within `timeout`.
 */
function run(
  script: string,
  timeout: number
): Promise<{ code: number | "timed out"; output: string }> {
  return new Promise(done => {
    const child = spawn("osascript", ["-e", script], { stdio: ["ignore", "pipe", "pipe"] });
    let output = "";
    child.stdout.on("data", (chunk: Buffer) => (output += chunk.toString()));
    child.stderr.on("data", (chunk: Buffer) => (output += chunk.toString()));
    const timer = setTimeout(() => {
      child.kill("SIGKILL");
      done({ code: "timed out", output });
    }, timeout);
    child.on("close", status => {
      clearTimeout(timer);
      done({ code: status ?? 1, output });
    });
  });
}

/** Whether the application goes on answering for a few seconds — no dialog has come up in the meantime. */
async function keepsAnswering(app: App): Promise<boolean> {
  for (let waited = 0; waited < 8000; waited += 2000) {
    await sleep(2000);
    if (documentPaths(app, 5000) === undefined) {
      return false;
    }
  }
  return true;
}

/** The documents open, asked until the application answers, for half a minute. */
function pathsWhenAnswering(app: App): string[] | undefined {
  let paths: string[] | undefined;
  for (let waited = 0; paths === undefined && waited < 30_000; waited += 10_000) {
    paths = documentPaths(app, 10_000);
  }
  return paths;
}

/**
 * Closes the document at `path` — the file under test, by the path it was opened from — unsaved, and says whether it
 * is gone. Judged by the document list rather than by the command returning in time: on a busy machine a close can
 * outlast its timeout and still complete. The document is addressed through the list's own entry for that path, never
 * by a name another document might share.
 */
function close(app: App, path: string): boolean {
  const target = documents(app, 10_000)?.find(document => document.path === path);
  if (target !== undefined) {
    osa(
      `tell application "${app.name}" to close ${app.document} ${target.index} saving no`,
      15_000
    );
  }
  const paths = pathsWhenAnswering(app);
  return paths !== undefined && !paths.includes(path);
}

/**
 * Opens `file` and judges it by whether the `open` command completes. A dialog raised while the file opens — a repair
 * prompt, an error — holds the command until someone answers it, so a command that has not returned in time means one
 * is up. Sleeping a fixed time and then asking was the first design, and it misjudged both ways: a workbook still
 * loading kept Excel from answering and read as a dialog, and a prompt one file left waiting blocked the next file's
 * open and was charged to that one.
 */
async function open(file: string, copy: string, path: string): Promise<Outcome> {
  const app = appFor(file);
  const before = await launched(app);
  if (before.includes(path)) {
    // A copy this run made, in a directory this run created, so nobody else's: Office restored it after a kill
    // (`--kill-on-dialog`), and it is closed before the file is opened again.
    if (!close(app, path)) {
      throw new Refusal(
        `${app.name} shows ${path} from an earlier attempt and it could not be closed.`
      );
    }
  }
  // Room for a large file to load: a minute, and half a minute more for each megabyte.
  const timeout = 60_000 + Math.ceil((await stat(copy)).size / 1024 / 1024) * 30_000;
  const opened = await run(
    `tell application "${app.name}" to open (POSIX file ${JSON.stringify(path)})`,
    timeout
  );
  if (opened.code === "timed out") {
    return {
      file,
      ok: false,
      stuck: true,
      detail: `the open did not complete in ${timeout / 1000} s: a dialog is waiting (a repair prompt or an error)`
    };
  }
  // Asked after a pause: a prompt raised once the document is showing keeps the application from answering. Asked
  // until it answers, for half a minute: a dialog keeps it silent throughout, a busy machine only for a while.
  await sleep(1500);
  const paths = pathsWhenAnswering(app);
  if (paths === undefined) {
    return { file, ok: false, stuck: true, detail: "no answer once open: a dialog is blocking" };
  }
  if (opened.code !== 0 || !paths.includes(path)) {
    // Word reports a file it cannot read ("Word experienced an error trying to open the file") with a dialog it raises
    // *after* the open command has returned, so an answer now does not mean there is none: watched for a while, the
    // dialog is this file's — taken for none, it blocked the next file and was charged to that one.
    const stuck = !(await keepsAnswering(app));
    const detail =
      opened.code !== 0
        ? `the open failed: ${opened.output.trim()}`
        : "the open completed, but the file is not open";
    return { file, ok: false, stuck, detail: stuck ? `${detail}; then a dialog came up` : detail };
  }
  if (!close(app, path)) {
    return { file, ok: false, stuck: true, detail: "could not be closed: a dialog is blocking" };
  }
  return { file, ok: true, stuck: false, detail: "opened without a dialog" };
}

/**
 * Relationship types whose external target Office fetches when the file opens — and asks before it does (file access
 * outside its sandbox, content from the network). Such a file opens to a prompt by design, which this check cannot
 * tell from a repair prompt, so it is listed rather than opened. An external hyperlink is fetched only when clicked.
 */
const FETCHED_ON_OPEN = /\/(image|oleObject|externalLinkPath|attachedTemplate|subDocument|frame)"/;

/** Whether `file` has a relationship Office fetches from outside the package when it opens. */
function fetchesOnOpen(file: string): boolean {
  let rels: string;
  try {
    rels = execFileSync("unzip", ["-p", file, "*.rels"], {
      encoding: "utf8",
      maxBuffer: 64 * 1024 * 1024
    });
  } catch {
    return false;
  }
  return (rels.match(/<Relationship\b[^>]*>/g) ?? []).some(
    rel => /TargetMode="External"/.test(rel) && FETCHED_ON_OPEN.test(rel)
  );
}

async function collect(dir: string, out: string[]): Promise<void> {
  let entries: Awaited<ReturnType<typeof readdir>>;
  try {
    entries = await readdir(dir, { withFileTypes: true });
  } catch {
    return;
  }
  for (const entry of entries) {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) {
      await collect(path, out);
    } else if (
      /\.(docx|docm|xlsx|xlsm|xlsb)$/.test(entry.name) &&
      // An Office lock file beside a document someone has open.
      !entry.name.startsWith("~$")
    ) {
      out.push(path);
    }
  }
}

async function installed(app: App): Promise<boolean> {
  return access(app.bundle).then(
    () => true,
    () => false
  );
}

async function main(): Promise<void> {
  const argv = process.argv.slice(2);
  const required = argv.includes("--require");
  const killOnDialog = argv.includes("--kill-on-dialog");
  const option = (name: string): string | undefined => {
    const at = argv.indexOf(name);
    return at === -1 ? undefined : argv[at + 1];
  };
  const filter = option("--filter");
  const dir = option("--dir");
  if (dir === undefined) {
    console.error(
      "✗ verify:office — name the directory to check: --dir <path> (see the header of scripts/verify-office.ts)."
    );
    process.exit(1);
  }

  const files: string[] = [];
  await collect(dir, files);
  const targets = (filter === undefined ? files : files.filter(f => f.includes(filter))).sort();
  if (targets.length === 0) {
    console.error(
      `✗ verify:office — no documents to check under ${dir}${filter === undefined ? ". Run `pnpm verify:examples` first." : ` matching ${JSON.stringify(filter)}.`}`
    );
    process.exit(1);
  }

  const missing: App[] = [];
  for (const app of new Set(targets.map(appFor))) {
    if (process.platform !== "darwin" || !(await installed(app))) {
      missing.push(app);
    }
  }
  if (missing.length > 0) {
    const message = `${missing.map(app => app.name).join(" and ")} not found (this check needs macOS, with Office in /Applications).`;
    if (required) {
      console.error(`✗ verify:office — ${message}`);
      process.exit(1);
    }
    // Loud, and it says what was not checked: a quiet skip is how a gate dies green.
    console.log(`⚠ verify:office — skipped, nothing was verified.\n  ${message}`);
    return;
  }

  try {
    for (const app of new Set(targets.map(appFor))) {
      checkIdle(app);
    }
  } catch (error) {
    if (error instanceof Refusal) {
      console.error(`✗ verify:office — ${error.message}`);
      process.exit(1);
    }
    throw error;
  }

  // An encrypted package opens to a password prompt, which says nothing about the writer: listed, not opened.
  const encrypted: string[] = [];
  for (const file of targets) {
    const head = (await readFile(file)).subarray(0, COMPOUND_FILE_SIGNATURE.length);
    if (COMPOUND_FILE_SIGNATURE.every((byte, i) => head[i] === byte)) {
      encrypted.push(file);
    }
  }
  const external = targets.filter(file => !encrypted.includes(file) && fetchesOnOpen(file));
  const opened = targets.filter(file => !encrypted.includes(file) && !external.includes(file));
  for (const [what, files] of [
    ["encrypted (opens to a password prompt)", encrypted],
    [
      "with content Office fetches from outside when it opens (opens to a prompt by design)",
      external
    ]
  ] as const) {
    if (files.length > 0) {
      console.log(
        `${files.length} file(s) ${what}, listed and not opened:\n${files.map(file => `  - ${file}`).join("\n")}\n`
      );
    }
  }
  if (opened.length === 0) {
    console.error(
      `✗ verify:office — none of the ${targets.length} file(s) could be opened (all listed above); nothing was verified.`
    );
    process.exit(1);
  }
  console.log(`${opened.length} file(s) to open\n`);

  const failures: Outcome[] = [];
  const retried: string[] = [];
  let done = 0;
  // **Each file is opened as a copy this run made, in a directory it created.** A document is recognised by its path,
  // and only such a path proves the document is this check's: the file itself may be open in Office because someone
  // opened it — during the run, too — and closing it unsaved would lose their edits. The name is kept, so a dialog
  // still names the file it is about. Removed when the process ends, however it ends.
  const work = await mkdtemp(join(tmpdir(), "documonster-office-"));
  process.on("exit", () => rmSync(work, { recursive: true, force: true }));
  for (const [index, file] of opened.entries()) {
    const app = appFor(file);
    const copyDir = join(work, String(index));
    await mkdir(copyDir);
    const copy = join(copyDir, basename(file));
    await copyFile(file, copy);
    const path = await realpath(copy);
    let outcome: Outcome;
    let firstAttempt: Outcome | undefined;
    try {
      outcome = await open(file, copy, path);
      if (!outcome.ok && outcome.stuck) {
        if (!killOnDialog) {
          // A dialog is up and only killing the application clears it, which takes every document it holds. Without
          // the operator's leave, stop here and say where — they answer it, and run again from this file.
          failures.push(outcome);
          console.log(
            `  ✗ ${String(done + 1).padStart(3)}/${opened.length} ${file}\n      ${outcome.detail}`
          );
          throw new Refusal(
            `${app.name} is showing a dialog for ${file}. Answer it there, then run again — or pass --kill-on-dialog to have this check kill ${app.name} instead, when nothing of yours is open in it.`
          );
        }
        await kill(app);
      }
      // Tried once more: a dialog the file raises comes up every time, while Office stops answering once after about
      // 62 opens in a row, whatever the file — 70 copies of one plain workbook stalled at the 63rd, and every file a
      // long run failed on opened cleanly on its own. Once, not until it passes, and a pass on the retry is reported.
      if (!outcome.ok) {
        firstAttempt = outcome;
        outcome = await open(file, copy, path);
        if (!outcome.ok && outcome.stuck) {
          if (!killOnDialog) {
            throw new Refusal(
              `${app.name} is showing a dialog for ${file} again. Answer it there.`
            );
          }
          await kill(app);
        }
      }
    } catch (error) {
      if (error instanceof Refusal) {
        console.error(`\n✗ verify:office — ${error.message}`);
        process.exit(1);
      }
      throw error;
    }
    done += 1;
    if (outcome.ok) {
      console.log(
        `  ✓ ${String(done).padStart(3)}/${opened.length} ${file}${firstAttempt ? `\n      passed on a retry; the first attempt: ${firstAttempt.detail}` : ""}`
      );
      if (firstAttempt) {
        retried.push(file);
      }
    } else {
      failures.push(outcome);
      console.log(
        `  ✗ ${String(done).padStart(3)}/${opened.length} ${file}\n      ${outcome.detail}`
      );
    }
  }
  if (failures.length > 0) {
    console.error(
      `\n✗ verify:office — ${failures.length} of ${opened.length} file(s) did not open cleanly in Office.`
    );
    process.exit(1);
  }
  console.log(
    `\n✓ verify:office — all ${opened.length} file(s) opened in Office without a dialog${retried.length > 0 ? `, ${retried.length} of them on a retry (listed above)` : ""}.`
  );
  console.log(
    "  An Excel pass is strong evidence; Word repairs order errors in silence, so a Word pass is weak."
  );
}

await main();
