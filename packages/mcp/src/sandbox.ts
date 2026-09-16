/**
 * Path sandbox.
 *
 * The single most important file in this package. An MCP server hands a model
 * the ability to name files, so every path a tool touches must go through
 * {@link resolveInRoot} — there is no other legitimate way to turn a
 * model-supplied string into a filesystem path.
 *
 * Two attacks this defends against:
 *
 * 1. Traversal — `../../.ssh/id_rsa`, or a plain absolute path. Handled by
 *    resolving against the root and then checking containment.
 * 2. Symlink escape — a link inside the root pointing outside it. Handled by
 *    realpath'ing the deepest part of the path that actually exists, which is
 *    why the check cannot be a naive `startsWith` on the raw string.
 *
 * Do NOT rely on the MCP client's tool-approval prompt for this. Programmatic
 * clients have no prompt.
 */

import fs from "node:fs/promises";
import path from "node:path";

import type { ServerConfig } from "./config.js";
import { toolError } from "./errors.js";

/** Symlink hops allowed before a path is treated as a loop. */
const MAX_SYMLINK_DEPTH = 32;

/** Matches a URL scheme (`http:`, `file:`) without matching a Windows drive letter (`C:`). */
const URL_SCHEME_RE = /^[a-zA-Z][a-zA-Z0-9+.-]+:/;

/**
 * Win32 device names, reserved in *every* directory rather than living in one
 * place on disk.
 *
 * The set is Microsoft's, not a guess
 * ({@link https://learn.microsoft.com/en-us/windows/win32/fileio/naming-a-file | Naming a File}),
 * and two entries in it surprise people:
 *
 * - The ISO 8859-1 superscripts `¹²³` are treated as digits by the Win32 path
 *   parser, so `COM¹` is the same device as `COM1`. They have no case, so
 *   upper-casing before the test leaves them alone.
 * - `CONIN$`/`CONOUT$` are the console's input and output handles. They are
 *   absent from the "do not use" list but `CreateFile` opens them, which is all
 *   that matters here.
 *
 * `COM0`/`LPT0` are deliberately *not* included: no such device exists, so a
 * file of that name is an ordinary file and refusing it would be a false
 * positive on every platform.
 */
const WINDOWS_RESERVED_DEVICE_RE = /^(?:CON|PRN|AUX|NUL|CONIN\$|CONOUT\$|COM[1-9¹²³]|LPT[1-9¹²³])$/;

export interface ResolveOptions {
  /**
   * When true, the path must already exist. Use for inputs; leave false for
   * write targets, whose parent is checked instead.
   */
  readonly mustExist?: boolean;
}

/**
 * Turn a model-supplied path into an absolute path guaranteed to sit inside
 * `config.root`.
 *
 * @throws {McpToolError} `invalid_input` for a URL, an empty string, a NUL byte
 *   or a reserved Windows device name, `not_found` when `mustExist` is set and
 *   it is not there, and `outside_root` for anything that escapes the sandbox.
 */
export async function resolveInRoot(
  config: ServerConfig,
  userPath: string,
  options: ResolveOptions = {}
): Promise<string> {
  assertUsablePath(userPath);

  const outputRelative = outputPathPart(userPath);
  const base = outputRelative === undefined ? config.root : config.outputRoot;
  const relative = outputRelative ?? userPath;
  const candidate = path.resolve(base, relative);
  const resolved = await resolveSegments(candidate);

  if (!isInside(base, resolved)) {
    throw toolError.outsideRoot(
      `path resolves outside the server root: ${userPath}`,
      outputRelative === undefined
        ? "Only paths inside the configured --root are readable. Ask the user to restart the server with a wider root if this file is genuinely needed."
        : "Only paths inside --output-root are accessible through @output/."
    );
  }

  if (options.mustExist === true && !(await pathExists(resolved))) {
    throw toolError.notFound(
      `no such file or directory: ${userPath}`,
      "Call doc_inspect on the parent directory to see what is actually there."
    );
  }

  return resolved;
}

/**
 * Resolve a path a tool is about to write to, plus its sidecar files.
 *
 * A write target needs more than containment: every path the tool derives from
 * it — a `.bak`, a temporary file — must be confined too, or a pre-planted
 * symlink at `report.xlsx.bak` lets an in-place edit write outside the root.
 */
export async function resolveWriteTarget(
  config: ServerConfig,
  userPath: string,
  options: { readonly inPlace?: boolean } = {}
): Promise<WriteTarget> {
  if (options.inPlace === true) {
    const outputRelative = outputPathPart(userPath);
    if (outputRelative !== undefined) {
      const target = await resolveOutputPath(config, outputRelative);
      const backup = await resolveOutputPath(config, `${outputRelative}.bak`);
      return {
        path: target,
        backupPath: backup,
        display: `@output/${normaliseDisplay(outputRelative)}`
      };
    }
    if (!config.allowInPlace) {
      throw toolError.readonly(
        "in-place edits are disabled",
        "Provide an `out` path to write safely under @output/, or restart the server with --allow-in-place."
      );
    }
    const target = await resolveInRoot(config, userPath);
    const backup = await resolveInputSidecar(config, `${userPath}.bak`);
    return { path: target, backupPath: backup, display: userPath };
  }

  const relative = outputPathPart(userPath) ?? userPath;
  const target = await resolveOutputPath(config, relative);
  const backup = await resolveOutputPath(config, `${relative}.bak`);
  return { path: target, backupPath: backup, display: `@output/${normaliseDisplay(relative)}` };
}

/** Resolve a write destination inside the disjoint output root. */
export async function resolveOutputPath(config: ServerConfig, userPath: string): Promise<string> {
  assertUsablePath(userPath);
  const relative = outputPathPart(userPath) ?? userPath;
  if (path.isAbsolute(relative)) {
    throw toolError.invalidInput(
      `output path must be relative: ${userPath}`,
      "Write paths are relative to --output-root and are returned as @output/<path>."
    );
  }
  const candidate = path.resolve(config.outputRoot, relative);
  const resolved = await resolveSegments(candidate);
  if (!isInside(config.outputRoot, resolved)) {
    throw toolError.outsideRoot(
      `output path resolves outside --output-root: ${userPath}`,
      "Use a relative output path without `..`."
    );
  }
  return resolved;
}

/** Resolve a sidecar beside an input file; available only in explicit in-place mode. */
async function resolveInputSidecar(config: ServerConfig, userPath: string): Promise<string> {
  const candidate = path.resolve(config.root, userPath);
  const resolved = await resolveSegments(candidate);
  if (!isInside(config.root, resolved)) {
    throw toolError.outsideRoot(`sidecar path resolves outside --root: ${userPath}`);
  }
  return resolved;
}

/**
 * Reject a path string that cannot safely name a file, before it is resolved.
 *
 * Separate from containment because these are properties of the *string*, and
 * two of them are invisible to a containment check:
 *
 * - A NUL byte. Node throws on it deep inside `fs`, which surfaced to the model
 *   as an unclassified `internal` error carrying a raw library message. It is an
 *   invalid argument, and saying so is what lets the model correct itself.
 * - A Win32 device name. `NUL`, `CON`, `COM1` are reserved in every directory,
 *   so `<root>/CON` passes containment yet opens the *console* rather than a
 *   file — and this server's console is the stdio transport it speaks JSON-RPC
 *   over. Reading it would wedge the process. Refused on Windows only, because
 *   on POSIX a file called `NUL` is an ordinary file and rejecting it would be
 *   a false positive.
 */
function assertUsablePath(userPath: string): void {
  if (typeof userPath !== "string" || userPath.trim().length === 0) {
    throw toolError.invalidInput("path must be a non-empty string");
  }

  if (userPath.includes("\u0000")) {
    throw toolError.invalidInput(
      "path must not contain a NUL byte",
      "Pass the plain file name. A NUL byte cannot appear in a filesystem path."
    );
  }

  if (URL_SCHEME_RE.test(userPath)) {
    throw toolError.invalidInput(
      `path must be a local path inside the sandbox, not a URL: ${userPath}`,
      "Use a path relative to the server root. This server reads local files only."
    );
  }

  if (process.platform === "win32") {
    const reserved = reservedDeviceSegment(userPath);
    if (reserved !== undefined) {
      throw toolError.invalidInput(
        `path names a reserved Windows device: ${reserved}`,
        "Rename the file. CON, PRN, AUX, NUL, COM<n> and LPT<n> address hardware on Windows, not files on disk."
      );
    }
  }
}

/**
 * The first path segment that names a Win32 device, or undefined.
 *
 * Exported for tests: the guard above only fires on Windows, so this is the
 * only way to assert the matching rules on any host.
 *
 * A segment is cut at its first `.` or `:` before matching, because Windows
 * resolves both `NUL.txt` and `NUL:stream` to the device — the extension and
 * the stream suffix are not part of the name it looks up. Surrounding blanks go
 * too, since the parser ignores them. `NULL` and `CONSOLE` survive all of that
 * and stay ordinary files.
 */
export function reservedDeviceSegment(userPath: string): string | undefined {
  for (const segment of userPath.split(/[/\\]/)) {
    const stem = (segment.split(/[.:]/)[0] ?? "").trim();
    if (WINDOWS_RESERVED_DEVICE_RE.test(stem.toUpperCase())) {
      return segment;
    }
  }
  return undefined;
}

/** Return the path after the virtual @output/ prefix, or undefined for input. */
function outputPathPart(userPath: string): string | undefined {
  const normalised = userPath.replace(/\\/g, "/");
  if (normalised === "@output") {
    return ".";
  }
  return normalised.startsWith("@output/") ? normalised.slice("@output/".length) : undefined;
}

function normaliseDisplay(relative: string): string {
  return relative.replace(/\\/g, "/").replace(/^\.\//, "");
}

/** Public display spelling for a path under the output root. */
export function outputDisplay(userPath: string): string {
  const relative = outputPathPart(userPath) ?? userPath;
  return `@output/${normaliseDisplay(relative)}`;
}

export interface WriteTarget {
  /** Confined absolute path to write. */
  readonly path: string;
  /** Confined absolute path for the `.bak` sidecar. */
  readonly backupPath: string;
  /** The caller's spelling, for messages. */
  readonly display: string;
}

/**
 * Choose the destination for a tool that edits an existing file.
 *
 * Files already under `@output/` may be edited in place because that root is
 * the server's private writable workspace. Input files require an explicit
 * `out` unless the operator enabled the compatibility escape hatch.
 */
export async function resolveEditTarget(
  config: ServerConfig,
  sourcePath: string,
  out: string | undefined
): Promise<WriteTarget> {
  if (out !== undefined) {
    return await resolveWriteTarget(config, out);
  }
  if (outputPathPart(sourcePath) !== undefined) {
    return await resolveWriteTarget(config, sourcePath, { inPlace: true });
  }
  return await resolveWriteTarget(config, sourcePath, { inPlace: true });
}

/** True when a path exists, following symlinks. */
async function pathExists(target: string): Promise<boolean> {
  try {
    await fs.stat(target);
    return true;
  } catch {
    return false;
  }
}

/**
 * Assert the server is allowed to mutate the filesystem.
 *
 * Mutating tools are already withheld from `tools/list` under `--readonly`;
 * this is the second line of defence for a client that calls a tool it was
 * never told about.
 *
 * @remarks
 * Called by every mutating tool (`sheet_write`, `archive_read` when extracting)
 * before it touches the disk. `selectTools` filtering is the other half.
 */
export function assertWritable(config: ServerConfig): void {
  if (config.readonly) {
    throw toolError.readonly(
      "this server is running in --readonly mode",
      "Report to the user that writing is disabled; do not retry."
    );
  }
}

/**
 * True when `child` is `parent` itself or lives beneath it.
 *
 * Compares path segments rather than raw strings, so `/root-2` is correctly
 * rejected as a sibling of `/root` instead of matching a `startsWith("/root")`.
 */
export function isInside(parent: string, child: string): boolean {
  if (child === parent) {
    return true;
  }
  const relative = path.relative(parent, child);
  return (
    relative.length > 0 &&
    !relative.startsWith(`..${path.sep}`) &&
    relative !== ".." &&
    !path.isAbsolute(relative)
  );
}

/**
 * Realpath every segment of `target`, following symlinks explicitly.
 *
 * `fs.realpath` cannot be used for the whole path: it throws `ENOENT` both for
 * "this does not exist" and for "this is a symlink whose target does not
 * exist". Treating the second case as the first is a sandbox escape — a
 * dangling link pointing outside the root would be reported as an ordinary
 * not-yet-existing path inside it, and the tool would then create the file at
 * the link's real destination.
 *
 * So each segment is resolved in turn: an existing symlink is read and its
 * target resolved recursively, whether or not that target exists yet, while a
 * genuinely missing segment is simply appended.
 */
async function resolveSegments(target: string, depth = 0): Promise<string> {
  if (depth > MAX_SYMLINK_DEPTH) {
    throw toolError.invalidInput(
      `too many symbolic links while resolving ${target}`,
      "The path contains a symlink loop."
    );
  }

  const parent = path.dirname(target);
  // `path.dirname` is a fixed point at the filesystem root, which ends recursion.
  const resolvedParent = parent === target ? target : await resolveSegments(parent, depth + 1);
  const candidate = path.join(resolvedParent, path.basename(target));

  let stats;
  try {
    stats = await fs.lstat(candidate);
  } catch (cause) {
    if ((cause as { code?: string }).code === "ENOENT") {
      // Genuinely absent: a valid write target, and nothing left to follow.
      return candidate;
    }
    throw cause;
  }

  if (!stats.isSymbolicLink()) {
    return candidate;
  }

  const link = await fs.readlink(candidate);
  const absolute = path.isAbsolute(link) ? link : path.join(resolvedParent, link);
  return await resolveSegments(absolute, depth + 1);
}
