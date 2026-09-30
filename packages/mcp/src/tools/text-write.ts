/**
 * `text_write` — upload a text file under @output/ in numbered parts.
 *
 * A tool argument has to be generated in full before the client sends the call,
 * so no argument can be longer than one model reply. A long translation, a large
 * JSON payload or thousands of rows of CSV therefore cannot travel in a single
 * call, whatever the server does — the call is aborted on the client before this
 * process ever sees it, and MCP has no streamed arguments to split it with.
 *
 * This tool is the one place that limit is lifted. Content arrives in parts, each
 * small enough for one reply, and the finished file is an ordinary file every
 * tool reads by path (`doc_write.from`, `template_fill.dataFrom`,
 * `sheet_write.sheetsFrom`, `sheet_edit.opsFrom`, `form_fill.valuesFrom`, …).
 *
 * ## State model
 *
 * A session is a directory holding immutable part files and one manifest,
 * `meta.json`. **The manifest is the only commit point.** A part's bytes go to a
 * new, uniquely named file first; the manifest is then replaced atomically to
 * point at it. A crash or failure between the two leaves an unreferenced file and
 * the previous, still-consistent state — never a manifest describing bytes that
 * are not there.
 *
 * Every change to content (a part's text, or `total`) increments `revision`.
 * Publishing records the revision and the SHA-256 of what it wrote. So "is the
 * published file current" is a comparison of two numbers, not a guess from the
 * last call, and any call — including a bare status call — publishes whenever the
 * upload is complete and the published revision is behind.
 *
 * Publishing writes the manifest twice: a `pending` record (revision and hash)
 * just before the file is installed, and `published` just after. If the process
 * dies between install and the second write, the next publish finds the file
 * matching `pending` and knows it wrote it, rather than refusing it as someone
 * else's.
 *
 * ## Guarantees
 *
 * - Nothing is published until parts 1…total are all present.
 * - A late or retried call from an earlier upload carries that upload's id, so it
 *   cannot land in a newer upload's content.
 * - Resending a part with identical text changes nothing.
 * - The size limit is checked for the whole request before anything is written,
 *   and again from the manifest before every publish.
 * - A republish replaces the file only if it still holds exactly what this upload
 *   wrote (by hash); otherwise it is refused and nothing is overwritten.
 * - Calls on one session, and publishes to one destination, are serialised.
 *
 * ## Deliberate limits
 *
 * Serialisation is within this server process. Two server processes sharing one
 * output root and one session could still race on the manifest, and a program
 * that ignores this tool entirely can write the destination in the instant between
 * the hash check and the rename; closing that needs an OS-level lock, and a
 * single MCP client drives a single session. Two different texts sent for the same
 * part concurrently leave whichever arrived last — the tool cannot know which the
 * caller meant.
 */

import { createHash, randomBytes } from "node:crypto";
import { createReadStream, createWriteStream } from "node:fs";
import { mkdir, readdir, readFile, rm, stat, unlink, utimes } from "node:fs/promises";
import path from "node:path";

import { z } from "zod";

import type { ServerConfig } from "../config.js";
import { formatToolError, toolError } from "../errors.js";
import { assertWritable, outputDisplay, resolveOutputPath } from "../sandbox.js";
import { exists, writeFileAtomic, writeWithPolicy } from "./fs-helpers.js";
import { formatBytes, textResult } from "./result.js";
import { defineTool } from "./types.js";

/** Where sessions live, below the output root; hidden so listings show real output. */
const UPLOADS_DIRECTORY = ".documonster-uploads";

/**
 * Highest part number. Bounds the manifest, which lists every part, at a few
 * hundred kilobytes — and at a few thousand words a part it is far beyond any
 * document a model will write.
 */
const MAX_PART = 10_000;

/** A session untouched this long is abandoned, and swept when a new one starts. */
const SESSION_TTL_MS = 24 * 60 * 60 * 1000;

const UPLOAD_ID = /^[0-9a-f]{20}$/;

interface StoredPart {
  /** File name inside the session directory; immutable once written. */
  readonly file: string;
  readonly size: number;
}

interface PublishRecord {
  readonly revision: number;
  readonly sha256: string;
}

/** A session's manifest: the single source of truth for its state. */
interface SessionMeta {
  readonly version: 2;
  /** The destination exactly as the caller named it, relative to --output-root. */
  readonly path: string;
  readonly total: number;
  readonly overwrite: boolean;
  /** Incremented by every change to content or total. */
  readonly revision: number;
  readonly parts: Readonly<Record<string, StoredPart>>;
  /** What is at `path`, as far as this upload knows. */
  readonly published?: PublishRecord;
  /** Written just before a publish installs the file; see the module comment. */
  readonly pending?: PublishRecord;
}

interface Session {
  readonly id: string;
  readonly directory: string;
}

/** The path, relative to --output-root, of a session's directory or one of its files. */
function sessionPath(id: string, file?: string): string {
  return file === undefined ? `${UPLOADS_DIRECTORY}/${id}` : `${UPLOADS_DIRECTORY}/${id}/${file}`;
}

/**
 * Resolve a file inside a session through the sandbox. The names are the server's
 * own, but the directory sits in a writable root: a symlink planted where a part
 * belongs must not make this tool read or write outside it.
 */
function sessionFile(config: ServerConfig, id: string, file: string): Promise<string> {
  return resolveOutputPath(config, sessionPath(id, file));
}

async function readMeta(config: ServerConfig, id: string): Promise<SessionMeta> {
  const file = await sessionFile(config, id, "meta.json");
  const text = await readFile(file, "utf8").catch((cause: unknown) => {
    if ((cause as { code?: string }).code === "ENOENT") {
      throw toolError.notFound(
        `no upload session ${JSON.stringify(id)}`,
        "Sessions expire after 24 hours unused. Start a new one with path and total (and no upload)."
      );
    }
    throw cause;
  });
  let meta: SessionMeta;
  try {
    meta = JSON.parse(text) as SessionMeta;
  } catch (cause) {
    throw toolError.invalidInput(
      `upload ${id} has an unreadable manifest`,
      "Something other than this tool changed the session. Start a new upload.",
      { cause }
    );
  }
  if (meta.version !== 2) {
    throw toolError.invalidInput(
      `upload ${id} was started by an incompatible version of this server`,
      "Start a new upload."
    );
  }
  return meta;
}

/** Commit a new manifest. Everything a call changes becomes visible here, at once. */
async function commit(config: ServerConfig, id: string, meta: SessionMeta): Promise<void> {
  await writeFileAtomic(await sessionFile(config, id, "meta.json"), JSON.stringify(meta));
}

function partNumbers(meta: SessionMeta): number[] {
  return Object.keys(meta.parts)
    .map(Number)
    .sort((a, b) => a - b);
}

function missingParts(meta: SessionMeta): number[] {
  const missing: number[] = [];
  for (let part = 1; part <= meta.total; part += 1) {
    if (meta.parts[part] === undefined) {
      missing.push(part);
    }
  }
  return missing;
}

/** Bytes the published file would have: parts 1…total only. */
function includedSize(meta: SessionMeta): number {
  let size = 0;
  for (const [part, stored] of Object.entries(meta.parts)) {
    if (Number(part) <= meta.total) {
      size += stored.size;
    }
  }
  return size;
}

/** Compress `[1,2,3,5]` to `1–3, 5` so a long session's report stays one line. */
function describeRanges(numbers: readonly number[]): string {
  const ranges: string[] = [];
  let index = 0;
  while (index < numbers.length) {
    const start = numbers[index] as number;
    let end = start;
    while (numbers[index + 1] === end + 1) {
      index += 1;
      end += 1;
    }
    ranges.push(start === end ? String(start) : `${start}–${end}`);
    index += 1;
  }
  return ranges.join(", ");
}

/**
 * Run work one at a time per key.
 *
 * Keyed twice: by session, because a call reads the manifest, decides and commits,
 * and two interleaved calls would each commit a manifest missing the other's part;
 * and by destination, because two sessions publishing one path must not both pass
 * the ownership check before either installs its file.
 */
const queues = new Map<string, Promise<unknown>>();

function serialised<T>(key: string, work: () => Promise<T>): Promise<T> {
  const previous = queues.get(key) ?? Promise.resolve();
  const next = previous.then(work, work);
  const settled = next.catch(() => undefined);
  queues.set(key, settled);
  void settled.then(() => {
    if (queues.get(key) === settled) {
      queues.delete(key);
    }
  });
  return next;
}

/**
 * Remove sessions nobody has touched within the TTL. Best effort: never fails a call.
 *
 * Every call on a session touches its directory, so "untouched" means no call at
 * all — a status check or an unchanged resend counts. A session with a call in
 * flight is skipped whatever its age.
 */
async function sweepExpired(config: ServerConfig): Promise<void> {
  const root = await resolveOutputPath(config, UPLOADS_DIRECTORY);
  const names = await readdir(root).catch(() => [] as string[]);
  const now = Date.now();
  await Promise.all(
    names
      .filter(name => UPLOAD_ID.test(name) && !queues.has(`session:${name}`))
      .map(async name => {
        const directory = path.join(root, name);
        const modified = await stat(directory).then(
          info => info.mtimeMs,
          () => now
        );
        if (now - modified > SESSION_TTL_MS) {
          await rm(directory, { recursive: true, force: true }).catch(() => undefined);
        }
      })
  );
}

function withoutPending(meta: SessionMeta): SessionMeta {
  const { pending, ...rest } = meta;
  void pending;
  return rest;
}

async function sha256Of(file: string): Promise<string> {
  const hash = createHash("sha256");
  for await (const chunk of createReadStream(file)) {
    hash.update(chunk as Buffer);
  }
  return hash.digest("hex");
}

/**
 * Assemble parts 1…total into the destination, streaming, and install it atomically.
 *
 * @returns The manifest after publishing.
 */
async function publish(
  config: ServerConfig,
  session: Session,
  meta: SessionMeta
): Promise<SessionMeta> {
  const target = await resolveOutputPath(config, meta.path);
  return serialised(`target:${target}`, async () => {
    // Re-checked here, not only when a part arrives: total may have grown to take in
    // parts that were stored while it was smaller and so were never counted.
    const size = includedSize(meta);
    if (size > config.maxFileSize) {
      throw toolError.invalidInput(
        `${outputDisplay(meta.path)} would be ${size} bytes, over the ${config.maxFileSize} byte limit`,
        "Lower total, replace large parts, or ask the user to raise --max-file-size."
      );
    }

    // Replace an existing file only if it is one this upload wrote.
    const ours = [meta.published?.sha256, meta.pending?.sha256].filter(
      (value): value is string => value !== undefined
    );
    let replace = meta.overwrite;
    if (ours.length > 0 && (await exists(target))) {
      if (!ours.includes(await sha256Of(target))) {
        throw toolError.invalidInput(
          `${outputDisplay(meta.path)} was changed by something else after this upload published it`,
          "Nothing was overwritten. Start a new upload to a different path, or with overwrite: true to replace it."
        );
      }
      replace = true;
    }

    const files: string[] = [];
    for (let part = 1; part <= meta.total; part += 1) {
      files.push(await sessionFile(config, session.id, (meta.parts[part] as StoredPart).file));
    }

    let pending: PublishRecord | undefined;
    await writeWithPolicy(target, replace, async temporary => {
      const hash = createHash("sha256");
      const output = createWriteStream(temporary);
      try {
        for (const file of files) {
          for await (const chunk of createReadStream(file)) {
            hash.update(chunk as Buffer);
            if (!output.write(chunk)) {
              await new Promise<void>(resolve => output.once("drain", resolve));
            }
          }
        }
      } catch (cause) {
        output.destroy();
        throw cause;
      }
      await new Promise<void>((resolve, reject) => {
        output.end((error?: Error | null) => (error ? reject(error) : resolve()));
      });
      // Recorded before the file is installed, so a crash straight after installing
      // still leaves this upload able to recognise the file as its own.
      pending = { revision: meta.revision, sha256: hash.digest("hex") };
      await commit(config, session.id, { ...meta, pending });
    });

    const published: SessionMeta = { ...withoutPending(meta), published: pending as PublishRecord };
    await commit(config, session.id, published);
    return published;
  });
}

export const textWriteTool = defineTool({
  name: "text_write",
  group: "core",
  title: "Upload a text file in parts",
  description:
    "Upload a text file (.md, .json, .csv, .mmd, .txt …) to @output/ in numbered parts, for content too long for one call — a full translation, a long report, large JSON or CSV. One call carrying all of it can exceed your output limit and be aborted before it is sent, so keep each part to a few thousand words. Start with path, total and part 1 (no upload); the reply gives an upload id — pass it with every later part. The file is published only when parts 1…total are all present. Resend a part to replace it; call with just upload to see what is missing or to retry a failed publish. Then pass the file by path: doc_write `from`, template_fill `dataFrom`, sheet_write `sheetsFrom`, sheet_edit `opsFrom`, form_fill `valuesFrom`, diagram_render `from`.",
  inputSchema: {
    upload: z
      .string()
      .optional()
      .describe("Upload id returned by the first call. Omit only to start a new upload."),
    path: z
      .string()
      .min(1)
      .optional()
      .describe(
        'Starting only: destination below --output-root, e.g. "report.md"; returned as @output/<path>.'
      ),
    total: z
      .number()
      .int()
      .min(1)
      .max(MAX_PART)
      .optional()
      .describe(
        "How many parts the file has. Required to start; may be changed later if the plan changes."
      ),
    part: z
      .number()
      .int()
      .min(1)
      .max(MAX_PART)
      .optional()
      .describe("This part's number, from 1 to total. Parts are joined in number order."),
    text: z
      .string()
      .optional()
      .describe(
        "This part's content, joined to its neighbours exactly as sent — end a Markdown part with a blank line."
      ),
    overwrite: z
      .boolean()
      .optional()
      .describe("Starting only: replace an existing file at path. Defaults to false.")
  },
  annotations: {
    readOnlyHint: false,
    destructiveHint: false,
    // Not idempotent: a call without `upload` starts a new session every time.
    idempotentHint: false,
    openWorldHint: false
  },
  mutates: true,
  handler: async (args, context) => {
    const { config } = context;
    assertWritable(config);

    if ((args.part === undefined) !== (args.text === undefined)) {
      throw toolError.invalidInput(
        "`part` and `text` go together",
        "Send a part as { part, text }, or neither to only start a session or check its status."
      );
    }

    const session =
      args.upload === undefined ? await start(config, args) : await open(config, args);

    return serialised(`session:${session.id}`, async () => {
      // Read inside the queue: an earlier call may have committed since `open`.
      let meta = await readMeta(config, session.id);
      await utimes(session.directory, new Date(), new Date()).catch(() => undefined);

      meta = await apply(config, session, meta, args);

      const missing = missingParts(meta);
      const size = includedSize(meta);
      let state: string;
      if (missing.length > 0) {
        state = `**not published yet** — waiting for part(s) ${describeRanges(missing)}`;
      } else if (meta.published?.revision !== meta.revision) {
        // Decided from the manifest, not from what this call did: a status call
        // after a failed publish retries it. A failure is reported rather than
        // thrown because the part is already committed — an error would read as
        // "this part was lost" and invite a pointless resend.
        try {
          meta = await publish(config, session, meta);
          state = `**published** ${outputDisplay(meta.path)} (${formatBytes(size)}) — pass it by path to the tool that uses it`;
        } catch (cause) {
          state = `**not published — publishing failed**: ${formatToolError(cause)}\n- every part is stored; resolve that, then call { upload: "${session.id}" } to publish`;
        }
      } else {
        state = `published ${outputDisplay(meta.path)} (${formatBytes(size)}); up to date`;
      }

      const beyond = partNumbers(meta).filter(part => part > meta.total);
      const included = partNumbers(meta).length - beyond.length;
      return textResult(
        config,
        [
          `Upload **${session.id}** → ${outputDisplay(meta.path)}: ${included} of ${meta.total} part(s) stored.`,
          `- ${state}`,
          ...(beyond.length > 0
            ? [
                `- part(s) ${describeRanges(beyond)} are beyond total ${meta.total} and are not included`
              ]
            : []),
          ...(missing.length > 0 ? [`- continue with { upload: "${session.id}", part, text }`] : [])
        ].join("\n")
      );
    });
  }
});

/**
 * Apply one call's changes — a new total, a part, or both — as a single commit.
 *
 * Everything is validated against the state the call would produce before anything
 * is written, so a rejected call changes nothing at all.
 */
async function apply(
  config: ServerConfig,
  session: Session,
  meta: SessionMeta,
  args: { readonly total?: number; readonly part?: number; readonly text?: string }
): Promise<SessionMeta> {
  const total = args.total ?? meta.total;
  let parts = meta.parts;
  let replaced: string | undefined;
  let changed = total !== meta.total;

  if (args.part !== undefined && args.text !== undefined) {
    if (args.part > total) {
      throw toolError.invalidInput(
        `part ${args.part} is beyond total ${total}`,
        "Raise total in this same call if the file really has more parts."
      );
    }
    const bytes = Buffer.from(args.text, "utf8");
    const previous = meta.parts[args.part];
    const unchanged =
      previous !== undefined &&
      previous.size === bytes.byteLength &&
      (await readFile(await sessionFile(config, session.id, previous.file))).equals(bytes);

    if (!unchanged) {
      assertWithinLimit(config, {
        ...meta,
        total,
        parts: { ...meta.parts, [args.part]: { file: "", size: bytes.byteLength } }
      });
      // A new file for every version, so the manifest still pointing at the old one
      // stays true until the commit below replaces it.
      const file = `${String(args.part).padStart(5, "0")}.${randomBytes(6).toString("hex")}`;
      await writeFileAtomic(await sessionFile(config, session.id, file), bytes);
      parts = { ...meta.parts, [args.part]: { file, size: bytes.byteLength } };
      replaced = previous?.file;
      changed = true;
    }
  }
  // Checked for the state the call produces whatever it carried: a new total with an
  // unchanged part, or with no part at all, can bring earlier parts back into range.
  if (parts === meta.parts && total !== meta.total) {
    assertWithinLimit(config, { ...meta, total });
  }

  if (!changed) {
    return meta;
  }
  const next: SessionMeta = { ...meta, total, parts, revision: meta.revision + 1 };
  await commit(config, session.id, next);
  if (replaced !== undefined) {
    // Only after the commit: until then the old file is what the manifest names.
    await unlink(await sessionFile(config, session.id, replaced)).catch(() => undefined);
  }
  return next;
}

/** Refuse a state whose published file would exceed the size limit. Writes nothing. */
function assertWithinLimit(config: ServerConfig, meta: SessionMeta): void {
  const size = includedSize(meta);
  if (size > config.maxFileSize) {
    throw toolError.invalidInput(
      `${outputDisplay(meta.path)} would be ${size} bytes, over the ${config.maxFileSize} byte limit`,
      "Nothing was changed. Lower total, replace large parts, or ask the user to raise --max-file-size."
    );
  }
}

/** Begin a session: validate the destination now, so a doomed upload fails on part 1. */
async function start(
  config: ServerConfig,
  args: { readonly path?: string; readonly total?: number; readonly overwrite?: boolean }
): Promise<Session> {
  if (args.path === undefined || args.total === undefined) {
    throw toolError.invalidInput(
      "starting an upload needs `path` and `total`",
      'e.g. { path: "report.md", total: 6, part: 1, text: "…" }. To continue an upload, pass its `upload` id instead.'
    );
  }
  const target = await resolveOutputPath(config, args.path);
  if (args.overwrite !== true && (await exists(target))) {
    throw toolError.invalidInput(
      `${outputDisplay(args.path)} already exists`,
      "Pass overwrite: true to replace it, or choose a different path."
    );
  }

  await sweepExpired(config);
  const id = randomBytes(10).toString("hex");
  const directory = await resolveOutputPath(config, sessionPath(id));
  await mkdir(directory, { recursive: true });
  await commit(config, id, {
    version: 2,
    path: args.path,
    total: args.total,
    overwrite: args.overwrite === true,
    revision: 0,
    parts: {}
  });
  return { id, directory };
}

/** Continue a session named by id. */
async function open(
  config: ServerConfig,
  args: { readonly upload?: string; readonly path?: string }
): Promise<Session> {
  const id = args.upload as string;
  if (!UPLOAD_ID.test(id)) {
    throw toolError.invalidInput(
      `${JSON.stringify(id)} is not an upload id`,
      "Pass the id exactly as the first call returned it."
    );
  }
  const meta = await readMeta(config, id);
  if (args.path !== undefined && args.path !== meta.path) {
    throw toolError.invalidInput(
      `upload ${id} writes ${outputDisplay(meta.path)}, not ${outputDisplay(args.path)}`,
      "Omit path when continuing an upload, or start a new upload for a different file."
    );
  }
  return { id, directory: await resolveOutputPath(config, sessionPath(id)) };
}
