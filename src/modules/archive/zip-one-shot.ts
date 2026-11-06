/**
 * One-shot ZIP helpers — the ZIP counterparts of `packTar` / `unpackTar`: entries in,
 * bytes out, and bytes in, a path → entry map out.
 *
 * Use `Archive.createZip()` / `Archive.openZip()` instead when you need per-entry options,
 * progress, streaming output or lazy reads.
 */

import {
  checkReadOptions,
  checkZipEncryptionOptions,
  checkZipOptions,
  ZIP_READ_LIMITS
} from "@archive/core/entry-checks";
import { ArchiveArgumentError } from "@archive/core/errors";
import { collect } from "@archive/io/archive-sink";
import type { ArchiveBytes, ArchiveSource } from "@archive/io/archive-source";
import {
  archiveBytesToUint8Array,
  checkSourceArgument,
  isArchiveBytes,
  isInMemoryArchiveSource,
  iterateArchiveSource,
  toUint8Array
} from "@archive/io/archive-source";
import type { TarEntryInput } from "@archive/tar/tar-archive";
import { forEachEntryInput } from "@archive/tar/tar-archive";
import type { ZipEntryRecord, ZipParseOptions } from "@archive/unzip/zip-parser";
import { ZipParser } from "@archive/unzip/zip-parser";
import type { ZipOptions } from "@archive/zip/zip-archive";
import { ZipArchive } from "@archive/zip/zip-archive";
import { asPlainBytes } from "@utils/binary";

/** One extracted ZIP entry: its central-directory record and its uncompressed data. */
export interface UnpackedZipEntry {
  info: ZipEntryRecord;
  data: Uint8Array<ArrayBuffer>;
}

/** Options for {@link unpackZip} / {@link unpackZipSync}. */
export interface UnpackZipOptions extends ZipParseOptions {
  /** Abort signal, honoured while a streamed source is being collected. */
  signal?: AbortSignal;
}

function addZipEntries(
  archive: ZipArchive,
  entries: ReadonlyMap<string, ArchiveSource> | Iterable<TarEntryInput>,
  where: string
): void {
  forEachEntryInput(where, entries, (name, source) => archive._addAt(where, name, source));
}

/**
 * Build a ZIP archive from entries in one call.
 *
 * @example
 * ```ts
 * import { Archive } from "documonster/archive";
 *
 * const bytes = await Archive.packZip([["hello.txt", "Hello"]], { level: 9 });
 * ```
 */
export async function packZip(
  entries: ReadonlyMap<string, ArchiveSource> | Iterable<TarEntryInput>,
  options?: Omit<ZipOptions, "format">
): Promise<Uint8Array<ArrayBuffer>> {
  checkZipOptions("Archive.packZip", options);
  checkZipEncryptionOptions("Archive.packZip", options);
  const archive = new ZipArchive(options);
  addZipEntries(archive, entries, "Archive.packZip");
  return archive.toBuffer();
}

/**
 * Build a ZIP archive from in-memory entries synchronously. Sources must be strings,
 * `Uint8Array`s or `ArrayBuffer`s. `"zipcrypto"` encryption is written synchronously; AES needs
 * {@link packZip}, as it runs on Web Crypto.
 */
export function packZipSync(
  entries: ReadonlyMap<string, string | Uint8Array | ArrayBuffer> | Iterable<TarEntryInput>,
  options?: Omit<ZipOptions, "format">
): Uint8Array<ArrayBuffer> {
  checkZipOptions("Archive.packZipSync", options);
  checkZipEncryptionOptions("Archive.packZipSync", options);
  const archive = new ZipArchive(options);
  addZipEntries(archive, entries, "Archive.packZipSync");
  return archive.toBufferSync();
}

function extractAllEntries(
  parser: ZipParser,
  read: (path: string) => Uint8Array | null
): Map<string, UnpackedZipEntry> {
  const result = new Map<string, UnpackedZipEntry>();
  for (const info of parser.getEntries()) {
    result.set(info.path, { info, data: asPlainBytes(read(info.path) ?? new Uint8Array(0)) });
  }
  return result;
}

/**
 * Extract every entry of a ZIP archive into a path → `{ info, data }` map — the ZIP
 * counterpart of `Archive.unpackTar`. Directories map to empty data.
 *
 * @example
 * ```ts
 * import { Archive } from "documonster/archive";
 *
 * const zip = await Archive.packZip([["a.txt", "A"]]);
 * const files = await Archive.unpackZip(zip);
 * files.get("a.txt")?.data; // Uint8Array [65]
 * ```
 *
 * @param source - A string is the archive's content, encoded as UTF-8 — never a file path or URL;
 *   read a file's bytes first (in Node, or open it with `ArchiveFile.fromFile`).
 */
export async function unpackZip(
  source: ArchiveSource | ArchiveBytes,
  options: UnpackZipOptions = {}
): Promise<Map<string, UnpackedZipEntry>> {
  checkSourceArgument("Archive.unpackZip", source);
  checkReadOptions("Archive.unpackZip", options, ZIP_READ_LIMITS);
  const { signal, ...parseOptions } = options;
  const bytes = isInMemoryArchiveSource(source)
    ? await toUint8Array(source)
    : await collect(iterateArchiveSource(source, { signal }));
  const parser = new ZipParser(bytes, parseOptions);
  const all = await parser.extractAll();
  return extractAllEntries(parser, path => all.get(path) ?? null);
}

/**
 * Synchronous {@link unpackZip} for an in-memory archive. AES-encrypted entries need the
 * asynchronous form; ZipCrypto entries work with `options.password`.
 */
export function unpackZipSync(
  source: ArchiveBytes,
  options: Omit<UnpackZipOptions, "signal"> = {}
): Map<string, UnpackedZipEntry> {
  if (!isArchiveBytes(source)) {
    throw new ArchiveArgumentError(
      "source",
      "Archive.unpackZipSync: source must be bytes (an ArrayBufferView or ArrayBuffer); use Archive.unpackZip for a Blob or a stream"
    );
  }
  checkReadOptions("Archive.unpackZipSync", options, ZIP_READ_LIMITS);
  const parser = new ZipParser(archiveBytesToUint8Array(source), options);
  const all = parser.extractAllSync();
  return extractAllEntries(parser, path => all[path] ?? null);
}
