/**
 * How a read reports a package that is not a readable workbook, so every Excel reader — buffered XLSX and XLSB,
 * streamed XLSX, `Stream.WorkbookReader` — reports it the same way: bytes that are not a ZIP package, and a part whose
 * bytes do not match the CRC-32 the package records for it, are an `ExcelFileError` (which `isExcelError` answers for)
 * with the archive's diagnosis as the cause — and an OLE compound file, an encrypted workbook (`ExcelEncryptedError`)
 * or a legacy `.xls`, is reported as what it is. A container limit is an `ExcelLimitExceededError` naming the `zip`
 * option, with the archive's `ArchiveLimitExceededError` as its cause; a ZIP password keeps its archive error.
 *
 * An encrypted workbook read with its `password` is decrypted here ({@link openWorkbookPackage}), by the Office
 * encryption Word uses too (`@archive/office-crypto/agile`) — loaded only then, through `import()`.
 */
import { hasEncryptedPackageStreams, isCfb } from "@archive/cfb/signature";
import { DEFAULT_MAX_ENTRY_SIZE } from "@archive/core/defaults";
import {
  ArchiveFileTooLargeError,
  ArchiveLimitExceededError,
  ZipCrc32MismatchError,
  ZipParseError
} from "@archive/core/errors";
import type { EncryptOptions } from "@archive/office-crypto/agile";
import {
  ExcelDecryptionError,
  ExcelEncryptedError,
  ExcelFileError,
  ExcelLimitExceededError
} from "@excel/errors";
import type { ExcelLimit } from "@excel/errors";

/** Whether `bytes` start with the compound-file signature. */
export const isCompoundFile = isCfb;

/**
 * The decrypted package's default bound when the read sets none: the default of `zip.maxTotalUncompressedSize`, as a
 * package (compressed) cannot be larger than its parts uncompressed.
 */
const DEFAULT_MAX_PACKAGE_SIZE = 2 * 1024 * 1024 * 1024;

/** What decrypting a workbook needs: the read's password and size bound, and the names its errors give. */
export interface WorkbookDecryption {
  readonly password?: string | undefined;
  /** Bound on the decrypted package. Default: 2 GiB, `zip.maxTotalUncompressedSize`'s default. */
  readonly maxPackageSize?: number | undefined;
  readonly source?: string | undefined;
  readonly format?: string | undefined;
}

/** The decryption a read's options ask for. */
export function decryptionOf(
  options: { password?: string; zip?: { maxTotalUncompressedSize?: number } } | undefined,
  source?: string
): WorkbookDecryption {
  return {
    password: options?.password,
    maxPackageSize: options?.zip?.maxTotalUncompressedSize,
    source
  };
}

/**
 * The workbook package in `bytes`: an encrypted workbook decrypted with `decryption.password` — an
 * {@link ExcelEncryptedError} without one, an {@link ExcelDecryptionError} with a wrong one — and a legacy `.xls` the
 * {@link compoundFileError} it is. Anything else is returned as it is.
 */
export async function openWorkbookPackage(
  bytes: Uint8Array,
  decryption: WorkbookDecryption
): Promise<Uint8Array> {
  if (!isCfb(bytes)) {
    return bytes;
  }
  const source = decryption.source ?? "<input>";
  if (!hasEncryptedPackageStreams(bytes) || decryption.password === undefined) {
    throw compoundFileError(bytes, source, decryption.format);
  }
  const { CfbError, OfficeCryptoError, decryptPackage } =
    await import("@archive/office-crypto/agile");
  try {
    return await decryptPackage(
      bytes,
      decryption.password,
      decryption.maxPackageSize ?? DEFAULT_MAX_PACKAGE_SIZE
    );
  } catch (error) {
    if (error instanceof OfficeCryptoError) {
      throw new ExcelDecryptionError(source, error.reason, decryptionDetails(error), {
        cause: error
      });
    }
    if (error instanceof CfbError) {
      throw new ExcelDecryptionError(
        source,
        "malformed",
        `the encrypted workbook's container is damaged: ${error.message}`,
        { cause: error }
      );
    }
    throw error;
  }
}

function decryptionDetails(error: { reason: string; message: string }): string {
  switch (error.reason) {
    case "password":
      return "the password is incorrect";
    case "unsupported":
      return `the workbook's encryption is not supported: ${error.message}`;
    case "limit":
      return `the decrypted package is larger than the read allows (options.zip.maxTotalUncompressedSize): ${error.message}`;
    default:
      return `the encrypted workbook cannot be decrypted: ${error.message}`;
  }
}

/**
 * `bytes`, encrypted with `options.password` when one is given — the CFB container Excel writes for "Encrypt with
 * Password", by the encryption Word uses too, loaded only then.
 */
export async function encryptWorkbookPackage(
  bytes: Uint8Array,
  options:
    | {
        readonly password?: string;
        readonly encryption?: EncryptOptions;
      }
    | undefined
): Promise<Uint8Array> {
  if (options?.password === undefined) {
    return bytes;
  }
  const { encryptPackage } = await import("@archive/office-crypto/agile");
  return encryptPackage(bytes, options.password, options.encryption);
}

/**
 * The error for a compound file read as a workbook package: `ExcelEncryptedError` when it holds the `EncryptionInfo`
 * and `EncryptedPackage` streams of an encrypted OOXML package (MS-OFFCRYPTO 2.3.4 — what Excel writes for a file with
 * a password to open), else an `ExcelFileError` saying it is a compound file and not a ZIP package. The directory
 * names are found as the directory stores them, UTF-16LE.
 */
export function compoundFileError(
  bytes: Uint8Array,
  source = "<input>",
  format: string | undefined = "XLSX"
): ExcelFileError {
  if (hasEncryptedPackageStreams(bytes)) {
    return new ExcelEncryptedError(source);
  }
  return new ExcelFileError(
    source,
    "read",
    `not an ${format} package — an OLE compound file: a legacy .xls workbook (BIFF) or another Office binary format, ` +
      "which this library does not read"
  );
}

/**
 * Looks at the first bytes of a streamed input, and keeps a compound file away from the ZIP parser: its chunks are
 * collected, and {@link end} hands on the package an encrypted workbook decrypts to with the read's password — or
 * throws {@link openWorkbookPackage}'s error for the whole input. Anything else is passed through as it arrives (the
 * first eight bytes held back until they are known).
 */
export class CompoundFileSniffer {
  private head: Uint8Array[] = [];
  private headLength = 0;
  private state: "head" | "pass" | "compound" = "head";
  private readonly decryption: WorkbookDecryption;

  constructor(decryption: WorkbookDecryption = {}) {
    this.decryption = { source: "<input>", format: "XLSX", ...decryption };
  }

  /** The chunks to hand on for `chunk` — none while the head is undecided or the input is a compound file. */
  push(chunk: Uint8Array | string): (Uint8Array | string)[] {
    if (this.state === "pass") {
      return [chunk];
    }
    if (typeof chunk === "string") {
      // Text is not a compound file.
      this.state = "pass";
      return this.release(chunk);
    }
    this.head.push(chunk);
    this.headLength += chunk.length;
    if (this.state === "compound" || this.headLength < 8) {
      return [];
    }
    if (isCfb(concat(this.head, this.headLength))) {
      this.state = "compound";
      return [];
    }
    this.state = "pass";
    return this.release();
  }

  /**
   * The chunks still held back at the end of the input: for a compound file, the package an encrypted workbook
   * decrypts to; it throws for a compound file that cannot be read.
   */
  async end(): Promise<Uint8Array[]> {
    if (this.state === "compound") {
      const bytes = concat(this.head, this.headLength);
      this.head = [];
      this.headLength = 0;
      return [await openWorkbookPackage(bytes, this.decryption)];
    }
    return this.release() as Uint8Array[];
  }

  private release(...after: (Uint8Array | string)[]): (Uint8Array | string)[] {
    const held = this.head;
    this.head = [];
    this.headLength = 0;
    return [...held, ...after];
  }
}

/**
 * `chunks` passed through a {@link CompoundFileSniffer}: an encrypted workbook is decrypted once it has been read, and
 * a compound file that cannot be read throws its error.
 */
export async function* compoundFileChecked<T extends Uint8Array | string>(
  chunks: AsyncIterable<T>,
  decryption?: WorkbookDecryption
): AsyncGenerator<T> {
  const sniffer = new CompoundFileSniffer(decryption);
  for await (const chunk of chunks) {
    yield* sniffer.push(chunk) as T[];
  }
  yield* (await sniffer.end()) as T[];
}

function concat(chunks: readonly Uint8Array[], length: number): Uint8Array {
  if (chunks.length === 1) {
    return chunks[0]!;
  }
  const out = new Uint8Array(length);
  let at = 0;
  for (const chunk of chunks) {
    out.set(chunk, at);
    at += chunk.length;
  }
  return out;
}

/**
 * Document-level container defaults. The archive library leaves both bounds
 * unbounded because what a process can afford is host-dependent; a workbook is
 * a narrower thing. A part count of 10,000 mirrors the Word reader's
 * `maxPartCount` and is far above real workbooks (each sheet adds ~2–6 parts:
 * sheet XML, rels, drawing, comments, table), which reach the low hundreds.
 * 2 GiB total sits above what Excel itself can open — sheet XML compresses
 * ~10:1, so this admits a ~200 MB `.xlsx` — while rejecting the many-entries
 * bombs a per-entry bound alone lets through. Pass `Infinity` to lift either.
 */
const DEFAULT_MAX_ENTRIES = 10_000;
const DEFAULT_MAX_TOTAL_UNCOMPRESSED_SIZE = 2 * 1024 * 1024 * 1024;

/** The `zip` options of a workbook read, as every Excel reader takes them. */
export interface ContainerReadOptions {
  readonly maxEntrySize?: number;
  readonly maxTotalUncompressedSize?: number;
  readonly maxEntries?: number;
  readonly checkCrc32?: boolean;
}

/**
 * The container limits a read applies — its `zip` options with the workbook defaults filled in. One function so
 * the buffered XLSX and XLSB readers, format detection and the streamed reader cannot disagree about a limit.
 */
export function containerReadLimits(zip: ContainerReadOptions | undefined): {
  maxEntrySize: number;
  maxTotalUncompressedSize: number;
  maxEntries: number;
  checkCrc32?: boolean;
} {
  return {
    maxEntrySize: zip?.maxEntrySize ?? DEFAULT_MAX_ENTRY_SIZE,
    maxTotalUncompressedSize: zip?.maxTotalUncompressedSize ?? DEFAULT_MAX_TOTAL_UNCOMPRESSED_SIZE,
    maxEntries: zip?.maxEntries ?? DEFAULT_MAX_ENTRIES,
    ...(zip?.checkCrc32 === undefined ? {} : { checkCrc32: zip.checkCrc32 })
  };
}

/**
 * The limits of {@link containerReadLimits} applied to a package read as it streams, where sizes are not known up
 * front: {@link enter} counts a part as it arrives and returns the meter its bytes go through as they inflate —
 * every byte of the part, whether it is parsed or drained. Each refusal is the {@link containerLimitError} the
 * buffered reader would give.
 */
export function createContainerBudget(zip: ContainerReadOptions | undefined): {
  /** Count one more part, `path`, and return the meter its chunks are charged to. */
  enter(path: string): (chunk: Uint8Array | string) => void;
} {
  const limits = containerReadLimits(zip);
  let entries = 0;
  let total = 0;
  return {
    enter(path) {
      entries++;
      if (entries > limits.maxEntries) {
        throw containerLimitError(new ArchiveLimitExceededError("maxEntries", limits.maxEntries));
      }
      let partTotal = 0;
      return chunk => {
        // An entry's chunks are bytes; a string chunk is counted as the UTF-8 it stands for.
        const size =
          typeof chunk === "string" ? new TextEncoder().encode(chunk).length : chunk.byteLength;
        partTotal += size;
        if (partTotal > limits.maxEntrySize) {
          throw containerLimitError(
            new ArchiveFileTooLargeError(path, `holds more than ${limits.maxEntrySize} bytes`, {
              allowed: limits.maxEntrySize
            })
          );
        }
        total += size;
        if (total > limits.maxTotalUncompressedSize) {
          throw containerLimitError(
            new ArchiveLimitExceededError(
              "maxTotalUncompressedSize",
              limits.maxTotalUncompressedSize
            )
          );
        }
      };
    }
  };
}

/**
 * The error a read reports for a container limit the archive enforced: an {@link ExcelLimitExceededError} naming the
 * `zip` option and its value, with the archive's error — which names the part, for a part too large — as the cause.
 */
export function containerLimitError(error: ArchiveLimitExceededError): ExcelLimitExceededError {
  const limit: ExcelLimit | undefined =
    error.limit === "maxEntrySize" ||
    error.limit === "maxTotalUncompressedSize" ||
    error.limit === "maxEntries"
      ? error.limit
      : undefined;
  const itemType =
    error instanceof ArchiveFileTooLargeError
      ? `part "${error.path}"`
      : limit === "maxEntries"
        ? "part"
        : "package";
  return new ExcelLimitExceededError(itemType, error.allowed, { limit, cause: error });
}

/**
 * The error a read reports for `error`: an `ExcelFileError` for a package that is damaged or not one — the
 * {@link compoundFileError} when `bytes`, the input read, are a compound file — an `ExcelLimitExceededError` for a
 * container limit, else `error`.
 */
export function containerReadError(
  error: unknown,
  source = "<input>",
  format = "XLSX",
  bytes?: Uint8Array
): unknown {
  if (error instanceof ArchiveLimitExceededError) {
    return containerLimitError(error);
  }
  if (error instanceof ZipParseError && bytes !== undefined && isCompoundFile(bytes)) {
    return compoundFileError(bytes, source, format);
  }
  if (error instanceof ZipCrc32MismatchError) {
    return new ExcelFileError(
      source,
      "read",
      `the part "${error.path}" is damaged: its bytes do not match the CRC-32 the package records for it`,
      { cause: error }
    );
  }
  if (error instanceof ZipParseError) {
    return new ExcelFileError(source, "read", `not an ${format} package — ${error.message}`, {
      cause: error
    });
  }
  return error;
}

/** `chunks`, with a damaged or unreadable package reported as {@link containerReadError} reports it. */
export async function* withContainerErrors<T>(
  chunks: AsyncIterable<T>,
  source?: string
): AsyncGenerator<T> {
  try {
    yield* chunks;
  } catch (error) {
    throw containerReadError(error, source);
  }
}
