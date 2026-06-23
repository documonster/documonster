import { EMPTY_UINT8ARRAY } from "@archive/core/bytes";
import { ArchiveError } from "@archive/core/errors";
import type { Zip64Mode } from "@archive/zip-spec/zip-records";
import {
  buildEndOfCentralDirectory,
  buildZip64EndOfCentralDirectory,
  buildZip64EndOfCentralDirectoryLocator,
  buildZip64ExtraField,
  concatExtraFields,
  writeCentralDirectoryHeaderInto,
  writeEndOfCentralDirectoryInto,
  writeZip64EndOfCentralDirectoryInto,
  writeZip64EndOfCentralDirectoryLocatorInto,
  ZIP_CENTRAL_DIR_HEADER_FIXED_SIZE,
  ZIP_END_OF_CENTRAL_DIR_FIXED_SIZE,
  ZIP64_END_OF_CENTRAL_DIR_FIXED_SIZE,
  ZIP64_END_OF_CENTRAL_DIR_LOCATOR_FIXED_SIZE,
  UINT16_MAX,
  UINT32_MAX,
  VERSION_MADE_BY,
  VERSION_NEEDED,
  VERSION_ZIP64
} from "@archive/zip-spec/zip-records";
import type { ZipCentralDirEntry } from "@archive/zip/writable-file";

/**
 * Input type for building Central Directory entries.
 *
 * This is a superset of ZipCentralDirEntry with fields renamed to match
 * the build function naming conventions.
 */
export interface ZipCentralDirectoryEntryInput {
  fileName: Uint8Array;
  extraField: Uint8Array;
  comment: Uint8Array;
  flags: number;
  crc32: number;
  compressedSize: number;
  uncompressedSize: number;
  compressionMethod: number;
  dosTime: number;
  dosDate: number;
  localHeaderOffset: number;
  zip64?: boolean;
  externalAttributes: number;
  versionMadeBy?: number;
}

export interface ZipCentralDirectoryBuildResult {
  centralDirectoryHeaders: Uint8Array[];
  centralDirSize: number;
  trailerRecords: Uint8Array[];
  usedZip64: boolean;
}

export interface ZipCentralDirectoryWriteResult {
  centralDirSize: number;
  trailerSize: number;
  totalWritten: number;
  usedZip64: boolean;
}

export interface ZipCentralDirectorySizingResult {
  centralDirSize: number;
  trailerSize: number;
  totalSize: number;
  usedZip64: boolean;
}

type ZipCentralDirectoryBuildEntry = ZipCentralDirectoryEntryInput | ZipCentralDirEntry;
type ZipCentralDirectoryProcessedEntry = {
  name: Uint8Array;
  extraField: Uint8Array;
  comment: Uint8Array;
  flags: number;
  crc: number;
  compressedData: Uint8Array;
  uncompressedSize: number;
  compressionMethod: number;
  modTime: number;
  modDate: number;
  offset: number;
  externalAttributes: number;
  versionMadeBy?: number;
};

type AnyCentralDirectoryEntry = ZipCentralDirectoryBuildEntry | ZipCentralDirectoryProcessedEntry;

function zip64ExtraLength(
  forceZip64: boolean,
  uncompressedSize: number,
  compressedSize: number,
  localHeaderOffset: number,
  needsZip64Entry: boolean
): number {
  if (!needsZip64Entry) {
    return 0;
  }

  let dataLen = 0;
  if (forceZip64 || uncompressedSize > UINT32_MAX) {
    dataLen += 8;
  }
  if (forceZip64 || compressedSize > UINT32_MAX) {
    dataLen += 8;
  }
  if (forceZip64 || localHeaderOffset > UINT32_MAX) {
    dataLen += 8;
  }

  return dataLen > 0 ? 4 + dataLen : 0;
}

function getEntryShape(entries: AnyCentralDirectoryEntry[]): "input" | "central" | "processed" {
  const first = entries[0] as
    | Partial<ZipCentralDirectoryEntryInput & ZipCentralDirectoryProcessedEntry>
    | undefined;
  return first?.fileName !== undefined
    ? "input"
    : first?.compressedData !== undefined
      ? "processed"
      : "central";
}

/**
 * Fields a central-directory entry exposes after being normalized across the
 * three possible input shapes (`input`, `central`, `processed`).
 */
interface NormalizedCentralDirEntry {
  fileName: Uint8Array;
  extraField: Uint8Array;
  comment?: Uint8Array;
  flags: number;
  crc32: number;
  compressedSize: number;
  uncompressedSize: number;
  compressionMethod: number;
  dosTime: number;
  dosDate: number;
  localHeaderOffset: number;
  zip64: boolean | undefined;
  externalAttributes: number;
  versionMadeBy?: number;
}

/**
 * Collapse the three structurally-distinct entry shapes into a single uniform
 * view so callers don't need shape-specific field access.
 */
function normalizeCentralDirEntry(
  entry: AnyCentralDirectoryEntry,
  shape: "input" | "central" | "processed"
): NormalizedCentralDirEntry {
  if (shape === "processed") {
    const e = entry as ZipCentralDirectoryProcessedEntry;
    return {
      fileName: e.name,
      extraField: e.extraField,
      comment: e.comment,
      flags: e.flags,
      crc32: e.crc,
      compressedSize: e.compressedData.length,
      uncompressedSize: e.uncompressedSize,
      compressionMethod: e.compressionMethod,
      dosTime: e.modTime,
      dosDate: e.modDate,
      localHeaderOffset: e.offset,
      zip64: undefined,
      externalAttributes: e.externalAttributes,
      versionMadeBy: e.versionMadeBy
    };
  }
  if (shape === "input") {
    const e = entry as ZipCentralDirectoryEntryInput;
    return {
      fileName: e.fileName,
      extraField: e.extraField,
      comment: e.comment,
      flags: e.flags,
      crc32: e.crc32,
      compressedSize: e.compressedSize,
      uncompressedSize: e.uncompressedSize,
      compressionMethod: e.compressionMethod,
      dosTime: e.dosTime,
      dosDate: e.dosDate,
      localHeaderOffset: e.localHeaderOffset,
      zip64: e.zip64,
      externalAttributes: e.externalAttributes,
      versionMadeBy: e.versionMadeBy
    };
  }
  const e = entry as ZipCentralDirEntry;
  return {
    fileName: e.name,
    extraField: e.extraField,
    comment: e.comment,
    flags: e.flags,
    crc32: e.crc,
    compressedSize: e.compressedSize,
    uncompressedSize: e.uncompressedSize,
    compressionMethod: e.compressionMethod,
    dosTime: e.dosTime,
    dosDate: e.dosDate,
    localHeaderOffset: e.offset,
    zip64: e.zip64,
    externalAttributes: e.externalAttributes,
    versionMadeBy: e.versionMadeBy
  };
}

export function measureCentralDirectoryAndEocd(
  entries: AnyCentralDirectoryEntry[],
  options: {
    zipComment: Uint8Array;
    zip64Mode: Zip64Mode;
    centralDirOffset: number;
  }
): ZipCentralDirectorySizingResult {
  const forceZip64 = options.zip64Mode === true;
  const forbidZip64 = options.zip64Mode === false;
  const shape = getEntryShape(entries);

  let centralDirSize = 0;
  for (let i = 0; i < entries.length; i++) {
    const entry = normalizeCentralDirEntry(entries[i]!, shape);
    const fileName = entry.fileName;
    const extraFieldBase = entry.extraField;
    const comment = entry.comment;
    const compressedSize = entry.compressedSize;
    const uncompressedSize = entry.uncompressedSize;
    const localHeaderOffset = entry.localHeaderOffset;
    const zip64 = entry.zip64;

    const needsZip64Entry =
      forceZip64 ||
      zip64 === true ||
      localHeaderOffset > UINT32_MAX ||
      compressedSize > UINT32_MAX ||
      uncompressedSize > UINT32_MAX;

    centralDirSize +=
      ZIP_CENTRAL_DIR_HEADER_FIXED_SIZE +
      fileName.length +
      extraFieldBase.length +
      zip64ExtraLength(
        forceZip64,
        uncompressedSize,
        compressedSize,
        localHeaderOffset,
        needsZip64Entry
      ) +
      (comment?.length ?? 0);
  }

  const needsZip64EOCDFromArchive =
    entries.length > UINT16_MAX || options.centralDirOffset > UINT32_MAX;
  const usedZip64 = forceZip64 || needsZip64EOCDFromArchive || centralDirSize > UINT32_MAX;

  if (forbidZip64 && usedZip64) {
    throw new ArchiveError("ZIP64 is required but zip64=false");
  }

  const trailerSize = usedZip64
    ? ZIP64_END_OF_CENTRAL_DIR_FIXED_SIZE +
      ZIP64_END_OF_CENTRAL_DIR_LOCATOR_FIXED_SIZE +
      ZIP_END_OF_CENTRAL_DIR_FIXED_SIZE +
      options.zipComment.length
    : ZIP_END_OF_CENTRAL_DIR_FIXED_SIZE + options.zipComment.length;

  return {
    centralDirSize,
    trailerSize,
    totalSize: centralDirSize + trailerSize,
    usedZip64
  };
}

export function buildCentralDirectoryAndEocd(
  entries: AnyCentralDirectoryEntry[],
  options: {
    zipComment: Uint8Array;
    zip64Mode: Zip64Mode;
    centralDirOffset: number;
  }
): ZipCentralDirectoryBuildResult {
  const forceZip64 = options.zip64Mode === true;
  const forbidZip64 = options.zip64Mode === false;

  const centralDirOffset = options.centralDirOffset;
  const needsZip64EOCDFromArchive = entries.length > UINT16_MAX || centralDirOffset > UINT32_MAX;

  const centralDirectoryHeaders: Uint8Array[] = new Array(entries.length);
  let centralDirSize = 0;

  const emitHeader = (
    index: number,
    fileName: Uint8Array,
    extraFieldBase: Uint8Array,
    comment: Uint8Array,
    flags: number,
    crc32: number,
    compressedSize: number,
    uncompressedSize: number,
    compressionMethod: number,
    dosTime: number,
    dosDate: number,
    localHeaderOffset: number,
    zip64: boolean | undefined,
    externalAttributes: number,
    versionMadeBy: number | undefined
  ): void => {
    const needsZip64Entry =
      forceZip64 ||
      zip64 === true ||
      localHeaderOffset > UINT32_MAX ||
      compressedSize > UINT32_MAX ||
      uncompressedSize > UINT32_MAX;

    const zip64Extra = needsZip64Entry
      ? buildZip64ExtraField({
          uncompressedSize:
            forceZip64 || uncompressedSize > UINT32_MAX ? uncompressedSize : undefined,
          compressedSize: forceZip64 || compressedSize > UINT32_MAX ? compressedSize : undefined,
          localHeaderOffset:
            forceZip64 || localHeaderOffset > UINT32_MAX ? localHeaderOffset : undefined
        })
      : EMPTY_UINT8ARRAY;
    const extraField =
      needsZip64Entry && zip64Extra.length > 0
        ? concatExtraFields(extraFieldBase, zip64Extra)
        : extraFieldBase;

    const header = new Uint8Array(
      ZIP_CENTRAL_DIR_HEADER_FIXED_SIZE +
        fileName.length +
        extraField.length +
        (comment?.length ?? 0)
    );
    const view = new DataView(header.buffer, header.byteOffset, header.byteLength);
    writeCentralDirectoryHeaderInto(header, view, 0, {
      fileName,
      extraField,
      comment: comment ?? EMPTY_UINT8ARRAY,
      flags,
      compressionMethod,
      dosTime,
      dosDate,
      crc32,
      compressedSize: needsZip64Entry ? UINT32_MAX : compressedSize,
      uncompressedSize: needsZip64Entry ? UINT32_MAX : uncompressedSize,
      localHeaderOffset: needsZip64Entry ? UINT32_MAX : localHeaderOffset,
      versionMadeBy: versionMadeBy ?? VERSION_MADE_BY,
      versionNeeded: needsZip64Entry ? VERSION_ZIP64 : VERSION_NEEDED,
      externalAttributes
    });

    centralDirectoryHeaders[index] = header;
    centralDirSize += header.length;
  };

  const shape = getEntryShape(entries);

  if (shape === "input") {
    for (let i = 0; i < entries.length; i++) {
      const entry = entries[i] as ZipCentralDirectoryEntryInput;
      emitHeader(
        i,
        entry.fileName,
        entry.extraField,
        entry.comment,
        entry.flags,
        entry.crc32,
        entry.compressedSize,
        entry.uncompressedSize,
        entry.compressionMethod,
        entry.dosTime,
        entry.dosDate,
        entry.localHeaderOffset,
        entry.zip64,
        entry.externalAttributes,
        entry.versionMadeBy
      );
    }
  } else if (shape === "central") {
    for (let i = 0; i < entries.length; i++) {
      const entry = entries[i] as ZipCentralDirEntry;
      emitHeader(
        i,
        entry.name,
        entry.extraField,
        entry.comment,
        entry.flags,
        entry.crc,
        entry.compressedSize,
        entry.uncompressedSize,
        entry.compressionMethod,
        entry.dosTime,
        entry.dosDate,
        entry.offset,
        entry.zip64,
        entry.externalAttributes,
        entry.versionMadeBy
      );
    }
  } else {
    for (let i = 0; i < entries.length; i++) {
      const entry = entries[i] as ZipCentralDirectoryProcessedEntry;
      emitHeader(
        i,
        entry.name,
        entry.extraField,
        entry.comment,
        entry.flags,
        entry.crc,
        entry.compressedData.length,
        entry.uncompressedSize,
        entry.compressionMethod,
        entry.modTime,
        entry.modDate,
        entry.offset,
        undefined,
        entry.externalAttributes,
        entry.versionMadeBy
      );
    }
  }

  const usedZip64 = forceZip64 || needsZip64EOCDFromArchive || centralDirSize > UINT32_MAX;
  if (forbidZip64 && usedZip64) {
    throw new ArchiveError("ZIP64 is required but zip64=false");
  }

  if (usedZip64) {
    const zip64EocdOffset = centralDirOffset + centralDirSize;
    const zip64Eocd = buildZip64EndOfCentralDirectory({
      entryCountOnDisk: entries.length,
      entryCountTotal: entries.length,
      centralDirSize,
      centralDirOffset
    });
    const zip64Locator = buildZip64EndOfCentralDirectoryLocator({
      zip64EndOfCentralDirectoryOffset: zip64EocdOffset,
      totalDisks: 1
    });

    const eocd = buildEndOfCentralDirectory({
      entryCount: UINT16_MAX,
      centralDirSize: UINT32_MAX,
      centralDirOffset: UINT32_MAX,
      comment: options.zipComment
    });

    return {
      centralDirectoryHeaders,
      centralDirSize,
      trailerRecords: [zip64Eocd, zip64Locator, eocd],
      usedZip64
    };
  }

  const eocd = buildEndOfCentralDirectory({
    entryCount: entries.length,
    centralDirSize,
    centralDirOffset,
    comment: options.zipComment
  });

  return {
    centralDirectoryHeaders,
    centralDirSize,
    trailerRecords: [eocd],
    usedZip64
  };
}

export function writeCentralDirectoryAndEocdInto(
  entries: AnyCentralDirectoryEntry[],
  options: {
    zipComment: Uint8Array;
    zip64Mode: Zip64Mode;
    centralDirOffset: number;
    out: Uint8Array;
    offset: number;
  }
): ZipCentralDirectoryWriteResult {
  const forceZip64 = options.zip64Mode === true;
  const sizing = measureCentralDirectoryAndEocd(entries, {
    zipComment: options.zipComment,
    zip64Mode: options.zip64Mode,
    centralDirOffset: options.centralDirOffset
  });
  const shape = getEntryShape(entries);
  let offset = options.offset;
  const view = new DataView(options.out.buffer, options.out.byteOffset, options.out.byteLength);

  for (let i = 0; i < entries.length; i++) {
    const entry = normalizeCentralDirEntry(entries[i]!, shape);

    const fileName = entry.fileName;
    const extraFieldBase = entry.extraField;
    const comment = entry.comment;
    const flags = entry.flags;
    const crc32 = entry.crc32;
    const compressedSize = entry.compressedSize;
    const uncompressedSize = entry.uncompressedSize;
    const compressionMethod = entry.compressionMethod;
    const dosTime = entry.dosTime;
    const dosDate = entry.dosDate;
    const localHeaderOffset = entry.localHeaderOffset;
    const zip64 = entry.zip64;
    const externalAttributes = entry.externalAttributes;
    const versionMadeBy = entry.versionMadeBy;

    const needsZip64Entry =
      forceZip64 ||
      zip64 === true ||
      localHeaderOffset > UINT32_MAX ||
      compressedSize > UINT32_MAX ||
      uncompressedSize > UINT32_MAX;

    const zip64Extra = needsZip64Entry
      ? buildZip64ExtraField({
          uncompressedSize:
            forceZip64 || uncompressedSize > UINT32_MAX ? uncompressedSize : undefined,
          compressedSize: forceZip64 || compressedSize > UINT32_MAX ? compressedSize : undefined,
          localHeaderOffset:
            forceZip64 || localHeaderOffset > UINT32_MAX ? localHeaderOffset : undefined
        })
      : EMPTY_UINT8ARRAY;
    const extraField =
      needsZip64Entry && zip64Extra.length > 0
        ? concatExtraFields(extraFieldBase, zip64Extra)
        : extraFieldBase;

    offset += writeCentralDirectoryHeaderInto(options.out, view, offset, {
      fileName,
      extraField,
      comment: comment ?? EMPTY_UINT8ARRAY,
      flags,
      compressionMethod,
      dosTime,
      dosDate,
      crc32,
      compressedSize: needsZip64Entry ? UINT32_MAX : compressedSize,
      uncompressedSize: needsZip64Entry ? UINT32_MAX : uncompressedSize,
      localHeaderOffset: needsZip64Entry ? UINT32_MAX : localHeaderOffset,
      versionMadeBy: versionMadeBy ?? VERSION_MADE_BY,
      versionNeeded: needsZip64Entry ? VERSION_ZIP64 : VERSION_NEEDED,
      externalAttributes
    });
  }

  if (sizing.usedZip64) {
    const zip64EocdOffset = options.centralDirOffset + sizing.centralDirSize;
    offset += writeZip64EndOfCentralDirectoryInto(options.out, view, offset, {
      entryCountOnDisk: entries.length,
      entryCountTotal: entries.length,
      centralDirSize: sizing.centralDirSize,
      centralDirOffset: options.centralDirOffset
    });

    offset += writeZip64EndOfCentralDirectoryLocatorInto(options.out, view, offset, {
      zip64EndOfCentralDirectoryOffset: zip64EocdOffset,
      totalDisks: 1
    });

    writeEndOfCentralDirectoryInto(options.out, view, offset, {
      entryCount: UINT16_MAX,
      centralDirSize: UINT32_MAX,
      centralDirOffset: UINT32_MAX,
      comment: options.zipComment
    });
  } else {
    writeEndOfCentralDirectoryInto(options.out, view, offset, {
      entryCount: entries.length,
      centralDirSize: sizing.centralDirSize,
      centralDirOffset: options.centralDirOffset,
      comment: options.zipComment
    });
  }

  return {
    centralDirSize: sizing.centralDirSize,
    trailerSize: sizing.trailerSize,
    totalWritten: sizing.totalSize,
    usedZip64: sizing.usedZip64
  };
}
