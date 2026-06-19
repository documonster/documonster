/**
 * TAR Header Utilities
 *
 * Functions for encoding and decoding TAR headers.
 */

import { EMPTY_UINT8ARRAY } from "@archive/core/bytes";
import { ArchiveError } from "@archive/core/errors";
import type { TarType } from "@archive/tar/tar-constants";
import {
  TAR_BLOCK_SIZE,
  TAR_HEADER,
  TAR_TYPE,
  USTAR_MAGIC,
  USTAR_VERSION,
  DEFAULT_TAR_MODE,
  DEFAULT_TAR_DIR_MODE,
  DEFAULT_TAR_UID,
  DEFAULT_TAR_GID,
  DEFAULT_TAR_UNAME,
  DEFAULT_TAR_GNAME
} from "@archive/tar/tar-constants";
import type { TarEntryInfo } from "@archive/tar/tar-entry-info";
import { textEncoder, textDecoder } from "@utils/binary";

/**
 * Encode a string to a fixed-size field (null-terminated if space allows)
 */
function encodeString(value: string, size: number): Uint8Array {
  const result = new Uint8Array(size);
  const bytes = textEncoder.encode(value);
  const copyLen = Math.min(bytes.length, size - 1); // Leave room for null terminator
  result.set(bytes.subarray(0, copyLen));
  return result;
}

/**
 * Decode a null-terminated string from a field
 */
function decodeString(data: Uint8Array, offset: number, size: number): string {
  const end = data.indexOf(0, offset);
  const actualEnd = end >= offset && end < offset + size ? end : offset + size;
  return textDecoder.decode(data.subarray(offset, actualEnd)).trim();
}

/**
 * Encode an octal number to a fixed-size field
 * Format: space-padded octal digits followed by null or space
 */
function encodeOctal(value: number, size: number): Uint8Array {
  const result = new Uint8Array(size);
  // Convert to octal string, pad with leading zeros
  const str = value.toString(8).padStart(size - 1, "0");
  const bytes = textEncoder.encode(str);
  result.set(bytes.subarray(0, size - 1));
  result[size - 1] = 0; // Null terminator
  return result;
}

/**
 * Encode a large number (for sizes > 8GB) using GNU binary format
 */
function encodeLargeNumber(value: number, size: number): Uint8Array {
  if (value < 0o77777777777) {
    // Fits in octal
    return encodeOctal(value, size);
  }

  // Use GNU binary format: high bit set, followed by big-endian number
  const result = new Uint8Array(size);
  result[0] = 0x80; // High bit set indicates binary format

  // Write as big-endian (right-aligned)
  let remaining = value;
  for (let i = size - 1; i > 0; i--) {
    result[i] = remaining & 0xff;
    remaining = Math.floor(remaining / 256);
  }

  return result;
}

/**
 * Decode an octal number from a field
 */
function decodeOctal(data: Uint8Array, offset: number, size: number): number {
  // Check for GNU binary format (high bit set)
  if (data[offset] & 0x80) {
    // Binary format: big-endian number
    let value = 0;
    for (let i = offset + 1; i < offset + size; i++) {
      value = value * 256 + data[i];
    }
    return value;
  }

  // Standard octal format - remove null bytes
  const str = decodeString(data, offset, size).split(String.fromCharCode(0)).join("").trim();
  return str ? parseInt(str, 8) : 0;
}

/**
 * Calculate TAR header checksum
 * The checksum is the sum of all bytes in the header, with the checksum field
 * treated as spaces (0x20).
 */
function calculateChecksum(header: Uint8Array): number {
  let sum = 0;
  for (let i = 0; i < TAR_BLOCK_SIZE; i++) {
    // Treat checksum field (bytes 148-155) as spaces
    if (
      i >= TAR_HEADER.checksum.offset &&
      i < TAR_HEADER.checksum.offset + TAR_HEADER.checksum.size
    ) {
      sum += 0x20;
    } else {
      sum += header[i];
    }
  }
  return sum;
}

/**
 * Validate header checksum
 */
export function validateChecksum(header: Uint8Array): boolean {
  const expected = decodeOctal(header, TAR_HEADER.checksum.offset, TAR_HEADER.checksum.size);
  const actual = calculateChecksum(header);
  return expected === actual;
}

/**
 * Check if a block is all zeros (end of archive marker)
 * Optimized to check 4 bytes at a time using DataView
 */
export function isZeroBlock(block: Uint8Array): boolean {
  // Fast path: check using 32-bit reads where possible
  const len = block.length;
  const aligned = len & ~3; // Round down to multiple of 4
  const view = new DataView(block.buffer, block.byteOffset, block.byteLength);

  for (let i = 0; i < aligned; i += 4) {
    if (view.getUint32(i, true) !== 0) {
      return false;
    }
  }

  // Check remaining bytes
  for (let i = aligned; i < len; i++) {
    if (block[i] !== 0) {
      return false;
    }
  }
  return true;
}

/**
 * Split a long path into prefix and name for ustar format
 * Returns null if the path is too long even with prefix
 */
function splitPath(path: string): { prefix: string; name: string } | null {
  const maxName = TAR_HEADER.name.size - 1; // 99 chars
  const maxPrefix = TAR_HEADER.prefix.size - 1; // 154 chars
  const maxTotal = maxName + maxPrefix + 1; // +1 for separator

  if (path.length <= maxName) {
    return { prefix: "", name: path };
  }

  if (path.length > maxTotal) {
    return null; // Too long even with prefix
  }

  // Find a good split point (directory separator within valid range)
  const splitStart = path.length - maxName;
  const splitEnd = Math.min(path.length, maxPrefix + 1);

  for (let i = splitEnd; i >= splitStart; i--) {
    if (path[i] === "/") {
      return {
        prefix: path.substring(0, i),
        name: path.substring(i + 1)
      };
    }
  }

  return null; // No valid split point found
}

export interface TarHeaderOptions {
  path: string;
  size?: number;
  mode?: number;
  uid?: number;
  gid?: number;
  mtime?: Date;
  type?: TarType;
  linkname?: string;
  uname?: string;
  gname?: string;
  devmajor?: number;
  devminor?: number;
}

/**
 * Encode a TAR header
 *
 * Returns { header, longName? } where longName is an optional GNU long name entry
 * if the filename exceeds ustar limits.
 */
export function encodeHeader(options: TarHeaderOptions): {
  header: Uint8Array;
  longName?: Uint8Array;
} {
  const {
    path,
    size = 0,
    mode = path.endsWith("/") ? DEFAULT_TAR_DIR_MODE : DEFAULT_TAR_MODE,
    uid = DEFAULT_TAR_UID,
    gid = DEFAULT_TAR_GID,
    mtime = new Date(),
    type = path.endsWith("/") ? TAR_TYPE.DIRECTORY : TAR_TYPE.FILE,
    linkname = "",
    uname = DEFAULT_TAR_UNAME,
    gname = DEFAULT_TAR_GNAME,
    devmajor = 0,
    devminor = 0
  } = options;

  let longNameHeader: Uint8Array | undefined;
  let name: string;
  let prefix: string;

  // Try to split path for ustar format
  const split = splitPath(path);

  if (split) {
    name = split.name;
    prefix = split.prefix;
  } else {
    // Path too long - use GNU long name extension
    name = path.substring(0, TAR_HEADER.name.size - 1);
    prefix = "";
    longNameHeader = createGnuLongNameHeader(path);
  }

  // Create header block
  const header = new Uint8Array(TAR_BLOCK_SIZE);

  // Name
  header.set(encodeString(name, TAR_HEADER.name.size), TAR_HEADER.name.offset);

  // Mode
  header.set(encodeOctal(mode, TAR_HEADER.mode.size), TAR_HEADER.mode.offset);

  // UID/GID
  header.set(encodeOctal(uid, TAR_HEADER.uid.size), TAR_HEADER.uid.offset);
  header.set(encodeOctal(gid, TAR_HEADER.gid.size), TAR_HEADER.gid.offset);

  // Size
  header.set(encodeLargeNumber(size, TAR_HEADER.size.size), TAR_HEADER.size.offset);

  // Modification time (seconds since epoch)
  const mtimeSecs = Math.floor(mtime.getTime() / 1000);
  header.set(encodeOctal(mtimeSecs, TAR_HEADER.mtime.size), TAR_HEADER.mtime.offset);

  // Type flag
  header.set(textEncoder.encode(type), TAR_HEADER.type.offset);

  // Link name
  if (linkname) {
    header.set(encodeString(linkname, TAR_HEADER.linkname.size), TAR_HEADER.linkname.offset);
  }

  // USTAR magic and version
  header.set(textEncoder.encode(USTAR_MAGIC), TAR_HEADER.magic.offset);
  header.set(textEncoder.encode(USTAR_VERSION), TAR_HEADER.version.offset);

  // User/group names
  header.set(encodeString(uname, TAR_HEADER.uname.size), TAR_HEADER.uname.offset);
  header.set(encodeString(gname, TAR_HEADER.gname.size), TAR_HEADER.gname.offset);

  // Device numbers
  header.set(encodeOctal(devmajor, TAR_HEADER.devmajor.size), TAR_HEADER.devmajor.offset);
  header.set(encodeOctal(devminor, TAR_HEADER.devminor.size), TAR_HEADER.devminor.offset);

  // Prefix
  if (prefix) {
    header.set(encodeString(prefix, TAR_HEADER.prefix.size), TAR_HEADER.prefix.offset);
  }

  // Calculate and write checksum
  const checksum = calculateChecksum(header);
  // Checksum format: 6 octal digits, null, space
  const checksumStr = checksum.toString(8).padStart(6, "0") + "\0 ";
  header.set(textEncoder.encode(checksumStr), TAR_HEADER.checksum.offset);

  return { header, longName: longNameHeader };
}

/**
 * Create a GNU long name header entry
 */
function createGnuLongNameHeader(path: string): Uint8Array {
  const nameBytes = textEncoder.encode(path + "\0");
  const paddedSize = Math.ceil(nameBytes.length / TAR_BLOCK_SIZE) * TAR_BLOCK_SIZE;

  // Create the long name header
  const { header } = encodeHeader({
    path: "././@LongLink",
    size: nameBytes.length,
    type: TAR_TYPE.GNU_LONG_NAME,
    mode: 0,
    mtime: new Date(0)
  });

  // Create data blocks
  const data = new Uint8Array(paddedSize);
  data.set(nameBytes);

  // Combine header and data
  const result = new Uint8Array(TAR_BLOCK_SIZE + paddedSize);
  result.set(header);
  result.set(data, TAR_BLOCK_SIZE);

  return result;
}

/**
 * Decode a TAR header
 */
export function decodeHeader(header: Uint8Array): TarEntryInfo | null {
  // Check for zero block (end of archive)
  if (isZeroBlock(header)) {
    return null;
  }

  // Validate checksum
  if (!validateChecksum(header)) {
    throw new ArchiveError("Invalid TAR header checksum");
  }

  // Decode basic fields
  const name = decodeString(header, TAR_HEADER.name.offset, TAR_HEADER.name.size);
  const prefix = decodeString(header, TAR_HEADER.prefix.offset, TAR_HEADER.prefix.size);
  const path = prefix ? `${prefix}/${name}` : name;

  const type = (String.fromCharCode(header[TAR_HEADER.type.offset]) || TAR_TYPE.FILE) as TarType;
  const size = decodeOctal(header, TAR_HEADER.size.offset, TAR_HEADER.size.size);
  const mode = decodeOctal(header, TAR_HEADER.mode.offset, TAR_HEADER.mode.size);
  const uid = decodeOctal(header, TAR_HEADER.uid.offset, TAR_HEADER.uid.size);
  const gid = decodeOctal(header, TAR_HEADER.gid.offset, TAR_HEADER.gid.size);
  const mtimeSecs = decodeOctal(header, TAR_HEADER.mtime.offset, TAR_HEADER.mtime.size);
  const mtime = new Date(mtimeSecs * 1000);

  const linkname = decodeString(header, TAR_HEADER.linkname.offset, TAR_HEADER.linkname.size);
  const uname = decodeString(header, TAR_HEADER.uname.offset, TAR_HEADER.uname.size);
  const gname = decodeString(header, TAR_HEADER.gname.offset, TAR_HEADER.gname.size);
  const devmajor = decodeOctal(header, TAR_HEADER.devmajor.offset, TAR_HEADER.devmajor.size);
  const devminor = decodeOctal(header, TAR_HEADER.devminor.offset, TAR_HEADER.devminor.size);

  return {
    path,
    type,
    size,
    mode,
    uid,
    gid,
    uname,
    gname,
    mtime,
    linkname: linkname || undefined,
    devmajor: devmajor || undefined,
    devminor: devminor || undefined
  };
}

/**
 * Calculate padding needed to reach next block boundary
 */
export function calculatePadding(size: number): number {
  const remainder = size % TAR_BLOCK_SIZE;
  return remainder === 0 ? 0 : TAR_BLOCK_SIZE - remainder;
}

// Pre-allocated padding buffers (512 bytes max padding)
const PADDING_CACHE: Uint8Array[] = [];

/**
 * Create padding bytes (uses cached buffers to avoid allocation)
 */
export function createPadding(size: number): Uint8Array {
  const padding = calculatePadding(size);
  if (padding === 0) {
    return EMPTY_UINT8ARRAY;
  }

  // Cache padding buffers by size
  if (!PADDING_CACHE[padding]) {
    PADDING_CACHE[padding] = new Uint8Array(padding);
  }
  return PADDING_CACHE[padding];
}

// Pre-allocated end-of-archive marker
const END_OF_ARCHIVE = new Uint8Array(TAR_BLOCK_SIZE * 2);

/**
 * Create end-of-archive marker (two zero blocks)
 */
export function createEndOfArchive(): Uint8Array {
  return END_OF_ARCHIVE;
}
