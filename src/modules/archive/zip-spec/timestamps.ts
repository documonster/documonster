import { ArchiveError } from "@archive/core/errors";
import { concatUint8Arrays } from "@utils/binary";

export const EXTENDED_TIMESTAMP_ID = 0x5455;
export const NTFS_TIMESTAMP_ID = 0x000a;

export interface ZipExtraTimestamps {
  /** Access time. */
  atime?: Date;
  /** Metadata change time (Unix ctime). */
  ctime?: Date;
  /** Creation time (Windows/NTFS "btime"). */
  birthTime?: Date;
}

function clampUint32(value: number): number {
  if (!Number.isFinite(value)) {
    return 0;
  }
  if (value <= 0) {
    return 0;
  }
  // 0xFFFFFFFF fits JS safe integer.
  if (value >= 0xffffffff) {
    return 0xffffffff;
  }
  return value >>> 0;
}

function unixSecondsFromDate(date: Date): number {
  return clampUint32(Math.floor(date.getTime() / 1000));
}

function fileTimeFromDate(date: Date): bigint {
  if (typeof BigInt !== "function") {
    throw new ArchiveError("NTFS timestamps require BigInt support");
  }
  // Windows FILETIME: 100-nanosecond intervals since 1601-01-01 UTC.
  // JS Date is milliseconds since 1970-01-01 UTC.
  const unixMs = BigInt(date.getTime());
  const EPOCH_DIFF_100NS = 116444736000000000n; // 1601->1970 in 100ns
  return unixMs * 10000n + EPOCH_DIFF_100NS;
}

function buildExtendedTimestampExtraField(modTime: Date, extra?: ZipExtraTimestamps): Uint8Array {
  // Data: [flags:1][mtime?:4][atime?:4][ctime?:4]
  const includeAtime = extra?.atime !== undefined;
  const includeCtime = extra?.ctime !== undefined;

  let flags = 0x01;
  if (includeAtime) {
    flags |= 0x02;
  }
  if (includeCtime) {
    flags |= 0x04;
  }

  const payloadSize = 1 + 4 + (includeAtime ? 4 : 0) + (includeCtime ? 4 : 0);
  const out = new Uint8Array(4 + payloadSize);
  const view = new DataView(out.buffer);

  view.setUint16(0, EXTENDED_TIMESTAMP_ID, true);
  view.setUint16(2, payloadSize, true);
  out[4] = flags;
  view.setUint32(5, unixSecondsFromDate(modTime), true);

  let cursor = 9;
  if (includeAtime) {
    view.setUint32(cursor, unixSecondsFromDate(extra!.atime!), true);
    cursor += 4;
  }
  if (includeCtime) {
    view.setUint32(cursor, unixSecondsFromDate(extra!.ctime!), true);
  }

  return out;
}

function buildNtfsTimestampExtraField(modTime: Date, extra?: ZipExtraTimestamps): Uint8Array {
  if (typeof BigInt !== "function") {
    throw new ArchiveError("NTFS timestamps require BigInt support");
  }
  // NTFS extra field (0x000a)
  // Data:
  //   [reserved:4=0]
  //   [tag:2=0x0001][size:2=32]
  //   [mtime:8][atime:8][ctime:8][btime:8] (FILETIME)
  const atime = extra?.atime ?? modTime;
  const ctime = extra?.ctime ?? modTime;
  const btime = extra?.birthTime ?? modTime;

  const dataSize = 4 + 2 + 2 + 32;
  const out = new Uint8Array(4 + dataSize);
  const view = new DataView(out.buffer);

  view.setUint16(0, NTFS_TIMESTAMP_ID, true);
  view.setUint16(2, dataSize, true);

  // reserved
  view.setUint32(4, 0, true);

  // attribute tag 0x0001, size 32
  view.setUint16(8, 0x0001, true);
  view.setUint16(10, 32, true);

  let cursor = 12;
  view.setBigUint64(cursor, fileTimeFromDate(modTime), true);
  cursor += 8;
  view.setBigUint64(cursor, fileTimeFromDate(atime), true);
  cursor += 8;
  view.setBigUint64(cursor, fileTimeFromDate(ctime), true);
  cursor += 8;
  view.setBigUint64(cursor, fileTimeFromDate(btime), true);

  return out;
}

/**
 * DOS date/time helpers for ZIP files.
 *
 * The DOS fields are wall-clock time with no zone: every common writer (Info-ZIP, 7-Zip,
 * Windows, Python's `zipfile`) fills them from local time and reads them back as local
 * time, and so does this module. They hold 1980-01-01 00:00:00 to 2107-12-31 23:59:58 in
 * two-second steps.
 */

/** The first and last moments the DOS fields can hold, as local time. */
const DOS_MIN_YEAR = 1980;
const DOS_MAX_YEAR = 2107;

/** Seconds since 1970 the Info-ZIP UT field (an unsigned 32-bit count) can hold. */
const UT_MAX_SECONDS = 0xffffffff;

function inDosRange(date: Date): boolean {
  const year = date.getFullYear();
  return year >= DOS_MIN_YEAR && year <= DOS_MAX_YEAR;
}

function inUtRange(date: Date): boolean {
  const seconds = Math.floor(date.getTime() / 1000);
  return seconds >= 0 && seconds <= UT_MAX_SECONDS;
}

/**
 * Whether the DOS fields read back as `date`, to their two seconds: false outside their
 * range, and for a wall-clock time the zone repeats (the hour a DST change goes back), which
 * reads back as its first occurrence.
 */
function dosRoundTrips(date: Date): boolean {
  if (!inDosRange(date)) {
    return false;
  }
  const [dosTime, dosDate] = dateToDos(date);
  return parseDosDateTime(dosDate, dosTime).getTime() === Math.floor(date.getTime() / 2000) * 2000;
}

/**
 * Whether an entry modified at `date` carries a UT field: always when `mode` asks for one,
 * and with `"dos"` whenever the DOS fields alone would not read back as the date — out of
 * their range, or in a repeated DST hour — since the UT field then stores it exactly. Never
 * when the UT field cannot hold the date: the DOS fields alone are then the best there is,
 * and a clamped UT field would win on reading.
 */
function writesUtField(date: Date, mode: ZipTimestampMode): boolean {
  return inUtRange(date) && (mode !== "dos" || !dosRoundTrips(date));
}

/**
 * Why `date` cannot be stored exactly (to the DOS fields' two seconds) as an entry's
 * modification time, or `undefined` when it can: the DOS fields hold 1980–2107 local time
 * and the UT field 1970-01-01T00:00:00Z to 2106-02-07T06:28:15Z, and a date either holds is
 * written so that it reads back. `mode` does not narrow this — `"dos"` adds a UT field to an
 * entry whose DOS fields would not read back as its date (see `writesUtField`).
 */
export function zipModTimeProblem(date: Date): string | undefined {
  return inDosRange(date) || inUtRange(date)
    ? undefined
    : `${date.toISOString()} is outside what a ZIP can store ` +
        `(1970-01-01T00:00:00Z to ${DOS_MAX_YEAR}-12-31 local time)`;
}

/**
 * `date` if a ZIP can store it, else the nearest date it can. For a stamp the caller did not
 * choose — a document's own created/modified date — where refusing to write would be worse
 * than moving the stamp.
 */
export function fitZipModTime(date: Date | undefined): Date | undefined {
  if (date === undefined || Number.isNaN(date.getTime()) || zipModTimeProblem(date) === undefined) {
    return date;
  }
  return date.getTime() < 0 ? new Date(0) : new Date(DOS_MAX_YEAR, 11, 31, 23, 59, 58);
}

/**
 * Convert a Date to DOS time/date fields from its local time fields. A date outside the
 * DOS range is clamped to the nearest end (it is never wrapped); a caller-given `modTime`
 * is checked with {@link zipModTimeProblem} before it gets here, so only times taken from
 * files on disk are clamped — as Info-ZIP clamps them.
 */
export function dateToDos(date: Date): [number, number] {
  const year = date.getFullYear();
  const clamped =
    year < DOS_MIN_YEAR
      ? new Date(DOS_MIN_YEAR, 0, 1, 0, 0, 0)
      : year > DOS_MAX_YEAR
        ? new Date(DOS_MAX_YEAR, 11, 31, 23, 59, 58)
        : date;

  const dosTime =
    (clamped.getHours() << 11) | (clamped.getMinutes() << 5) | (clamped.getSeconds() >> 1);

  const dosDate =
    ((clamped.getFullYear() - DOS_MIN_YEAR) << 9) |
    ((clamped.getMonth() + 1) << 5) |
    clamped.getDate();

  return [dosTime, dosDate];
}

/**
 * Parse DOS date/time fields to a Date, as local time — the inverse of {@link dateToDos}.
 */
export function parseDosDateTime(date: number, time?: number): Date {
  const day = date & 0x1f;
  const month = (date >> 5) & 0x0f;
  const year = ((date >> 9) & 0x7f) + DOS_MIN_YEAR;
  const seconds = time ? (time & 0x1f) * 2 : 0;
  const minutes = time ? (time >> 5) & 0x3f : 0;
  const hours = time ? time >> 11 : 0;

  return new Date(year, month - 1, day, hours, minutes, seconds);
}

/**
 * How to write timestamps in ZIP headers.
 *
 * ZIP always has DOS date/time fields; `dos+utc` additionally writes the Info-ZIP
 * extended timestamp extra field (0x5455) for a UTC mtime.
 */
/**
 * - "dos": DOS date/time only
 * - "dos+utc": also writes Info-ZIP extended timestamp (0x5455) for UTC mtime (and optional atime/ctime)
 * - "dos+utc+ntfs": additionally writes NTFS timestamps (0x000a) including creation time
 */
export type ZipTimestampMode = "dos" | "dos+utc" | "dos+utc+ntfs";

export function resolveZipLastModifiedDateFromUnixSeconds(
  dosDate: number,
  dosTime: number,
  mtimeUnixSeconds?: number
): Date {
  if (mtimeUnixSeconds === undefined) {
    return parseDosDateTime(dosDate, dosTime);
  }
  return new Date(mtimeUnixSeconds * 1000);
}

export function buildZipTimestampExtraField(
  modTime: Date,
  mode: ZipTimestampMode,
  extra?: ZipExtraTimestamps
): Uint8Array {
  const parts: Uint8Array[] = [];
  if (writesUtField(modTime, mode)) {
    parts.push(buildExtendedTimestampExtraField(modTime, extra));
  }
  if (mode === "dos+utc+ntfs") {
    parts.push(buildNtfsTimestampExtraField(modTime, extra));
  }
  return concatUint8Arrays(parts);
}

export function dateToZipDos(modTime: Date): { dosTime: number; dosDate: number } {
  const [dosTime, dosDate] = dateToDos(modTime);
  return { dosTime, dosDate };
}
