/**
 * Reading an XLSB worksheet **one record at a time**.
 *
 * `Stream.WorkbookReader` walks the ZIP forward and hands each entry to a decoder. For XLSX that decoder is a SAX
 * parser producing rows; this is the same thing for a BIFF12 stream. The reader above is unchanged: the entry
 * iteration, the `waiting-worksheet` spooling that copes with shared strings arriving *after* a sheet, the
 * worksheet-reader surface a caller consumes — all of it is shared, and only the bytes-to-rows step differs.
 *
 * ## What is bounded
 *
 * Rows. A record is framed by its own header, so a sheet is decoded without holding the part: bytes are consumed as
 * they arrive, a row is emitted when its cells are complete, and nothing is retained. A million-row sheet costs one
 * row at a time.
 *
 * ## What is not, and why saying so matters
 *
 * The **shared-string table** and the **style table**. Both are indexed by the cell records — `BrtCellIsst` holds an
 * index, not a string — so they have to exist in full before a cell can be resolved. The XLSX streaming reader is in
 * the same position and deals with it the same way: if a sheet arrives first, the reader spools that entry and comes
 * back to it. That is a real cost on a package whose ZIP puts sheets before `sharedStrings.bin` — which is what this
 * library's own writer does — and it is honest to name it rather than describe the reader as unconditionally
 * streaming.
 *
 * ## Partial-record handling
 *
 * A record can straddle a chunk boundary, and this is the part a from-scratch attempt gets wrong. The decoder keeps a
 * carry buffer: bytes that do not yet form a complete `id` + `length` + payload stay there until the next chunk
 * arrives. Records are never split by the framing itself — a truncated tail at end of stream is a corrupt part and is
 * reported as one, not silently dropped.
 */
import { XlsbParseError } from "@excel/errors";
import { recordSpec } from "@excel/xlsb/spec/records";

/** One decoded record: its id, its name when known, and its payload. */
export interface StreamedRecord {
  readonly id: number;
  readonly name: string | undefined;
  readonly payload: Uint8Array;
}

/**
 * Decode a chunked byte stream into BIFF12 records.
 *
 * Written as an async generator so the caller drives it and back-pressure is the caller's: nothing is read ahead of
 * what is consumed. The carry buffer grows only to the size of one record.
 */
export async function* streamBiffRecords(
  chunks: AsyncIterable<Uint8Array>,
  part: string
): AsyncIterableIterator<StreamedRecord> {
  for await (const batch of streamBiffRecordBatches(chunks, part)) {
    yield* batch;
  }
}

/**
 * {@link streamBiffRecords}, one array per input chunk instead of one step per record.
 *
 * A worksheet is millions of records — usually one per cell — and an async generator step is a promise round trip,
 * which is what made those steps the largest single cost of a streamed XLSB read. A chunk's records are framed in one
 * synchronous pass and handed over together; the caller still drives the stream chunk by chunk.
 *
 * Each record's payload is a view of a buffer this function owns and never writes to again, so a record stays valid
 * after the next chunk arrives, however the source treats its own buffers.
 */
export async function* streamBiffRecordBatches(
  chunks: AsyncIterable<Uint8Array>,
  part: string
): AsyncIterableIterator<StreamedRecord[]> {
  let carry: Uint8Array = EMPTY;
  for await (const chunk of chunks) {
    if (chunk.length === 0) {
      continue;
    }
    // Copied, because a source may reuse its chunk buffer and a record handed out must stay valid; joined to the
    // unfinished tail in the same copy when there is one.
    const buffer = concat(carry, chunk);
    const records: StreamedRecord[] = [];
    let consumed = 0;
    for (;;) {
      const framed = frame(buffer, consumed, part);
      if (framed === undefined) {
        break;
      }
      consumed = framed.next;
      records.push(framed.record);
    }
    // Only the unconsumed tail is kept, copied, so a sheet does not pin every buffer it was delivered in.
    carry = consumed === buffer.length ? EMPTY : buffer.slice(consumed);
    if (records.length > 0) {
      yield records;
    }
  }
  if (carry.length > 0) {
    throw new XlsbParseError(part, `${carry.length} trailing byte(s) do not form a record`);
  }
}

const EMPTY = new Uint8Array(0);

/**
 * The record starting at `offset`, or `undefined` when the buffer does not hold all of it yet.
 *
 * The header is two variable-length integers, so how many bytes it occupies is not known until it has been read —
 * which is why this cannot simply check for a fixed minimum before decoding.
 */
function frame(
  buffer: Uint8Array,
  offset: number,
  part: string
): { readonly record: StreamedRecord; readonly next: number } | undefined {
  if (offset >= buffer.length) {
    return undefined;
  }
  // The two header integers, read without throwing on a short buffer: "not yet" and "malformed" are the same bytes
  // mid-stream, and only the end-of-stream check above can tell them apart.
  const idField = varUInt(buffer, offset, 2);
  if (idField === undefined) {
    return undefined;
  }
  const lengthField = varUInt(buffer, idField.next, 4);
  if (lengthField === undefined) {
    return undefined;
  }
  const id = idField.value;
  const length = lengthField.value;
  const start = lengthField.next;
  if (start + length > buffer.length) {
    return undefined;
  }
  return {
    record: {
      id,
      name: recordSpec(id)?.name,
      payload: buffer.subarray(start, start + length)
    },
    next: start + length
  };
}

/**
 * A variable-length unsigned integer, or `undefined` when the buffer stops inside it.
 *
 * The same encoding `xlsb/binary` reads — seven bits per byte, the high bit continuing — but returning absence rather
 * than throwing, because a partial header is the normal state of a stream between chunks. Arithmetic rather than
 * shifts for the same reason the buffered reader uses it: a four-byte length reaches bit 28, and `|` would coerce to
 * int32 and turn a large length negative.
 */
function varUInt(
  buffer: Uint8Array,
  start: number,
  maxBytes: number
): { readonly value: number; readonly next: number } | undefined {
  let value = 0;
  let offset = start;
  for (let index = 0; index < maxBytes; index++) {
    if (offset >= buffer.length) {
      return undefined;
    }
    const byte = buffer[offset++]!;
    value += (byte & 0x7f) * 2 ** (7 * index);
    if ((byte & 0x80) === 0) {
      return { value, next: offset };
    }
  }
  return { value, next: offset };
}

/** `left` followed by `right`, in a new buffer. */
function concat(left: Uint8Array, right: Uint8Array): Uint8Array {
  const joined = new Uint8Array(left.length + right.length);
  joined.set(left, 0);
  joined.set(right, left.length);
  return joined;
}
