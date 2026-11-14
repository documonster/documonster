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
import { readRecordHeader, truncatedRecordError } from "@excel/xlsb/binary";
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
 * after the next chunk arrives, however the source treats its own buffers. A record longer than the chunks it arrives
 * in is gathered into a buffer doubled as its bytes arrive, up to the size its header gives: joining each chunk to
 * what came before copied the record so far every time — 2.5 s for one 32 MB record arriving in 64 KiB chunks — and
 * allocating the size its header gives at once let five bytes of a damaged part claim 256 MB.
 *
 * Malformed bytes are refused as `iterateBiffRecords` refuses them, with the same message, after the records before
 * them have been handed out.
 */
export async function* streamBiffRecordBatches(
  chunks: AsyncIterable<Uint8Array>,
  part: string
): AsyncIterableIterator<StreamedRecord[]> {
  /** The unframed tail, too short to hold a record's header. */
  let carry: Uint8Array = EMPTY;
  /**
   * A record whose header has been read, being gathered: `filled` bytes of its `size` so far, in a buffer grown as they
   * arrive — a header alone is five bytes, and cannot be what decides a 256 MB allocation.
   */
  let pending: Uint8Array | undefined;
  let filled = 0;
  let size = 0;
  /** Where `carry` or `pending` begins in the part, for the byte an error names. */
  let base = 0;
  for await (const chunk of chunks) {
    if (chunk.length === 0) {
      continue;
    }
    const records: StreamedRecord[] = [];
    let rest = chunk;
    if (pending !== undefined) {
      const taken = Math.min(size - filled, chunk.length);
      pending = room(pending, filled + taken, size);
      pending.set(chunk.subarray(0, taken), filled);
      filled += taken;
      if (filled < size) {
        continue;
      }
      records.push(frame(pending, 0, part, base)!.record);
      base += size;
      pending = undefined;
      rest = chunk.subarray(taken);
    }
    // Copied, because a source may reuse its chunk buffer and a record handed out must stay valid; joined to the
    // unframed tail in the same copy.
    const buffer = concat(carry, rest);
    let consumed = 0;
    let failure: unknown;
    try {
      for (
        let framed = frame(buffer, 0, part, base);
        framed;
        framed = frame(buffer, consumed, part, base)
      ) {
        consumed = framed.next;
        records.push(framed.record);
      }
      const header = readRecordHeader(buffer, consumed, part, base);
      if (header === undefined) {
        // Only the unframed tail is kept, copied, so a sheet does not pin every buffer it was delivered in.
        carry = consumed === buffer.length ? EMPTY : buffer.slice(consumed);
      } else {
        size = header.start - consumed + header.length;
        filled = buffer.length - consumed;
        pending = room(EMPTY, filled, size);
        pending.set(buffer.subarray(consumed));
        carry = EMPTY;
      }
      base += consumed;
    } catch (error) {
      failure = error;
    }
    if (records.length > 0) {
      yield records;
    }
    if (failure !== undefined) {
      throw failure;
    }
  }
  const tail = pending === undefined ? carry : pending.subarray(0, filled);
  if (tail.length > 0) {
    throw truncatedRecordError(tail, 0, part, base);
  }
}

const EMPTY = new Uint8Array(0);

/**
 * `buffer`, or a larger copy of its contents, holding at least `needed` bytes of a record of `size`: doubled as the
 * record arrives, so gathering it copies it about twice in all, and never larger than the record.
 */
function room(buffer: Uint8Array, needed: number, size: number): Uint8Array {
  if (needed <= buffer.length) {
    return buffer;
  }
  const grown = new Uint8Array(Math.min(size, Math.max(needed, buffer.length * 2, 64 * 1024)));
  grown.set(buffer);
  return grown;
}

/** The record starting at `offset`, or `undefined` when the buffer does not hold all of it yet. */
function frame(
  buffer: Uint8Array,
  offset: number,
  part: string,
  base: number
): { readonly record: StreamedRecord; readonly next: number } | undefined {
  const header = offset < buffer.length ? readRecordHeader(buffer, offset, part, base) : undefined;
  if (header === undefined || header.start + header.length > buffer.length) {
    return undefined;
  }
  const end = header.start + header.length;
  return {
    record: {
      id: header.id,
      name: recordSpec(header.id)?.name,
      payload: buffer.subarray(header.start, end)
    },
    next: end
  };
}

/** `left`, then `right`, in a new buffer. */
function concat(left: Uint8Array, right: Uint8Array): Uint8Array {
  const joined = new Uint8Array(left.length + right.length);
  joined.set(left, 0);
  joined.set(right, left.length);
  return joined;
}
