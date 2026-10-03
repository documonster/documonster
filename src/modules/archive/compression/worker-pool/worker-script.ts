/**
 * Inline Worker Script Generator
 *
 * Generates the worker script code that handles compression/decompression tasks.
 * Uses native CompressionStream/DecompressionStream when available,
 * falls back to a pure JS implementation message protocol.
 */

/**
 * Generate the inline worker script code
 *
 * The worker supports:
 * - deflate: Compress data using deflate-raw
 * - inflate: Decompress data using deflate-raw
 *
 * It uses native Web Streams API when available, which is significantly faster
 * than any pure JS implementation.
 */
export function generateWorkerScript(): string {
  return `
'use strict';

// Check deflate-raw support once at startup
const hasDeflateRaw = (() => {
  try {
    new CompressionStream('deflate-raw');
    new DecompressionStream('deflate-raw');
    return true;
  } catch {
    return false;
  }
})();

const EMPTY_UINT8ARRAY = new Uint8Array(0);

/** Output exceeded the task's maxOutputLength; reported to the pool as \`limit\`. */
class LimitError extends Error {
  constructor(limit) {
    super('Decompressed output exceeds maxOutputLength (' + limit + ' bytes)');
    this.limit = limit;
  }
}

/**
 * Process data through a TransformStream (compress or decompress)
 */
async function processWithStream(stream, data, maxOutputLength) {
  const writer = stream.writable.getWriter();
  const reader = stream.readable.getReader();

  // Start reading output
  let firstChunk = null;
  let chunks = null;
  let produced = 0;
  const readPromise = (async () => {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      if (!value) {
        continue;
      }
      produced += value.length;
      if (maxOutputLength !== undefined && produced > maxOutputLength) {
        // Stop pulling so nothing past the bound is produced or buffered.
        const limitError = new LimitError(maxOutputLength);
        reader.cancel(limitError).catch(() => {});
        throw limitError;
      }
      if (firstChunk === null) {
        firstChunk = value;
      } else if (chunks === null) {
        chunks = [firstChunk, value];
      } else {
        chunks.push(value);
      }
    }
  })();

  // Write input and close. Once the reader has failed the writes reject
  // too; the read's verdict is the one to report.
  const read = readPromise.then(() => null, err => err);
  try {
    await writer.write(data);
    await writer.close();
  } catch (err) {
    throw (await read) ?? err;
  }

  // Wait for all output
  const readError = await read;
  if (readError) throw readError;

  // Fast path for common cases
  if (firstChunk === null) return EMPTY_UINT8ARRAY;
  if (chunks === null) return firstChunk;

  // Pre-calculate total length and allocate once
  let totalLen = 0;
  for (let i = 0; i < chunks.length; i++) {
    totalLen += chunks[i].length;
  }
  const result = new Uint8Array(totalLen);
  let offset = 0;
  for (let i = 0; i < chunks.length; i++) {
    const chunk = chunks[i];
    result.set(chunk, offset);
    offset += chunk.length;
  }
  return result;
}

/**
 * Streaming sessions keyed by taskId.
 * Each session keeps a transform stream open and produces output chunks as they are available.
 */
const streamingSessions = new Map();

function closeSession(taskId) {
  const session = streamingSessions.get(taskId);
  if (!session) return;
  streamingSessions.delete(taskId);
  try {
    if (session.reader) {
      session.reader.cancel();
    }
  } catch {
    // ignore
  }
  try {
    if (session.writer) {
      session.writer.abort();
    }
  } catch {
    // ignore
  }
}

// Signal ready
self.postMessage({ type: 'ready' });

function postError(taskId, err, startTime) {
  const error = typeof err === 'string' ? err : err?.message || String(err);
  const duration = performance.now() - startTime;
  self.postMessage({
    type: 'error',
    taskId,
    error,
    limit: err instanceof LimitError ? err.limit : undefined,
    duration
  });
}

function postResult(taskId, data, startTime) {
  const duration = performance.now() - startTime;
  self.postMessage(
    { type: 'result', taskId, data, duration },
    [data.buffer]
  );
}

/**
 * Handle incoming messages
 */
self.onmessage = async function(event) {
  const msg = event.data;

  if (!msg || typeof msg.type !== 'string') {
    return;
  }

  // Handle termination request
  if (msg.type === 'terminate') {
    self.close();
    return;
  }

  // Streaming: start
  if (msg.type === 'start') {
    const { taskId, taskType, level } = msg;
    const startTime = performance.now();
    try {
      if (!hasDeflateRaw) {
        throw new Error('deflate-raw not supported in this worker');
      }
      if (typeof taskId !== 'number') {
        throw new Error('Invalid taskId');
      }
      if (streamingSessions.has(taskId)) {
        throw new Error('Streaming task already started');
      }

      const stream = taskType === 'deflate'
        ? new CompressionStream('deflate-raw')
        : taskType === 'inflate'
          ? new DecompressionStream('deflate-raw')
          : null;
      if (!stream) {
        throw new Error('Unknown task type: ' + taskType);
      }

      // Note: CompressionStream does not expose compression level.
      // We keep level for API shape parity (ignored in worker).
      void level;

      const writer = stream.writable.getWriter();
      const reader = stream.readable.getReader();

      const session = {
        writer,
        reader,
        startTime,
        closed: false,
        startedMessage: { type: 'started', taskId },
        outMessage: { type: 'out', taskId, data: null },
        ackMessage: { type: 'ack', taskId },
        doneMessage: { type: 'done', taskId, duration: 0 },
        readLoop: null
      };

      session.readLoop = (async () => {
        try {
          while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            if (value) {
              session.outMessage.data = value;
              self.postMessage(session.outMessage, [value.buffer]);
            }
          }
        } catch (err) {
          // Reader failures will be handled by chunk/end paths too.
        }
      })();

      streamingSessions.set(taskId, session);
      self.postMessage(session.startedMessage);
      return;
    } catch (err) {
      postError(taskId, err, startTime);
      closeSession(taskId);
      return;
    }
  }

  // Streaming: chunk
  if (msg.type === 'chunk') {
    const { taskId, data } = msg;
    const session = streamingSessions.get(taskId);
    if (!session || session.closed) {
      // If the session is gone, ignore.
      return;
    }
    try {
      await session.writer.write(data);
      self.postMessage(session.ackMessage);
    } catch (err) {
      postError(taskId, err, session.startTime);
      session.closed = true;
      closeSession(taskId);
    }
    return;
  }

  // Streaming: end
  if (msg.type === 'end') {
    const { taskId } = msg;
    const session = streamingSessions.get(taskId);
    if (!session || session.closed) {
      return;
    }
    session.closed = true;
    try {
      await session.writer.close();
      if (session.readLoop) {
        await session.readLoop;
      }
      session.doneMessage.duration = performance.now() - session.startTime;
      self.postMessage(session.doneMessage);
    } catch (err) {
      postError(taskId, err, session.startTime);
    } finally {
      closeSession(taskId);
    }
    return;
  }

  // Streaming: abort
  if (msg.type === 'abort') {
    const { taskId } = msg;
    const session = streamingSessions.get(taskId);
    if (!session) {
      return;
    }
    session.closed = true;
    closeSession(taskId);
    postError(taskId, msg.error || 'aborted', session.startTime);
    return;
  }

  // Handle task request
  if (msg.type === 'task') {
    const { taskId, taskType, data, maxOutputLength } = msg;
    const startTime = performance.now();

    try {
      if (!hasDeflateRaw) {
        throw new Error('deflate-raw not supported in this worker');
      }

      let stream;
      switch (taskType) {
        case 'deflate':
          stream = new CompressionStream('deflate-raw');
          break;
        case 'inflate':
          stream = new DecompressionStream('deflate-raw');
          break;
        default:
          throw new Error('Unknown task type: ' + taskType);
      }

      const result = await processWithStream(
        stream,
        data,
        taskType === 'inflate' ? maxOutputLength : undefined
      );

      // Transfer the result buffer for zero-copy
      postResult(taskId, result, startTime);
    } catch (err) {
      postError(taskId, err, startTime);
    }
  }
};
`;
}

// Cache the worker URL with reference counting
let _workerBlobUrl: string | null = null;
let _workerBlobUrlRefCount = 0;

/**
 * Get or create the worker blob URL (increments reference count)
 */
export function getWorkerBlobUrl(): string {
  if (_workerBlobUrl === null) {
    const script = generateWorkerScript();
    const blob = new Blob([script], { type: "text/javascript" });
    _workerBlobUrl = URL.createObjectURL(blob);
  }
  _workerBlobUrlRefCount++;
  return _workerBlobUrl;
}

/**
 * Release the cached worker blob URL (decrements reference count)
 */
export function releaseWorkerBlobUrl(): void {
  if (_workerBlobUrl !== null && --_workerBlobUrlRefCount <= 0) {
    try {
      URL.revokeObjectURL(_workerBlobUrl);
    } catch {
      // Ignore errors during cleanup
    }
    _workerBlobUrl = null;
    _workerBlobUrlRefCount = 0;
  }
}
