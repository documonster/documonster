# Archive Module

[中文](README_zh.md)

A zero-dependency, cross-platform archive toolkit for creating, reading, and editing ZIP and TAR archives.

```typescript
import { Archive } from "documonster/archive";
```

## Features

- **Zero Dependencies** - Pure TypeScript, no native addons
- **Cross-Platform** - Same API in Node.js and browsers
- **ZIP + TAR** - Unified API for both archive formats
- **Streaming** - True async streaming with backpressure support
- **ZIP Editing** - Filesystem-like editing of existing archives with efficient passthrough
- **Remote Reading** - Read ZIP files via HTTP Range requests (download only what you need)
- **Encryption** - ZIP Traditional and AES-256 encryption/decryption
- **ZIP64** - Large file support (> 4GB files, > 65535 entries)
- **Compression** - DEFLATE, GZIP, and Zlib with sync/async/streaming APIs
- **PNG Encoding** - `Archive.encodePng`, because a PNG is a DEFLATE stream with CRC-32-checked chunks
- **Progress & Abort** - Built-in progress callbacks and AbortSignal support
- **File System Integration** - Node.js convenience layer for disk I/O

---

## Quick Start

### Creating a ZIP

```typescript
import { Archive } from "documonster/archive";

// Convenience function
const archive = Archive.createZip();
archive.add("hello.txt", "Hello, World!");
archive.add("data.json", JSON.stringify({ key: "value" }));
const bytes = await archive.toBuffer();

// With options
const archive = Archive.createZip({ level: 9, comment: "My archive" });
archive.add("file.txt", content, { modTime: new Date() });
archive.addDirectory("empty-dir/");
archive.addSymlink("link.txt", "hello.txt");
const bytes = await archive.toBuffer();
```

### Reading a ZIP

```typescript
import { Archive } from "documonster/archive";

const reader = Archive.openZip(zipBytes);
for await (const entry of reader.entries()) {
  console.log(entry.path, entry.type);
  if (entry.type === "file") {
    const data = await entry.bytes();
    console.log(new TextDecoder().decode(data));
  }
}

// Random access by path
const data = await reader.bytes("hello.txt");
```

### Editing a ZIP

```typescript
import { Archive } from "documonster/archive";

const editor = await Archive.ZipEditor.open(existingZipBytes, { preserve: "best-effort" });
editor.delete("old.txt");
editor.rename("a.txt", "renamed.txt");
editor.set("new.txt", "new content");
const output = await editor.toBuffer();
```

### One call, no builder

When you have every entry up front, one function builds the archive and one reads it
back into a `Map` of path → `{ info, data }`:

```typescript
import { Archive } from "documonster/archive";

const zip = await Archive.packZip(
  [
    ["hello.txt", "Hello"],
    ["data.bin", bytes]
  ],
  { level: 9 }
);
const files = await Archive.unpackZip(zip);
files.get("hello.txt")?.data; // Uint8Array

// Synchronous forms for in-memory sources
const same = Archive.unpackZipSync(Archive.packZipSync(new Map([["a.txt", "A"]])));
```

### Which function for which job

The same shape for every format: a builder and a reader when you need options per entry,
progress, streaming output or lazy reads; one-shot functions when you do not.

| Job                               | ZIP                        | TAR                         | tar.gz                          |
| --------------------------------- | -------------------------- | --------------------------- | ------------------------------- |
| Builder                           | `Archive.createZip()`      | `Archive.createTar()`       | `Archive.createTarGz()`         |
| Reader                            | `Archive.openZip(src)`     | `Archive.openTar(src)`      | `Archive.openTarGz(src)`        |
| Entries → bytes                   | `Archive.packZip(entries)` | `Archive.packTar(entries)`  | `Archive.packTarGz(entries)`    |
| Entries → bytes, sync             | `Archive.packZipSync(…)`   | `Archive.packTarSync(…)`    | `Archive.packTarGzSync(…)`      |
| Bytes → `Map<path, {info, data}>` | `Archive.unpackZip(src)`   | `Archive.unpackTar(src)`    | `Archive.unpackTarGz(src)`      |
| Bytes → map, sync                 | `Archive.unpackZipSync(…)` | `Archive.unpackTarSync(…)`  | `Archive.unpackTarGzSync(…)`    |
| Bytes → entry list                | — (use the reader)         | `Archive.parseTar(bytes)`   | `Archive.parseTarGz(src)`       |
| Bytes → entry stream              | — (use the reader)         | `Archive.parseTarStream(…)` | `Archive.parseTarGzStream(src)` |

One verb means one thing and the format is in the name: `create*` returns a builder,
`open*` a reader, `pack*` bytes, `unpack*` a path map and `parse*` an entry list.
`entries` is a `Map<string, source>` or an iterable of `[name, source]` pairs or
`{ name, source }` objects. The synchronous forms take in-memory sources only and work on
both Node and the browser. The builder and reader classes (`ZipArchive`, `ZipReader`,
`TarArchive`, `TarReader`, `TarGzArchive`) are public values for `instanceof` checks and
subclassing.

The deprecated `Archive.zip` / `Archive.unzip` (format chosen by `options.format`),
`Archive.tar` / `tarSync` / `untar` and `createTarArchive` / `createTarReader` keep working
unchanged until 1.0.

---

## ZIP API

### `Archive.createZip(options?)`

Create a `ZipArchive` builder. `Archive.createTar()` and `Archive.createTarGz()` are the TAR and tar.gz builders.

```typescript
import { Archive } from "documonster/archive";

const archive = Archive.createZip(); // ZipArchive
const tarArchive = Archive.createTar(); // TarArchive
```

### `ZipArchive`

Streaming ZIP archive builder with chainable API.

```typescript
import { Archive } from "documonster/archive";

const archive = Archive.createZip({ level: 6, reproducible: true });

// Add entries (chainable)
archive
  .add("file.txt", "text content")
  .add("binary.dat", uint8Array)
  .add("from-stream.bin", readableStream)
  .addDirectory("empty-dir/")
  .addSymlink("link.txt", "file.txt");

// Output options:

// 1. Complete Uint8Array
const bytes = await archive.toBuffer();

// 2. Synchronous (in-memory sources only)
const bytes = archive.toBufferSync();

// 3. Async streaming
for await (const chunk of archive.stream()) {
  // Process chunk by chunk
}

// 4. Pipe to WritableStream
await archive.pipeTo(writableStream);

// 5. Operation with progress and abort
const op = archive.operation({
  onProgress: p => console.log(`${p.entriesDone}/${p.entriesTotal}`),
  signal: abortController.signal
});
for await (const chunk of op.iterable) {
  /* process chunk */
}
```

**`ZipOptions`:**

| Option         | Type                       | Default   | Description                               |
| -------------- | -------------------------- | --------- | ----------------------------------------- |
| `level`        | `number`                   | `6`       | Compression level (0 = store, 9 = max)    |
| `comment`      | `string`                   | -         | Archive-level comment                     |
| `reproducible` | `boolean`                  | `false`   | Stable timestamps for reproducible builds |
| `smartStore`   | `boolean`                  | `true`    | Auto-STORE incompressible data            |
| `zip64`        | `"auto" \| boolean`        | `"auto"`  | ZIP64 mode                                |
| `signal`       | `AbortSignal`              | -         | Abort signal                              |
| `onProgress`   | `(p: ZipProgress) => void` | -         | Progress callback                         |
| `encoding`     | `ZipStringEncoding`        | `"utf-8"` | String encoding for names                 |
| `timestamps`   | `ZipTimestampMode`         | -         | Timestamp encoding mode                   |
| `path`         | `false \| ZipPathOptions`  | `false`   | Entry-name normalization mode             |

Entry names are checked when they are added. An absolute name (`/etc/passwd`, `C:x`) or one whose
`..` segments climb above the archive root (`../evil.txt`) throws `ArchiveArgumentError` from `add`, as does
a `level` outside 0–9, an invalid `Date` or a `mode` the header cannot hold; `createTar().add`
applies the same name rule. Passing a `path` mode hands the naming to that mode instead:
`{ mode: "safe" }` normalizes and rejects, `{ mode: "posix" }` normalizes and keeps a leading `..`.
An entry source that is not an `ArchiveSource` is refused by `add` too, not when the archive is
written.

Timestamps: the DOS date and time fields every ZIP carries are local wall-clock time, written
from `modTime`'s local fields and read back as local time, as Info-ZIP, 7-Zip and Windows do;
they hold 1980–2107. `timestamps: "dos+utc"` adds the Info-ZIP UT field, which holds the exact
UTC time from 1970 to 2106 and is what readers prefer. Even in the default `"dos"` mode an entry
gets a UT field whenever its DOS fields alone would not read back as its `modTime` — a date before
1980, or a wall-clock time the zone repeats (the hour a DST change goes back) — so every `modTime`
from 1970 to 2106 reads back as written, to the DOS fields' two seconds. One only the DOS fields
can hold (2106-02-07 to 2107) gets none, and a `modTime` neither can hold is refused with an
`ArchiveError`. A time taken from a file on disk is clamped to the DOS range instead.

### `Archive.openZip(source, options?)`

Open a ZIP archive for reading. `Archive.openTar()` and `Archive.openTarGz()` are the TAR and tar.gz readers.

```typescript
import { Archive } from "documonster/archive";

const reader = Archive.openZip(zipBytes);
const reader = Archive.openZip(zipBytes, { password: "secret" });
const reader = Archive.openTar(tarBytes);
```

### `ZipReader`

Streaming ZIP reader with random-access support.

```typescript
import { Archive } from "documonster/archive";

const reader = Archive.openZip(zipBytes);

// Streaming iteration
for await (const entry of reader.entries()) {
  const data = await entry.bytes();
  const text = await entry.text();
  // Or stream the entry
  for await (const chunk of entry.stream()) {
    /* process chunk */
  }
  // Or pipe
  await entry.pipeTo(writableStream);
}

// Random access (a streamed source is buffered first)
const entry = await reader.get("path/to/file.txt");
const bytes = await reader.bytes("path/to/file.txt");
const paths = await reader.list(); // every entry, in central-directory order

// With progress and abort
const op = reader.operation({
  onProgress: p => console.log(`${p.entriesEmitted} entries, ${p.bytesOut} bytes`),
  signal: controller.signal
});
```

### `UnzipEntry`

Represents a single entry in an archive.

```typescript
entry.path; // "path/to/file.txt"
entry.type; // "file" | "directory" | "symlink"
entry.mode; // Unix mode (e.g., 0o644)
entry.linkTarget; // Symlink target (after bytes())
entry.info; // frozen ZipEntryRecord: sizes, crc32, lastModified, comment, …
//            (from the local header while streaming; see its documentation)

await entry.bytes(); // Uint8Array
await entry.text(); // String (UTF-8)
await entry.text("latin1"); // String (custom encoding)
for await (const chunk of entry.stream()) {
  /* process chunk */
}
await entry.pipeTo(writable);
entry.readableStream(); // WHATWG ReadableStream
entry.discard(); // Skip without reading
```

**`UnzipOptions`:**

| Option         | Type                         | Default   | Description                               |
| -------------- | ---------------------------- | --------- | ----------------------------------------- |
| `password`     | `string \| Uint8Array`       | -         | Decryption password                       |
| `encoding`     | `ZipStringEncoding`          | `"utf-8"` | Entry name encoding                       |
| `checkCrc32`   | `boolean`                    | `true`    | Verify each entry's CRC-32                |
| `maxEntrySize` | `number`                     | 512 MiB   | Largest entry `bytes()` reads into memory |
| `maxEntries`   | `number`                     | -         | Most entries the archive may hold         |
| `signal`       | `AbortSignal`                | -         | Abort signal                              |
| `onProgress`   | `(p: UnzipProgress) => void` | -         | Progress callback                         |

**Limits.** Every reader takes the same two limits: `maxEntrySize` (the largest entry read into
memory) and, for ZIP, `maxEntries` — `unpackZip`,
`openZip`, `RemoteZipReader` and `ArchiveFile` alike; the TAR readers take `maxEntrySize` too
(100 MiB by default; `maxFileSize` is its deprecated name). Set them when reading archives you
did not write. Every reader checks them when called: a limit that is not a whole number from 0 (or
`Infinity`) is an `ArchiveArgumentError` naming `options.<limit>`. A limit that is hit is an
`ArchiveLimitExceededError` whose `limit` names the option and `allowed` its value; for `maxEntrySize` it is
the subclass `ArchiveFileTooLargeError`, which also carries the entry's `path`.

**Strings.** A string given to a reader is the archive's content, encoded as UTF-8 — never a file
path or URL. Read a file's bytes first, or open it with `ArchiveFile.fromFile` in Node; a ZIP's
binary content does not survive a JavaScript string.

**Names.** A ZIP holds one entry per name: `ZipArchive.add` and `packZip` refuse a name already
added (`ArchiveArgumentError`). An archive written elsewhere with a name twice is read the way
extractors read it — the later entry is the one `get`, `bytes`, `unpackZip` and
`RemoteZipReader.extract` return, as `TarReader.get` and `unpackTar` do for TAR — and an
unedited `ZipEditor.open` writes every entry back.

**One-shot sources.** A `ReadableStream`, a Node stream or an async iterable can be read once.
`openZip` / `openTar` over one read it either with one `entries()` pass, or by buffering it the
first time `list`, `get` or `bytes` needs it (every later call, `entries()` included, then uses
that); anything else is an `ArchiveError` saying the stream was consumed. `openTarGz` holds
the decompressed archive and can be read any number of times.

**Integrity.** Every ZIP read verifies each entry against the CRC-32 the archive records for
it and throws `ZipCrc32MismatchError` on a mismatch — `Archive.openZip` / `ZipReader` (buffered
and streaming sources, `bytes()` and `stream()`), `Archive.unpackZip` / `unpackZipSync`,
`RemoteZipReader` and `ArchiveFile`. `stream()` reports a mismatch after the entry's
last chunk, since the checksum covers the whole entry. Pass `checkCrc32: false` to read a
damaged archive as far as it goes. AES entries are authenticated by their HMAC instead, and
ZipCrypto entries are always verified.

The check covers **the entries you read**, from every source alike: a mismatch fails that
`bytes()` or `stream()` — catch it and the iteration goes on to the next entry — and an entry
you skip (move past, or `discard()`) is not verified and fails nothing. Over a streamed
source the reader checks an entry as its bytes go by, but reports the result to whoever reads
it, exactly as an archive in memory checks an entry only when it is extracted.

**Streamed sources.** A stream can only move to its next entry once the current one has been
read, so the reader handles the entries you do not read:

- An entry you move past without reading is kept in memory while it is small (up to
  256 KiB), so it can still be read afterwards; a larger one is skipped, and reading it later
  throws an `ArchiveError` rather than stalling. A `discard()`ed entry throws when read too.
- Leaving the loop — `break`, `return`, or a `throw` from its body — settles at once. The
  entry you left with stays readable on the same terms — so read a large one before leaving —
  and the source is released as soon
  as that entry is through: a Node stream is destroyed and a Web `ReadableStream` cancelled
  and unlocked — the same as Excel's streaming workbook reader — and the parser and its
  inflater are freed. A failure or an abort releases everything at once.

---

## ZIP Editor

Filesystem-like editing of existing ZIP archives. Unchanged entries are passed through efficiently (raw compressed bytes preserved).

```typescript
import { Archive } from "documonster/archive";

// From bytes
const editor = await Archive.ZipEditor.open(zipBytes, {
  reproducible: true,
  preserve: "best-effort",
  onWarning: w => console.warn(w.code, w.entry, w.message)
});

// From URL (HTTP Range requests)
const remote = await Archive.ZipEditor.openUrl("https://example.com/archive.zip");

// Operations
editor.has("file.txt"); // Check existence
editor.set("new.txt", "content"); // Add or replace
editor.delete("old.txt"); // Delete entry
editor.deleteDirectory("old-dir/"); // Delete directory recursively
editor.rename("a.txt", "b.txt"); // Rename entry
editor.setComment("Updated archive"); // Set archive comment

// Reusable edit plans
const plan = new Archive.ZipEditPlan();
plan.set("config.json", newConfig);
plan.delete("temp/");
editor.apply(plan);

// Output
const output = await editor.toBuffer();
await editor.pipeTo(writable);
for await (const chunk of editor.stream()) {
  /* process chunk */
}
```

---

## Remote ZIP Reader

Read ZIP files from HTTP servers using Range requests -- download only the entries you need.

```typescript
import { Archive } from "documonster/archive";

const reader = await Archive.RemoteZipReader.open("https://example.com/large.zip");

// Metadata (no content downloaded yet)
const entries = reader.getEntries();
console.log(`${entries.length} entries`);

const files = reader.listFiles(); // file entries (no directories); one per entry, like getFileCount()
const count = reader.getFileCount();

// Find entries by pattern
const jsFiles = reader.findEntries("**/*.js");

// Extract single file (downloads only that entry)
const data = await reader.extract("path/to/file.txt");

// Extract multiple (batched for efficiency)
const results = await reader.extractMultiple(["a.txt", "b.txt"]);

// Streaming extraction
await reader.extractToStream("large-file.bin", writableStream);

// Password-protected archives
const reader = await Archive.RemoteZipReader.open(url);
const isValid = await reader.checkPassword("secret.txt", "mypassword");
const data = await reader.extract("secret.txt", { password: "mypassword" });

// Statistics
const stats = reader.getStats();
console.log(`Downloaded ${stats.http?.bytesDownloaded} bytes (${stats.http?.downloadedPercent}%)`);

await reader.close();

// Or let `await using` close it on every way out of the block — no try/finally needed.
// `ZipReader`, `RemoteZipReader`, `HttpRangeReader`, `TarReader` and `ZipEditor` support it; the
// member is typed where your `lib` declares `Symbol.asyncDispose` (`ESNext`, or `@types/node`).
// Where the runtime has no `Symbol.asyncDispose`, a polyfill must be loaded before this package,
// or define it as `Symbol.for("Symbol.asyncDispose")` — what esbuild and Babel look up in its absence.
{
  await using remote = await Archive.RemoteZipReader.open(url);
  await remote.extract("path/to/file.txt");
}
```

---

## Compression API

Low-level compression utilities. DEFLATE, GZIP, and Zlib with sync, async, and streaming variants.

### One-Shot Compression

```typescript
import { Compression } from "documonster/archive";

// DEFLATE-RAW (used by ZIP files)
const compressed = await Compression.deflate(data, { level: 9 });
const original = await Compression.inflate(compressed);
const compressedSync = Compression.deflateSync(data);

// GZIP
const gzipped = await Compression.gzip(data);
const ungzipped = await Compression.gunzip(gzipped);

// Zlib (RFC 1950)
const zlibbed = await Compression.zlib(data);
const unzlibbed = await Compression.unzlib(zlibbed);

// Auto-detect and decompress
const format = Compression.detectFormat(data); // "gzip" | "zlib" | "deflate-raw"
const original = await Compression.decompressAuto(data);
```

### Streaming Compression

```typescript
import { Compression } from "documonster/archive";

// DEFLATE-RAW streaming
const deflater = Compression.createDeflateStream({ level: 6 });
const inflater = Compression.createInflateStream();

// GZIP streaming
const gzipper = Compression.createGzipStream();
const gunzipper = Compression.createGunzipStream();

// Zlib streaming
const zlibber = Compression.createZlibStream();
const unzlibber = Compression.createUnzlibStream();

// Check availability
Compression.hasDeflateRaw(); // true in Node.js, depends on CompressionStream in browser
```

### CRC32

```typescript
import { Archive } from "documonster/archive";

// One-shot
const checksum = Archive.crc32(data);

// Incremental (streaming)
let state = 0xffffffff;
state = Archive.crc32Update(state, chunk1);
state = Archive.crc32Update(state, chunk2);
const checksum = Archive.crc32Finalize(state);
```

### PNG encoding

A PNG is a DEFLATE stream plus CRC-32-checked chunks — the two primitives above — so the
encoder lives here rather than in a module that merely happens to need pictures. A display
list does not need it directly — `Draw.toPng` and `Mermaid.toPng` call it for you — so
reach for it when you already have pixels:

```typescript
import { Archive } from "documonster/archive";

// 8-bit RGBA, row-major
const bytes = Archive.encodePng(rgba, width, height, { dpi: 192 });
```

8-bit RGBA in, filter 0, one `IDAT`. `dpi` writes a `pHYs` chunk; omit it and none is
written.

---

## TAR Support

Unified API compatible with the ZIP interface.

### Creating TAR archives

```typescript
import { Archive } from "documonster/archive";

// Convenience function
const tarBytes = await Archive.packTar(
  new Map([
    ["file.txt", "content"],
    ["data.bin", uint8Array]
  ]),
  { modTime: new Date() }
);

// Builder API (same as ZipArchive)
const archive = Archive.createTar();
archive
  .add("file.txt", "content", { mode: 0o755 })
  .addDirectory("dir/")
  .addSymlink("link", "file.txt");

const bytes = await archive.toBuffer();
for await (const chunk of archive.stream()) {
  /* process chunk */
}
```

### Reading TAR archives

```typescript
import { Archive } from "documonster/archive";

// Via unified API
const reader = Archive.openTar(tarBytes);

// Direct
const reader = Archive.openTar(tarBytes);
for await (const entry of reader.entries()) {
  console.log(entry.path, entry.isDirectory);
  const data = await entry.bytes();
}

// Random access
const data = await reader.bytes("file.txt");
const paths = await reader.list();
```

### TAR + GZIP

```typescript
import { Archive } from "documonster/archive";

// Create .tar.gz
const tgzBytes = await Archive.packTarGz(new Map([["file.txt", "content"]]), { level: 9 });

// Builder API with streaming gzip
const archive = Archive.createTarGz({ level: 6 });
archive.add("file.txt", "content");
const bytes = await archive.toBuffer();
for await (const chunk of archive.stream()) {
  /* process chunk */
}

// Parse .tar.gz
const entries = await Archive.parseTarGz(tgzBytes);
for (const entry of entries) {
  console.log(entry.info.path, await entry.bytes());
}

// Extract .tar.gz to Map
const files = await Archive.unpackTarGz(tgzBytes);

// Large .tar.gz: decompress and parse incrementally, one entry in memory at a time
for await (const entry of Archive.parseTarGzStream(response.body!)) {
  console.log(entry.info.path, (await entry.bytes()).length);
}
```

`parseTarGz`, `unpackTarGz`, `unpackTarGzSync` and `Archive.openTarGz`
decompress the whole archive in memory, so memory grows with the uncompressed size.
`Archive.openTarGz` keeps that decompressed TAR until `close()`, which is what lets it be
iterated and searched repeatedly even over a one-shot stream; concurrent calls share the one
decompression, and each call's `signal` cancels only that call — with one exception for a
one-shot stream: if every call waiting on its read aborts before decompression completes
(including after `close()`), the stream is cancelled and later calls throw an `ArchiveError`
(not an `AbortError`), since it cannot be read again. A read that has completed is never
discarded — an abort then rejects only that call. An
in-memory source is never cancelled. A failed read of a one-shot stream is final (later calls
rethrow the same error); an in-memory source is retried.
`parseTarGzStream` (like `parseTarStream` for a plain TAR) holds only one entry at a time —
use it for large archives. Its errors surface mid-iteration, after earlier entries have been
yielded. Every TAR and tar.gz reader refuses an archive cut off inside an entry or a header
with an `ArchiveError` rather than returning the partial entry. Once the TAR ends it
still reads the rest of the gzip stream, so a truncated or damaged gzip tail is reported
too. Every tar.gz reader reports a damaged gzip stream as an `ArchiveError` (the decoder's
own error is its `cause`).

A `.tar.gz` comes out byte-for-byte the same whichever way it is built — `packTarGz`,
`packTarGzSync`, `createTarGz().toBuffer()`, `toBufferSync()` or `stream()` — in Node. In a
browser that holds at `level: 0`; at other levels the native `CompressionStream` that
`toBuffer()` and `stream()` use is a different encoder (and takes no level), so those
outputs decompress to the identical TAR without being the same bytes. The `data` that
`unpackTar` / `unpackTarSync` / `unpackTarGz*` return for an in-memory source is a **view**
into that buffer (`subarray`), not a copy — `slice()` it before mutating either.

---

## File System Integration (Node.js)

High-level `ArchiveFile` class with disk I/O, glob patterns, directory traversal, and extraction.

```typescript
import { Archive } from "documonster/archive";

// Create from scratch
const af = new Archive.ArchiveFile();
af.addFile("./src/index.ts");
af.addText("Hello", "hello.txt");
af.addBuffer(bytes, "data.bin");
af.addDirectory("./src", { recursive: true });
af.addGlob("**/*.ts", { cwd: "./src" });
await af.writeFile("output.zip");

// Open existing
const af = await Archive.ArchiveFile.fromFile("archive.zip");
const entries = await af.getEntries();
const data = await af.readEntry("file.txt");
const text = await af.readAsText("file.txt");

// Extract
await af.extractTo("./output", {
  overwrite: "newer",
  filter: (entryPath, isDirectory) => isDirectory || entryPath.endsWith(".txt")
});

// Edit (ZIP only)
af.set("new.txt", "content");
af.delete("old.txt");
af.rename("a.txt", "b.txt");
await af.writeFile("modified.zip");

// Streaming output
for await (const chunk of af.stream()) {
  /* process chunk */
}
await af.streamToFile("output.zip");

// TAR format
const tar = new Archive.ArchiveFile({ format: "tar" });
tar.addFile("./data.txt");
tar.addSymlink("latest.txt", "data.txt", { mode: 0o777 });
await tar.writeFile("archive.tar");
```

Edits are pending until the archive is written, and every read (`has`, `entryCount`,
`getEntryNames`, `getEntries`, `readEntry`, `extractTo`, …) answers for the archive a write would
produce now: the opened archive's entries with the edits — adds, replacements, deletes, renames,
directories — applied, each described as the write records it. Every write — of an opened ZIP or
TAR alike — writes that archive and then becomes what the next one starts from, so writing twice
gives the same archive. Listing names reads no entry's data; a pending entry is read from memory
or, for `addFile`, from disk when a read needs it. A pending stream can be consumed only once, so
the first read or write that needs it keeps its bytes in memory until the archive is written — a
memory cost of the stream's full size; a synchronous read cannot consume it and throws until an
async read has. A `.tgz` opened
by path is written back gzip-compressed. `writeFile` refuses an existing file unless
`overwrite` says otherwise, for TAR as for ZIP.

The helpers it is built on — `Fs.glob`, `Fs.traverseDirectory`, `Fs.ensureDir`,
`Fs.readFileBytes`, … — are published on the Node-only `Fs` namespace. A wrong argument is an
`ArchiveArgumentError` naming the call and the parameter; what the file system then refuses
(`ENOENT`, `EACCES`, …) is passed on as the platform's error. For file streams use
`createReadStream` / `createWriteStream` from `node:fs`:

```typescript
import { Fs } from "documonster/archive";

for await (const entry of Fs.glob("**/*.json", { cwd: "./config" })) {
  console.log(entry.relativePath);
}
await Fs.ensureDir("./out");
```

---

## Encryption

ZIP Traditional (legacy) and WinZip AES (128, 192 or 256-bit) encryption support.

```typescript
import { Archive } from "documonster/archive";

// Write: every entry is encrypted. AES needs the asynchronous builders (Web Crypto);
// "zipcrypto" is also written by packZipSync / toBufferSync.
const encryptedZip = await Archive.packZip([["secret.txt", "classified"]], {
  encryptionMethod: "aes-256",
  password: "secret"
});

// Encrypted entries (handled automatically during read)
const reader = Archive.openZip(encryptedZip, { password: "secret" });
for await (const entry of reader.entries()) {
  const data = await entry.bytes(); // Automatically decrypted
}

// Remote encrypted ZIP
const remote = await Archive.RemoteZipReader.open(url);
const data = await remote.extract("secret.txt", { password: "mypassword" });

// Check password without full extraction
const valid = await remote.checkPassword("file.txt", "mypassword");
```

---

## Error Classes

```typescript
import {
  ArchiveError,
  ArchiveArgumentError, // a wrong argument or option; `.field` names it
  ZipParseError,
  ZipInvalidSignatureError,
  ZipEocdNotFoundError,
  ZipDecryptionError,
  ZipPasswordRequiredError,
  ArchiveFileTooLargeError,
  ZipUnsupportedCompressionError,
  ZipEntrySizeMismatchError,
  ZipCrc32MismatchError,
  ArchiveRangeNotSupportedError,
  ArchiveHttpRangeError,
  ArchiveLimitExceededError,
  ArchiveWorkerStreamStartError,
  ArchiveWorkerUnsupportedError, // a Compression.Workers call on Node; `.operation` names it
  ArchiveFileError, // ArchiveFile.fromFile / writeFile refused by the file system; `.cause` is its error
  AbortError,
  isArchiveError // Type guard: (err: unknown) => err is ArchiveError
} from "documonster/archive";
```

Every archive error extends `ArchiveError`, so `isArchiveError` catches all of them. A wrong argument
or option is an `ArchiveArgumentError` whose `field` is the path from the parameter name
(`"options.level"`, `"source"`, `"entries[0]"`, …) and whose message starts with the call
(`Archive.packZip: …`);
damage in the archive data stays a plain `ArchiveError` or one of its data subclasses. A cancelled
operation throws the shared `AbortError` instead, which is not an `ArchiveError` (it is the
same class `documonster/stream` exports):

```typescript
import { Archive, isArchiveError, AbortError } from "documonster/archive";

try {
  await Archive.unpackZip(bytes);
} catch (err) {
  if (err instanceof AbortError) {
    // cancelled through an AbortSignal
  } else if (isArchiveError(err)) {
    console.error(err.name, err.message, err.cause);
  }
}
```

---

## API Reference

### High-Level Functions

| Function                                   | Description                    |
| ------------------------------------------ | ------------------------------ |
| `Archive.createZip(options?)`              | Create a ZIP builder           |
| `Archive.createTar(options?)`              | Create a TAR builder           |
| `Archive.createTarGz(options?)`            | Create a tar.gz builder        |
| `Archive.openZip(source, options?)`        | Open a ZIP for reading         |
| `Archive.openTar(source, options?)`        | Open a TAR for reading         |
| `Archive.openTarGz(source, options?)`      | Open a tar.gz for reading      |
| `Archive.ZipEditor.open(source, options?)` | Open a ZIP for editing         |
| `Archive.ZipEditor.openUrl(url, options?)` | Open a remote ZIP for editing  |
| `Archive.packZip(entries, options?)`       | Entries → ZIP bytes (async)    |
| `Archive.packZipSync(entries, options?)`   | Entries → ZIP bytes (sync)     |
| `Archive.packTar(entries, options?)`       | Entries → TAR bytes (async)    |
| `Archive.packTarSync(entries, options?)`   | Entries → TAR bytes (sync)     |
| `Archive.packTarGz(entries, options?)`     | Entries → tar.gz bytes (async) |
| `Archive.packTarGzSync(entries, options?)` | Entries → tar.gz bytes (sync)  |
| `Archive.unpackZip(source, options?)`      | ZIP → path map (async)         |
| `Archive.unpackZipSync(bytes, options?)`   | ZIP → path map (sync)          |
| `Archive.unpackTar(source, options?)`      | TAR → path map (async)         |
| `Archive.unpackTarSync(bytes, options?)`   | TAR → path map (sync)          |
| `Archive.unpackTarGz(source, options?)`    | tar.gz → path map (async)      |
| `Archive.unpackTarGzSync(bytes, options?)` | tar.gz → path map (sync)       |

### Classes

| Class                     | Description                                                                                                                 |
| ------------------------- | --------------------------------------------------------------------------------------------------------------------------- |
| `Archive.ZipArchive`      | Streaming ZIP builder                                                                                                       |
| `Archive.ZipReader`       | Streaming ZIP reader                                                                                                        |
| `Archive.UnzipEntry`      | Single archive entry                                                                                                        |
| `Archive.ZipEditor`       | Edit existing ZIP archives                                                                                                  |
| `Archive.TarArchive`      | Streaming TAR builder (`Archive.createTar()`)                                                                               |
| `Archive.TarReader`       | Streaming TAR reader (`Archive.openTar(source)`)                                                                            |
| `Archive.TarGzArchive`    | TAR + GZIP builder (`Archive.createTarGz()`)                                                                                |
| `Archive.RemoteZipReader` | Random-access ZIP reader (`RemoteZipReader.open(url)`, `RemoteZipReader.fromReader(reader)`)                                |
| `Archive.HttpRangeReader` | HTTP Range random-access reader (`HttpRangeReader.open(url)`)                                                               |
| `Archive.ArchiveFile`     | On-disk archives (Node.js; `new ArchiveFile(options)`, `ArchiveFile.fromFile(path)` / `fromFileSync` / `fromBuffer(bytes)`) |

### Compression

| Function                                                       | Description                    |
| -------------------------------------------------------------- | ------------------------------ |
| `Compression.deflate` / `deflateSync`                          | DEFLATE-RAW compress           |
| `Compression.inflate` / `inflateSync`                          | DEFLATE-RAW decompress         |
| `Compression.gzip` / `gzipSync`                                | GZIP compress                  |
| `Compression.gunzip` / `gunzipSync`                            | GZIP decompress                |
| `Compression.zlib` / `zlibSync`                                | Zlib compress                  |
| `Compression.unzlib` / `unzlibSync`                            | Zlib decompress                |
| `Compression.decompressAuto`                                   | Auto-detect and decompress     |
| `Compression.createDeflateStream`                              | Streaming DEFLATE compressor   |
| `Compression.createInflateStream`                              | Streaming DEFLATE decompressor |
| `Compression.createGzipStream`                                 | Streaming GZIP compressor      |
| `Compression.createGunzipStream`                               | Streaming GZIP decompressor    |
| `Archive.crc32`                                                | CRC32 checksum                 |
| `Archive.encodePng`                                            | RGBA pixels to PNG bytes       |
| `Compression.Workers.deflate` / `inflate` (+ `…Batch`, `Pool`) | Web Worker pool (browser)      |

Low-level ZipCrypto / AES primitives and their constants (`encryptZipCrypto`, `decryptAes`,
`AES_KEY_LENGTH`, …) live under `Archive.Crypto`; ordinary encrypted reading and writing only
needs the `password` option.
