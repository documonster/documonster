# XML Module

[中文](README_zh.md)

A zero-dependency, cross-platform XML toolkit for reading and writing XML. Supports both streaming and buffered modes.

```typescript
import { Xml } from "documonster/xml";
```

## Features

- **Zero Dependencies** — Pure TypeScript, no native addons
- **Cross-Platform** — Same API in Node.js and browsers
- **Dual-Mode Writing** — Buffered (`Xml.Writer`) and streaming (`Xml.StreamWriter`)
- **Dual-Mode Reading** — SAX streaming (`Xml.SaxParser`) and DOM tree (`Xml.parse`)
- **Shared Interface** — `XmlSink` lets rendering code target both write modes transparently
- **XML Encoding** — Fast entity encode/decode with special character handling
- **Namespace Support** — Full XML Namespaces with prefix resolution, reserved namespace enforcement, and unbound prefix detection
- **Query Engine** — Simplified path expressions for querying DOM trees
- **Security Hardened** — Entity expansion limits, nesting depth limits, duplicate attribute rejection, name injection prevention, BOM handling

---

## Quick Start

### Writing XML (Buffered)

```typescript
import { Xml } from "documonster/xml";

const w = new Xml.Writer();
w.openXml(Xml.DEFAULT_DECLARATION);
w.openNode("root", { version: "1.0" });
w.leafNode("item", { id: "1" }, "hello");
w.leafNode("item", { id: "2" }, "world");
w.closeNode();

console.log(w.toString());
// <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
// <root version="1.0"><item id="1">hello</item><item id="2">world</item></root>
```

### Writing XML (Streaming)

```text
import { Xml } from "documonster/xml";

const chunks: string[] = [];
const target = { write: (chunk: string) => chunks.push(chunk) };
const sw = new Xml.StreamWriter(target);

sw.openXml();
sw.openNode("root");
sw.leafNode("item", { id: "1" }, "hello");
sw.closeNode();
// Writes go straight to the target; only start tags with nothing after them yet are held
```

### Reading XML (SAX — Streaming)

```text
import { Xml } from "documonster/xml";

const parser = new Xml.SaxParser();
parser.on("opentag", tag => console.log("open:", tag.name, tag.attributes));
parser.on("text", text => console.log("text:", text));
parser.on("closetag", tag => console.log("close:", tag.name));
// Each event has one handler: a second `on("text", …)` replaces the first.
parser.write('<root><item id="1">hello</item></root>');
parser.close();
```

### Reading XML (DOM — Buffered)

```text
import { Xml } from "documonster/xml";

const doc = Xml.parse('<root><item id="1">hello</item></root>');
const item = Xml.findChild(doc.root, "item");
console.log(Xml.getAttr(item!, "id")); // "1"
console.log(Xml.getText(item!)); // "hello"
```

### XML to Plain Object

Convert XML into plain JavaScript objects.

Two entry points for different scenarios:

```text
import { Xml } from "documonster/xml";

// Option 1: already have a DOM tree
const doc = Xml.parse('<root attr="1"><item>a</item><item>b</item></root>');
const obj = Xml.toPlainObject(doc.root);
// { root: { "@_attr": "1", item: ["a", "b"] } }

// Option 2: XML string → plain object directly (faster, single SAX pass)
const obj2 = Xml.parseToObject('<root attr="1"><item>a</item><item>b</item></root>');
// same output, ~1.6x faster on medium/large XML
```

**When to use which:**

- `Xml.toPlainObject(element)` — when you already have an `XmlElement` from `Xml.parse()`
- `Xml.parseToObject(xml)` — when you only need the plain object (skips DOM allocation)

**Default conversion rules:**

- Attributes are prefixed with `@_`
- Repeated sibling elements become arrays
- Text-only elements collapse to their string value
- Empty elements become `""`
- Whitespace-only indentation text is discarded by default

**Limitations:** plain-object conversion is intentionally lossy — it does not preserve element ordering, comments, or processing instructions. If you need exact XML structure, use `Xml.parse()` and work with the DOM tree directly.

### Query Engine

```typescript
import { Xml } from "documonster/xml";

const doc = Xml.parse("<root><a><b>1</b><b>2</b></a><a><b>3</b></a></root>");
const first = Xml.query(doc.root, "a/b"); // first <b> element
const all = Xml.queryAll(doc.root, "a/b"); // all <b> elements
const indexed = Xml.queryAll(doc.root, "a/b[0]"); // first <b> under each <a>
const filtered = Xml.query(doc.root, "a/b[@id='x']"); // <b> with id="x"
const deep = Xml.queryAll(doc.root, "a//b"); // <b> at any depth under <a>
```

### Encoding/Decoding

```typescript
import { Xml } from "documonster/xml";

Xml.encode('<tag attr="val">'); // "&lt;tag attr=&quot;val&quot;&gt;"
Xml.decode("&lt;hello&gt;"); // "<hello>"
```

---

## Architecture

```
src/modules/xml/
├── types.ts              # Core types (XmlNode, XmlSink, XmlSaxTag, etc.)
├── errors.ts             # XmlError, XmlParseError, XmlWriteError
├── encode.ts             # Xml.encode, Xml.decode, Xml.validateName, Xml.encodeCData, etc.
├── writer.ts             # Xml.Writer (buffered, with rollback support)
├── stream-writer.ts      # Xml.StreamWriter (streaming, writes to XmlWritableTarget)
├── sax.ts                # Xml.SaxParser (event-driven) + Xml.parseSaxStream (async generator)
├── dom.ts                # Xml.parse + DOM query helpers + Xml.toPlainObject
├── to-object.ts          # Xml.parseToObject (SAX-direct, single-pass)
├── to-object-shared.ts   # Shared conversion logic (internal)
├── query.ts              # Simplified path query engine
├── index.ts              # Public API barrel
└── __tests__/            # Tests
```

### Write Path

```
XmlSink (interface)
├── Xml.Writer        — Builds XML as a string in memory
│                      Supports rollback/transactions (snapshot/commit/rollback)
│                      Best for: small-medium XML, speculative writes
│
└── Xml.StreamWriter  — Writes directly to an XmlWritableTarget
                       Never holds the whole document: memory grows with nesting
                       depth and the start tag still being written, not file size
                       Best for: large XML (worksheets with 100K+ rows)
```

### Read Path

```
Xml.SaxParser  — Event-driven streaming parser
│                      Feed chunks via write(), events fire synchronously
│                      Best for: large XML, when you only need specific elements
│
├── Xml.parse        — Builds a DOM tree (XmlDocument/XmlElement)
│   │                  Built on top of Xml.SaxParser — no duplicate parsing logic
│   │                  Best for: small-medium XML, when you need tree traversal
│   │
│   └── Xml.toPlainObject — Converts XmlElement DOM to plain JS object
│                       Best for: when you already have a DOM tree
│
├── Xml.parseToObject — SAX-direct to plain JS object (single pass, no DOM)
│                      ~1.6x faster than Xml.parse + Xml.toPlainObject
│                      Best for: XML string → plain object → JSON.stringify
│
└── Xml.parseSaxStream     — Async generator wrapping Xml.SaxParser for stream iteration
                       Best for: async pipelines (e.g. reading from zip streams)
```

---

## API Reference

### Xml.Writer

Built with `new Xml.Writer()`.

| Method                          | Description                             |
| ------------------------------- | --------------------------------------- |
| `openXml(attrs?)`               | Write `<?xml ...?>` declaration (first) |
| `openNode(name, attrs?)`        | Open an element                         |
| `closeNode()`                   | Close the most recent element           |
| `leafNode(name, attrs?, text?)` | Write a complete element in one call    |
| `addAttribute(name, value)`     | Add attribute to currently-open element |
| `addAttributes(attrs)`          | Add multiple attributes                 |
| `writeText(text)`               | Write escaped text content              |
| `writeRaw(xml)`                 | Write pre-escaped XML (a string)        |
| `writeCData(text)`              | Write `<![CDATA[...]]>` section         |
| `writeComment(text)`            | Write `<!--...-->` comment              |
| `closeAll()`                    | Close all open elements                 |
| `toString()`                    | Get the built XML string                |
| `snapshot()`                    | Take a rollback snapshot                |
| `commit()`                      | Discard snapshot (keep changes)         |
| `rollback()`                    | Restore to snapshot (discard changes)   |
| `reset()`                       | Clear everything                        |
| `depth`                         | Current nesting depth                   |
| `currentElement`                | Name of innermost open element          |
| `cursor`                        | Grows with each write; rollback lowers  |

### Xml.StreamWriter

Built with `new Xml.StreamWriter(target)`. Same methods as `Xml.Writer` (both implement `XmlSink`), except:

- No `toString()` — content is already written to target
- No `snapshot()` / `commit()` / `rollback()` — streaming is irreversible
- No `cursor` — not applicable for streaming
- No `reset()` — what was written to the target cannot be taken back

It keeps no copy of the document, only the open elements and the start tags that have nothing after
them yet — the innermost still taking attributes — so its memory grows with nesting depth plus those
pending tags rather than with the file. Pending tags go out with whatever follows them, in one call
of the target's `write()`.

Both writers refuse `openXml` once anything has been written (`XmlWriteError`): XML allows the
declaration only at the very start of a document. The parsers refuse a misplaced `<?xml …?>` as well.
(This package's Excel and Word readers parse a package part from its first `<`, so white space a
producer wrote before a part's declaration is no error there — a policy of those readers, not of the
XML parser.)

Both writers take an element's own attributes only — one an attribute map inherits is not written —
and refuse a namespace declaration no document can hold (`xmlns:xml` bound elsewhere, `xmlns:xmlns`,
a prefix bound to `""`, a reserved namespace bound to another prefix). A prefix whose binding they are
not shown is let be, since what they write may be part of a document that declares it. A target's
`write()` may not call back into its `Xml.StreamWriter` (`XmlWriteError`).

### Xml.SaxParser

| Method / Property       | Description                                                                         |
| ----------------------- | ----------------------------------------------------------------------------------- |
| `write(chunk)`          | Feed XML text, or the document's bytes as a `Uint8Array` (UTF-8, split anywhere)    |
| `close()`               | Signal end of input                                                                 |
| `stop()`                | From a handler: read nothing more, the rest of the chunk included                   |
| `on(event, handler)`    | Register event handler                                                              |
| `off(event)`            | Remove event handler                                                                |
| `line` / `column`       | Current position (when position tracking enabled)                                   |
| `startByte` / `endByte` | Inside a handler, the event's byte range in the file (when given bytes)             |
| `byte`                  | Offset of the next byte to parse (when given bytes)                                 |
| `getLexicalState()`     | A comparable value naming where the parser is in the grammar (see below)            |
| `getPrologContext()`    | The root's name and the internal subset's entities, once the root start tag is read |
| `closed` / `stopped`    | Whether the parser has been closed / stopped                                        |

**Events:** `opentag`, `closetag`, `text`, `cdata`, `comment`, `pi`, `error`

A processing instruction in the DOCTYPE internal subset is reported too (`Xml.parse` does not put it
in `prologue`, which holds the document's own nodes). After a handler calls `stop()` nothing more is
reported, an error found on the way out included. A handler may not call `write()` or `close()` on the
parser that is calling it: that is refused with an `XmlError`.

**Options:**

| Option                     | Default    | Description                                                                                 |
| -------------------------- | ---------- | ------------------------------------------------------------------------------------------- |
| `position`                 | `true`     | Track line/column for error messages                                                        |
| `fragment`                 | `false`    | Allow multiple root elements                                                                |
| `xmlns`                    | `false`    | Enable namespace processing                                                                 |
| `maxDepth`                 | `256`      | Maximum element nesting depth                                                               |
| `maxEntityExpansions`      | `10000`    | Maximum entity expansion count (XML bomb defense)                                           |
| `maxEntityExpansionLength` | `10000000` | Maximum total length of entity expansions, in UTF-16 code units                             |
| `invalidCharHandling`      | `"error"`  | How to handle invalid XML characters (see below)                                            |
| `fileName`                 | —          | File name for error messages                                                                |
| `byteOffset`               | `0`        | Offset in the file of the first byte given                                                  |
| `resume`                   | `false`    | Start inside the root element, at an unknown depth                                          |
| `internalEntities`         | `false`    | Read and expand the internal subset's general entities                                      |
| `maxContextBytes`          | `65536`    | Largest prolog context (UTF-8 bytes of its JSON) and markup declaration (UTF-16 code units) |
| `context`                  | —          | A prolog context, for a `resume` parser                                                     |
| `skipLeadingText`          | `false`    | Skip what comes before the document's first `<` (see below)                                 |

**Bytes.** Given `Uint8Array` chunks, the parser decodes UTF-8 itself — a character split between
chunks is read whole — and gives every event its byte range (`startByte` / `endByte`) and every
error its `byte`. A byte that is not part of well-formed UTF-8 is an invalid character, handled as
`invalidCharHandling` says, so byte positions never drift. A declared encoding other than UTF-8 (or
ASCII) is refused; after a declaration of ASCII, a byte past `0x7F` is no character of the document's
encoding and is handled the same way. A document is given text or bytes, not both.

**Resuming.** `resume: true` starts in the content of some element: end tags of elements opened
before the start are passed over, a root element is not required, and a start inside a multi-byte
character drops the rest of it. Elements opened since the start and still open at the end are
reported; with `context`, so is an end that comes after a tag but before the root's end tag.

`getLexicalState()` tells when two parsers reading **one input** — one from its start, one resumed
further on — have reached the same state at the same position; from there they report the same
events. It holds the scanner state, any open quote or half-read reference, whether the prolog is
past, and where the current markup or text began — not what has been gathered since (the same input
for both) nor the elements open around it. Parsers given different inputs can share a lexical state.

### Xml.parseSaxStream (Async Generator)

```text
Xml.parseSaxStream(
  iterable: AsyncIterable<string | Uint8Array> | Iterable<string | Uint8Array>,
  options?: XmlSaxOptions
): AsyncGenerator<XmlSaxEvent[]>
```

Each batch holds the events of one chunk as the discriminated union `XmlSaxEvent` — narrow on
`eventType`. `Xml.feedSax(parser, source)` takes the same chunks for a parser you drive with
`.on(...)` handlers. A `Uint8Array` chunk goes to `Xml.SaxParser` as bytes, so everything listed
under **Bytes** above applies: `byteOffset`, byte positions, and `invalidCharHandling` for bytes
that are not well-formed UTF-8. One document is given text chunks or byte chunks, not a mix of
the two. An `ArrayBuffer` is not accepted; wrap it in a `Uint8Array`.

### Xml.readRecords / Xml.readRoot

Read only what is asked for from a file's bytes — a `Uint8Array`, or an iterable or async iterable
of them, split anywhere.

```ts
import { Xml } from "documonster/xml";

// The root element and the prolog's context, reading no further than the root's `>`.
const root = await Xml.readRoot(firstChunks);
root.name; // "DSR"

// Every <Track> as an object, in one pass, holding one record at a time.
await Xml.readRecords(source, {
  tags: ["Track"],
  onRecord: record => {
    record.object; // { Track: { "@_id": "1", Title: "…" } } — as Xml.parseToObject reads its text (see below)
    record.startByte; // the `<` of its start tag; endByte is just past its end tag
  }
});
```

A record is an element named in `tags` that is not inside another record; one inside it is part
of it. `onRecord` is called in file order, and a promise it returns is awaited. `rawXml: true` gives
each record its text as well. `record.object` is what `Xml.parseToObject` gives for that text only
in the same entity context: a record that uses an entity the file's DOCTYPE declares has it
expanded, so parsing `record.xml` alone needs an equivalent declaration (and
`internalEntities: true`) to give the same object.

**Byte ranges.** With `offset` and `size`, only the records whose start tag begins in
`[offset, offset + size)` are read — so a file can be split into ranges read independently, by
separate workers, and their records in order are the whole file's. `source` gives the bytes from
`offset` on; past `offset` 0 the prolog is not read, so pass `context: root.context`.

A range read from its own start assumes it begins in content, and that is checked rather than
trusted: the read of each range also reads the next range from that range's start, until the two
readings agree. If they cannot — a range that begins inside a record holding an element of the
same name, or inside a comment or CDATA section full of tags — it throws `XmlChunkAmbiguityError`
(`chunkStart`, `atByte`). Read that range with the one before it. A range after the first that
begins before the root element's start tag — in the prolog, or in text `skipLeadingText` skips — can
be refused the same way, as it reads an XML declaration or a DOCTYPE as content: the first range reads
the prolog whole wherever it ends, so ranges longer than the prolog never start there. So **if every range of a split
reads without an error and `Xml.verifyPartition` accepts the results, its records are exactly the
whole file's.** Each read goes past its range
only to finish its last record and place the next range's start — on a typical file, about a
kilobyte.

**Publish only a whole batch.** One range's success promises nothing about the others: a range
stops once its own records are read, so it cannot see that the file is cut further on, or that a
later range cannot be placed. Each result therefore carries `structure`, a summary of what its
range saw — opaque, but plain JSON, so a worker can send it back. Once every range has succeeded,
pass the results, in `offset` order, to `Xml.verifyPartition`: it checks what no single range can —
that tags match across range boundaries, that the root is closed (the file is not truncated), that
there is one root, and that nothing but white space, comments and processing instructions lies
outside it — and throws if not.

Run every range of one fixed version of the file and stage each one's records. Publish only when
every range has succeeded **and** `Xml.verifyPartition` has passed; if anything fails, discard the
whole batch — the ranges that succeeded included — and read again (the failed range with the one
before it, or the file whole). `onRecord` is called as records are read, so a call that fails may
already have given some.

```ts
import { Xml } from "documonster/xml";
import type { XmlReadRecordsResult } from "documonster/xml";

// readRangeIntoStaging returns its Xml.readRecords result; ranges are in offset order.
const outcomes = await Promise.allSettled(ranges.map(readRangeIntoStaging));
try {
  const results: XmlReadRecordsResult[] = [];
  for (const outcome of outcomes) {
    if (outcome.status === "rejected") {
      throw outcome.reason;
    }
    results.push(outcome.value);
  }
  Xml.verifyPartition(results);
  await publishStaged();
} catch {
  await discardStaged(); // every range's records, the successful ones' included
}
```

Errors are thrown, never recovered from. Input that ends inside a record is an
`XmlTruncatedError` with `recordStartByte`, and a range read to the end of the input checks that the
document ends there as one does. The range from offset 0 reads the prolog through to the root's start
tag, wherever the range ends, so the prolog is checked whole. `Xml.verifyPartition` throws
`XmlTruncatedError` when the file ends inside an element, `XmlParseError` when the ranges do not make
one document, and `XmlChunkAmbiguityError` when a range starts inside markup outside the root element
— the root's end tag, a trailing comment — and reads its end as text it cannot place: read that range
with the one before it.

`maxDepth`, `maxEntityExpansions` and `maxEntityExpansionLength` bound the resources one call spends, so they count within one
call: a range does not know how deep the elements around its start are, nor how many entities the
ranges before it expanded. A file whose whole read exceeds a limit may therefore read without error
in ranges — with the same records, since a limit decides whether to stop, never what is read.

`Xml.readRoot` stops at the root element's `>` and closes the source. Any error before it is
thrown with its byte offset; `maxStartByte` (default 1,000,000) and `maxTagBytes` (default 65,536)
bound where the root may begin and how long its start tag may be, by byte position — the same
answer however the source is chunked. It expands the internal subset's entities in the root's
attributes, bounded by `maxEntityExpansions` (default 10,000) and `maxEntityExpansionLength`
(default 10,000,000 UTF-16 code units), as the parser options of the same names are.

**Text before the document.** Some files carry text before the XML that no document holds there —
saved from `unzip -c`, they begin with its two lines `Archive:  …` and `  inflating: …`. With
`skipLeadingText: true`, `Xml.SaxParser`, `Xml.readRoot` and `Xml.readRecords` take the first `<` as
the document's start: what comes before it is skipped, not reported as text, and the XML declaration
may follow it. Byte positions stay the file's, and the result says how much was skipped
(`skippedLeadingBytes`; `skippedLeadingText` on a parser). Nothing else is relaxed: a `<` in that text
begins the document there, and text after the root is still refused. Give the option to `readRoot`
and to every range of a partition alike. It is off by default, since such a file is not XML.

Both functions throw this module's errors for what they read. Two kinds of failure are not
`XmlError`s: a cancelled `signal` is the package's `AbortError` (`isAbortError` is true,
`isXmlError` false), and what `source` itself throws — or a callback such as `onRecord` or `isArray`
— passes out unchanged.

### Xml.parse

```text
Xml.parse(xml: string, options?: XmlParseOptions): XmlDocument;
```

**Options:**

| Option                     | Default    | Description                                      |
| -------------------------- | ---------- | ------------------------------------------------ |
| `comments`                 | `false`    | Include comment nodes in DOM tree                |
| `processingInstructions`   | `false`    | Include PI nodes in DOM tree                     |
| `cdataAsNodes`             | `false`    | Keep CDATA as explicit nodes vs merge into text  |
| `fragment`                 | `false`    | Allow multiple root elements                     |
| `xmlns`                    | `false`    | Enable namespace processing                      |
| `maxDepth`                 | `256`      | Maximum element nesting depth                    |
| `maxEntityExpansions`      | `10000`    | Maximum entity expansion count                   |
| `maxEntityExpansionLength` | `10000000` | Maximum total length of entity expansions        |
| `invalidCharHandling`      | `"error"`  | How to handle invalid XML characters (see below) |
| `internalEntities`         | `false`    | Read and expand the internal subset's entities   |

**Returns:** `XmlDocument` with:

| Field         | Type                                            | Description                                                |
| ------------- | ----------------------------------------------- | ---------------------------------------------------------- |
| `root`        | `XmlElement`                                    | First (or only) root element                               |
| `roots`       | `XmlElement[]`                                  | All root-level elements (useful in fragment mode)          |
| `declaration` | `Record<string, string> \| undefined`           | XML declaration attributes (version, encoding, standalone) |
| `prologue`    | `Array<XmlComment \| XmlProcessingInstruction>` | Top-level comments and PIs before the root (when enabled)  |
| `epilogue`    | `Array<XmlComment \| XmlProcessingInstruction>` | Top-level comments and PIs after the root (when enabled)   |

### DOM Helpers

| Function                       | Description                 |
| ------------------------------ | --------------------------- |
| `Xml.findChild(el, name)`      | First child element by name |
| `Xml.listChildren(el, name)`   | All child elements by name  |
| `Xml.getText(node)`            | Recursive text content      |
| `Xml.getAttr(el, name)`        | Get attribute value         |
| `Xml.eachElement(el, visitor)` | Depth-first traversal       |

### Xml.toPlainObject

```text
Xml.toPlainObject(
  element: XmlElement,
  options?: XmlToPlainObjectOptions
): Record<string, unknown>;
```

Convert an `XmlElement` DOM tree into a plain JavaScript object.

### Xml.parseToObject

```text
Xml.parseToObject(xml: string, options?: XmlParseToObjectOptions): Record<string, unknown>;
```

Parse an XML string directly into a plain JavaScript object in a single SAX pass. ~1.6x faster than `Xml.parse()` + `Xml.toPlainObject()` on medium/large XML.

**Conversion options** (shared by both functions):

| Option                 | Default   | Description                                                 |
| ---------------------- | --------- | ----------------------------------------------------------- |
| `ignoreAttributes`     | `false`   | Discard all attributes entirely                             |
| `attributePrefix`      | `"@_"`    | Prefix for attribute keys (`""` for bare names)             |
| `textKey`              | `"#text"` | Key for text content in mixed-content elements              |
| `alwaysArray`          | `false`   | Always wrap child elements in arrays                        |
| `isArray`              | —         | Callback `(name) => boolean` for per-tag array wrapping     |
| `preserveCData`        | `true`    | Include CDATA values in text (relevant with `cdataAsNodes`) |
| `ignoreWhitespaceText` | `true`    | Discard whitespace-only text in elements that have children |

**Parser options** (`Xml.parseToObject` only):

| Option                     | Default    | Description                                       |
| -------------------------- | ---------- | ------------------------------------------------- |
| `fragment`                 | `false`    | Allow multiple root elements                      |
| `maxDepth`                 | `256`      | Maximum element nesting depth                     |
| `maxEntityExpansions`      | `10000`    | Maximum entity expansion count (XML bomb defense) |
| `maxEntityExpansionLength` | `10000000` | Maximum total length of entity expansions         |
| `invalidCharHandling`      | `"error"`  | How to handle invalid XML characters (see below)  |
| `internalEntities`         | `false`    | Read and expand the internal subset's entities    |

When `textKey` names the same key as an attribute (`attributePrefix: ""`) or a child element, the
values are merged into an array rather than one overwriting the other.

`preserveCData: false` drops a CDATA section only where it is still one: `Xml.parseToObject` sees it,
but `Xml.parse` merges CDATA into the text around it unless given `cdataAsNodes: true`, so the two
paths agree on CDATA only with that option.

`Xml.toPlainObject`, `Xml.getText`, `Xml.eachElement` and `Xml.query*` take a tree a caller may have
built: an element's own attributes only are read, a node without the fields read is refused
(`XmlArgumentError`), and so is an element among its own descendants, which no walk would leave.
Each field is read once and the value checked is the value used. These helpers check only the fields
they read, as they reach them — they are not a validation of the whole model, so a malformed node they
never reach (past `Xml.query`'s first match, say) is not reported. A model that puts one element under
several parents is walked as the tree its paths spell out: `Xml.query*` reach the element, and what is
under it, along each path. One step that would reach more than 2^25 elements that way is refused —
no tree in memory holds that many, but sharing each level twice over makes that many paths of a small
model.

### Query Engine

```typescript
import { Xml } from "documonster/xml";
```

| Syntax         | Description                                     |
| -------------- | ----------------------------------------------- |
| `a/b/c`        | Match child `a`, then `b`, then `c`             |
| `a/b[@id='1']` | Match `b` with attribute `id` equal to `"1"`    |
| `a/*/c`        | Wildcard: any element name at that level        |
| `a//c`         | Recursive descent: `c` at any depth under `a`   |
| `a/b[0]`       | Index: first matching `b` under each parent `a` |
| `//c`          | `c` at any depth under the context element      |

- `Xml.query(element, path)` — First match or `undefined`
- `Xml.queryAll(element, path)` — All matches (may be empty), each once, in document order

Anything else — an absolute path (`/a`), an empty step (`a/`), `.`/`..`, or a predicate not listed above (`[@id]`, `[last()]`) — throws `XmlError` rather than matching nothing. Indexes are 0-based, unlike XPath.

Index filters use **per-parent semantics**: `a/b[0]` returns the first `b` under _each_ `a`, not the globally first `b`.

### Encoding Utilities

| Function                        | Description                                                   |
| ------------------------------- | ------------------------------------------------------------- |
| `Xml.encode(text)`              | Encode text for XML content (`<`, `>`, `&`, `"`, `'`)         |
| `Xml.decode(text)`              | Decode XML entities back to text                              |
| `Xml.encodeAttr(value)`         | Encode an attribute value (adds `\t\n\r` → `&#x9;&#xA;&#xD;`) |
| `Xml.validateName(name)`        | Validate an XML element/attribute name                        |
| `Xml.validateCommentText(text)` | Validate text for XML comment content                         |
| `Xml.encodeCData(text)`         | Encode text for a CDATA section (splits `]]>`)                |

### Error Types

| Class                    | Parent          | Used by                                          |
| ------------------------ | --------------- | ------------------------------------------------ |
| `XmlError`               | `BaseError`     | Encoding/validation                              |
| `XmlParseError`          | `XmlError`      | SAX parser, DOM parser (`byte` when given bytes) |
| `XmlLimitExceededError`  | `XmlParseError` | A configured limit (`limit`, `allowed`)          |
| `XmlTruncatedError`      | `XmlParseError` | Input that ends inside a construct or record     |
| `XmlChunkAmbiguityError` | `XmlError`      | `Xml.readRecords`: a range that cannot be placed |
| `XmlWriteError`          | `XmlError`      | Writers (state errors)                           |

The errors this module raises extend `XmlError`, so `catch (e) { if (e instanceof XmlError) ... }`
catches them. Not everything a call throws is one: `Xml.readRoot` / `Xml.readRecords` throw the
package's `AbortError` when cancelled, and pass on unchanged what a source or callback throws.

---

## Namespace Support

When `xmlns: true` is enabled:

- **Prefix resolution** — Element QNames are resolved to `prefix`, `local` and `uri` on the
  element (`XmlElement`, `XmlSaxTag`). Attribute prefixes are resolved and checked too, but
  attributes stay a plain `Record<string, string>` keyed by QName — no per-attribute metadata
- **Pre-bound prefixes** — `xml` is pre-bound to `http://www.w3.org/XML/1998/namespace`, `xmlns` to `http://www.w3.org/2000/xmlns/`
- **Reserved namespace enforcement** — Cannot rebind `xml` to a different URI, cannot rebind `xmlns`, cannot bind other prefixes to reserved URIs
- **Unbound prefix detection** — Elements and attributes with undeclared prefixes produce errors
- **Expanded-name duplicate detection** — Two attributes with different prefixes but same URI + local name are rejected
- **Multi-colon QName rejection** — `<a:b:c/>` is rejected in namespace mode
- **Scope management** — Namespace declarations follow XML scoping rules (inherited by descendants, overridable in children)

Note: Unprefixed attributes do **not** inherit the default namespace, per XML Namespaces §6.2.

---

## Security

- **Entity expansion limits** — Prevents XML bomb attacks (configurable via `maxEntityExpansions` and `maxEntityExpansionLength`)
- **Nothing outside the document is read** — an external DTD subset, an external entity
  (`<!ENTITY e SYSTEM "…">`) and a parameter entity are recognised and never fetched or expanded. A
  reference to an external entity is an `undefined entity` error, and no entity declared after a
  parameter-entity reference is processed. No option changes this: no file is opened and no
  network request is made.
- **Internal entities only on request** — By default the DOCTYPE internal subset is checked for
  well-formedness (comments, processing instructions, literals and declarations, so a `]`, `>` or
  quote inside them does not end it) but not read, and a reference to an entity declared there is
  reported as `undefined entity: e (entities declared in a DOCTYPE internal subset are not
supported)`. With `internalEntities: true` (always on for `Xml.readRecords` / `Xml.readRoot`) its
  general entities expand as XML 1.0 §4.4 describes; replacement text that holds markup (`<`) or a
  `]]>` and recursive entities are refused, and every nested expansion counts against `maxEntityExpansions` —
  and the length of the text it makes against `maxEntityExpansionLength` — before any text is
  built, so the text entities make is bounded by `maxEntityExpansionLength` (UTF-16 code units) rather
  than by whatever the input asks for. That is a bound on expansion, not on the parser's memory as a
  whole: text being read, caches and what a writer's target buffers are not counted in it. The five predefined entities and character references are always
  available; a caller can add others through the parser's `ENTITIES` map — only its own string
  values count, and only a reference that is an XML Name, so `&constructor;` is undefined rather than
  `Object` and `&1;` is refused whatever the map holds. `maxEntityExpansionLength`
  counts the text such a caller-supplied entity expands to as well.
- **Attribute-list declarations are used** — as XML 1.0 §5.1 asks of a parser that reads the
  internal subset, whatever `internalEntities` says: an element that leaves out an attribute declared
  with a default is given it, and a value of a type other than CDATA has its spaces collapsed and
  trimmed. A default that refers to an entity declared there is expanded with `internalEntities`, and
  is an `undefined entity` error where it is used without it; an expanded default counts against
  `maxEntityExpansions` and `maxEntityExpansionLength` once, where it is declared: the elements given
  it share the one string. A document that says
  `standalone="yes"` is held to declaring what it uses, and its declarations after an unread
  parameter-entity reference are still read. `Xml.readRoot`'s `context` carries the declarations
  (`attributes`), so a range reads its records as the whole file does. What the parser keeps of the
  internal subset to read the rest by (entity bindings, attribute declarations) is bounded by
  `maxContextBytes` too, counted in characters, whether or not `internalEntities` is on.
- **Nesting depth limits** — Prevents stack overflow from deeply nested XML (configurable via `maxDepth`)
- **Duplicate attribute rejection** — XML 1.0 §3.1 WFC: Unique Att Spec (reports error, recovers with last-value-wins)
- **Name injection prevention** — Writers validate element and attribute names via `Xml.validateName()`
- **Comment/CDATA safety** — `Xml.validateCommentText()` rejects `--`, `Xml.encodeCData()` splits `]]>`
- **BOM handling** — UTF-8 BOM at start of input is silently stripped
- **Prototype pollution prevention** — Attribute maps and converted objects are null-prototype
  objects, so keys such as `__proto__`, `constructor` and `prototype` are kept as ordinary data and
  cannot reach `Object.prototype`
- **Invalid character handling** — Writers strip invalid XML 1.0 characters; parser behavior is configurable via `invalidCharHandling`

---

## Invalid Character Handling

Real-world XML data (especially from third-party XLSX files) may contain characters that are invalid per XML 1.0 — for example `0x01`–`0x08`, `0x0B`, `0x0C`, `0x0E`–`0x1F`, lone surrogates, and non-characters `U+FFFE`/`U+FFFF`. `0x7F` (DEL) is treated the same way. XML 1.0 itself allows DEL, only discouraging it; refusing it is this library's policy, as it has no use in the documents it reads and writes, and a character reference `&#x7F;` is still read.

The `invalidCharHandling` option controls how the parser responds:

| Value       | Behavior                                            |
| ----------- | --------------------------------------------------- |
| `"error"`   | Report via error handler or throw **(default)**     |
| `"skip"`    | Silently discard the invalid character              |
| `"replace"` | Replace with U+FFFD (Unicode REPLACEMENT CHARACTER) |

### Examples

```typescript
import { Xml } from "documonster/xml";

// Default: strict mode — throws on 0x7F
Xml.parse("<root>hello\x7fworld</root>");
// => XmlParseError: invalid XML character: 0x7f

// Skip mode — invalid chars are removed
const doc = Xml.parse("<root>hello\x7fworld</root>", { invalidCharHandling: "skip" });
// doc.root text content: "helloworld"

// Replace mode — invalid chars become U+FFFD
const doc2 = Xml.parse("<root>hello\x7fworld</root>", { invalidCharHandling: "replace" });
// doc.root text content: "hello\uFFFDworld"

// SAX parser with skip mode
const parser = new Xml.SaxParser({ invalidCharHandling: "skip" });
parser.on("text", text => console.log(text)); // "helloworld"
parser.write("<root>hello\x7fworld</root>");
parser.close();
```

### When to use which

- **`"error"` (default)** — Use for strict XML validation, testing, or when you control the XML source.
- **`"skip"`** — Use when reading untrusted/dirty XML (e.g., third-party XLSX files) where you want to silently discard bad characters. This is what the Excel XLSX reader uses internally.
- **`"replace"`** — Use when you want to preserve the _position_ of invalid characters (e.g., for diagnostics or data forensics) without crashing the parser.

> **Note:** The XML _writers_ (`Xml.Writer`, `Xml.StreamWriter`) always strip invalid characters via `Xml.encode()` — this option only affects the _parser_.
