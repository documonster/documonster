# XML Module

A zero-dependency, cross-platform XML toolkit for reading and writing XML. Supports both streaming and buffered modes.

```typescript
import { XmlWriter, SaxParser, parseXml } from "@xml/index";
```

## Features

- **Zero Dependencies** — Pure TypeScript, no native addons
- **Cross-Platform** — Same API in Node.js and browsers
- **Dual-Mode Writing** — Buffered (`XmlWriter`) and streaming (`XmlStreamWriter`)
- **Dual-Mode Reading** — SAX streaming (`SaxParser`) and DOM tree (`parseXml`)
- **Shared Interface** — `XmlSink` lets rendering code target both write modes transparently
- **XML Encoding** — Fast entity encode/decode with special character handling

---

## Quick Start

### Writing XML (Buffered)

```typescript
import { XmlWriter, StdDocAttributes } from "@xml/writer";

const w = new XmlWriter();
w.openXml(StdDocAttributes);
w.openNode("root", { version: "1.0" });
w.leafNode("item", { id: "1" }, "hello");
w.leafNode("item", { id: "2" }, "world");
w.closeNode();

console.log(w.xml);
// <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
// <root version="1.0"><item id="1">hello</item><item id="2">world</item></root>
```

### Writing XML (Streaming)

```typescript
import { XmlStreamWriter } from "@xml/stream-writer";

const chunks: string[] = [];
const target = { write: (chunk: string) => chunks.push(chunk) };
const sw = new XmlStreamWriter(target);

sw.openXml();
sw.openNode("root");
sw.leafNode("item", { id: "1" }, "hello");
sw.closeNode();
// Each method call writes directly to target — no buffering
```

### Reading XML (SAX — Streaming)

```typescript
import { SaxParser } from "@xml/sax";

const parser = new SaxParser();
parser.on("opentag", tag => console.log("open:", tag.name, tag.attributes));
parser.on("text", text => console.log("text:", text));
parser.on("closetag", tag => console.log("close:", tag.name));
parser.write('<root><item id="1">hello</item></root>');
parser.close();
```

### Reading XML (DOM — Buffered)

```typescript
import { parseXml, findChild, textContent, attr } from "@xml/dom";

const doc = parseXml('<root><item id="1">hello</item></root>');
const item = findChild(doc.root, "item");
console.log(attr(item!, "id")); // "1"
console.log(textContent(item!)); // "hello"
```

### Encoding/Decoding

```typescript
import { xmlEncode, xmlDecode } from "@xml/encode";

xmlEncode('<tag attr="val">'); // "&lt;tag attr=&quot;val&quot;&gt;"
xmlDecode("&lt;hello&gt;"); // "<hello>"
```

---

## Architecture

```
src/modules/xml/
├── types.ts           # Core types (XmlNode, XmlSink, SaxTag, etc.)
├── errors.ts          # XmlError, XmlParseError, XmlWriteError
├── encode.ts          # xmlEncode, xmlDecode, xmlEncodeAttr
├── writer.ts          # XmlWriter (buffered, with rollback support)
├── stream-writer.ts   # XmlStreamWriter (streaming, writes to WritableTarget)
├── sax.ts             # SaxParser (event-driven) + parseSax (async generator)
├── dom.ts             # parseXml + DOM query helpers
├── index.ts           # Public API barrel
└── __tests__/         # Tests
```

### Write Path

```
XmlSink (interface)
├── XmlWriter        — Builds XML as a string in memory
│                      Supports rollback/transactions (save/commit/rollback)
│                      Best for: small-medium XML, speculative writes
│
└── XmlStreamWriter  — Writes directly to a WritableTarget
                       O(1) memory — never holds full document
                       Best for: large XML (worksheets with 100K+ rows)
```

### Read Path

```
SaxParser            — Event-driven streaming parser
│                      Feed chunks via write(), events fire synchronously
│                      Best for: large XML, when you only need specific elements
│
└── parseXml         — Builds a DOM tree (XmlDocument/XmlElement)
                       Built on top of SaxParser — no duplicate parsing logic
                       Best for: small-medium XML, when you need tree traversal
```

---

## API Reference

### XmlWriter

| Method                          | Description                             |
| ------------------------------- | --------------------------------------- |
| `openXml(attrs?)`               | Write `<?xml ...?>` declaration         |
| `openNode(name, attrs?)`        | Open an element                         |
| `closeNode()`                   | Close the most recent element           |
| `leafNode(name, attrs?, text?)` | Write a complete element in one call    |
| `addAttribute(name, value)`     | Add attribute to currently-open element |
| `addAttributes(attrs)`          | Add multiple attributes                 |
| `writeText(text)`               | Write escaped text content              |
| `writeRaw(xml)`                 | Write pre-escaped XML                   |
| `writeCData(text)`              | Write `<![CDATA[...]]>` section         |
| `writeComment(text)`            | Write `<!--...-->` comment              |
| `closeAll()`                    | Close all open elements                 |
| `toString()` / `xml`            | Get the built XML string                |
| `save()`                        | Save rollback snapshot                  |
| `commit()`                      | Discard snapshot (keep changes)         |
| `rollback()`                    | Restore to snapshot (discard changes)   |
| `reset()`                       | Clear everything                        |
| `depth`                         | Current nesting depth                   |
| `currentElement`                | Name of innermost open element          |
| `cursor`                        | Monotonic position counter              |

### XmlStreamWriter

Same methods as `XmlWriter` (both implement `XmlSink`), except:

- No `toString()` / `xml` — content is already written to target
- No `save()` / `commit()` / `rollback()` — streaming is irreversible
- No `cursor` — not applicable for streaming

### SaxParser

| Method / Property    | Description                                       |
| -------------------- | ------------------------------------------------- |
| `write(chunk)`       | Feed XML text (can be called multiple times)      |
| `close()`            | Signal end of input                               |
| `on(event, handler)` | Register event handler                            |
| `off(event)`         | Remove event handler                              |
| `line` / `column`    | Current position (when position tracking enabled) |
| `closed`             | Whether parser has been closed                    |

**Events:** `opentag`, `closetag`, `text`, `cdata`, `comment`, `pi`, `error`

### parseSax (Async Generator)

```typescript
async function* parseSax(
  iterable: AsyncIterable<string | Uint8Array | ArrayBuffer>,
  options?: SaxOptions
): AsyncGenerator<SaxEventAny[]>
```

### parseXml

```typescript
function parseXml(xml: string, options?: XmlParseOptions): XmlDocument;
```

**Options:**

- `comments` — Include comment nodes (default: false)
- `processingInstructions` — Include PI nodes (default: false)
- `cdataAsNodes` — Keep CDATA as explicit nodes vs merge into text (default: false)
- `fragment` — Parse without requiring a root element (default: false)

### DOM Helpers

| Function                 | Description                 |
| ------------------------ | --------------------------- |
| `findChild(el, name)`    | First child element by name |
| `findChildren(el, name)` | All child elements by name  |
| `textContent(node)`      | Recursive text content      |
| `attr(el, name)`         | Get attribute value         |
| `walk(el, visitor)`      | Depth-first traversal       |
