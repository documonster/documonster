# Utils

[中文](README_zh.md)

Byte and text conversions the platforms offer unevenly, with one behaviour on Node.js, Bun and
every supported browser.

```ts
import { Bytes } from "documonster/utils";
// Bytes.fromBase64, Bytes.toBase64, Bytes.fromHex, Bytes.toHex, Bytes.concat,
// Bytes.fromText, Bytes.toText, Bytes.random
```

## Why

- **Base64** — `Uint8Array.fromBase64` / `toBase64` are ES2026: absent from Node 22 and 24 and
  from browsers older than Chrome 140, Firefox 133 and Safari 18.2. The decoders that stand in
  for them disagree about malformed text — Node's `Buffer.from(text, "base64")` skips characters
  it does not know, `atob` throws a `DOMException`. `Bytes.fromBase64` refuses them everywhere,
  and uses the native decoder where there is one. Its names and options are the native ones.
- **Hex** — `Uint8Array.fromHex` / `toHex` are ES2026 too, with the same gaps. `Bytes.fromHex`
  takes pairs of `0-9`, `a-f`, `A-F` and nothing else; `Bytes.toHex` writes lower case.
- **Concatenation** — there is no native `Uint8Array` concatenation; `Buffer.concat` is Node's.
- **Random bytes** — `crypto.getRandomValues` refuses more than 65,536 bytes in one call;
  `Bytes.random` fills in pieces.
- **Text** — `TextEncoder` / `TextDecoder`, with a refusal reported as this entry's error.

Every result is a plain `Uint8Array` over an `ArrayBuffer` of its own: never a Node `Buffer`, and
never a view of memory something else holds.

## Usage

```ts
import { Bytes, UtilsArgumentError } from "documonster/utils";

const bytes = Bytes.fromBase64("aGVsbG8=");
Bytes.toText(bytes); // "hello"
Bytes.toBase64(bytes, { alphabet: "base64url", omitPadding: true }); // "aGVsbG8"

Bytes.concat([Bytes.fromText("a"), Bytes.random(16)]);

try {
  Bytes.fromBase64("aGVs!bG8=");
} catch (error) {
  if (error instanceof UtilsArgumentError) {
    console.log(error.field); // "text"
  }
}
```

## Base64 rules

ASCII white space is ignored anywhere. Padding is optional, but padding that is present must
complete the last group (`YQ` and `YQ==` are read, `YQ=` is refused). Any other character outside
the alphabet is refused — including the other alphabet's `+` `/` or `-` `_` — and so is a lone
final character. These are the rules of `atob` and of `Uint8Array.fromBase64` in its default
mode.
