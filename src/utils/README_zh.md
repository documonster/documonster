# Utils

[English](README.md)

各平台支持程度不一的字节与文本转换。在 Node.js、Bun 和所有受支持的浏览器上行为一致。

```ts
import { Bytes } from "documonster/utils";
// Bytes.fromBase64, Bytes.toBase64, Bytes.fromHex, Bytes.toHex, Bytes.concat,
// Bytes.fromText, Bytes.toText, Bytes.random
```

## 为什么需要

- **Base64**：`Uint8Array.fromBase64` / `toBase64` 属于 ES2026，Node 22、24 以及早于 Chrome 140、
  Firefox 133、Safari 18.2 的浏览器都没有。替代它们的解码器对不合法文本的处理也不一致：Node 的
  `Buffer.from(text, "base64")` 会跳过不认识的字符，`atob` 则抛出 `DOMException`。`Bytes.fromBase64`
  在所有平台上都拒绝这类文本，有原生解码器时直接使用原生实现，方法名和选项也与原生一致。
- **Hex**：`Uint8Array.fromHex` / `toHex` 同样属于 ES2026，缺口相同。`Bytes.fromHex` 只接受成对的
  `0-9`、`a-f`、`A-F`，`Bytes.toHex` 输出小写。
- **拼接**：`Uint8Array` 没有原生的拼接方法，`Buffer.concat` 只有 Node 有。
- **随机字节**：`crypto.getRandomValues` 单次最多 65,536 字节，`Bytes.random` 会分块填充。
- **文本**：基于 `TextEncoder` / `TextDecoder`，出错时报本入口自己的错误。

所有结果都是普通的 `Uint8Array`，底层是独占的 `ArrayBuffer`：不会是 Node 的 `Buffer`，也不会是别人内存的视图。

## 用法

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

## Base64 规则

任意位置的 ASCII 空白都会被忽略。填充可以省略，但如果写了，就必须补全最后一组（`YQ` 和 `YQ==`
可以读，`YQ=` 会被拒绝）。字母表之外的其他字符一律拒绝，包括另一种字母表的 `+` `/` 或 `-` `_`；
末尾单独多出一个字符也会被拒绝。这与 `atob` 以及默认模式下的 `Uint8Array.fromBase64` 规则相同。
