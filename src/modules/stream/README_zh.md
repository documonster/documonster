# Stream 模块

[English](README.md)

跨平台流式实现，在 Node.js 和浏览器中提供相同的 API。

- **Node.js**：使用原生 `stream` 模块以获得最佳性能
- **浏览器**：使用 Web Streams API 的兼容实现

## 功能特性

- **100% 跨平台** — Node.js 和浏览器行为一致
- **单一入口** — 一套 API，两个运行时
- **类型安全** — 完整 TypeScript 泛型支持
- **兼容 Node.js API** — Readable、Writable、Transform、Duplex
- **管道和工具** — `pipeline()`、`finished()`、`createTransform()`
- **流消费** — `text()`、`buffer()`、`arrayBuffer()`、`blob()`
- **二进制工具** — `concatUint8Arrays`、`toUint8Array`、编码/解码
- **DuplexPair** — 双向通信通道
- **背压** — 自动流量控制

## 快速开始

```typescript
import {
  Readable,
  Writable,
  Transform,
  pipeline,
  createTransform
} from "@cj-tech-master/excelts/stream";

// 创建可读流
const readable = Readable.from(["hello", "world"]);

// 创建转换流
const upper = createTransform<string, string>({
  transform(chunk, _encoding, callback) {
    callback(null, chunk.toUpperCase());
  }
});

// 管道连接
const chunks: string[] = [];
const writable = new Writable({
  write(chunk, _encoding, callback) {
    chunks.push(String(chunk));
    callback();
  }
});

await pipeline(readable, upper, writable);
console.log(chunks); // ["HELLO", "WORLD"]
```

## 示例

查看 [examples 目录](examples/) 获取可运行代码。
