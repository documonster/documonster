# Archive 模块

[English](README.md)

零依赖、跨平台的归档工具包，用于创建、读取和编辑 ZIP 和 TAR 归档。

```typescript
import { Archive } from "documonster/archive";
```

## 功能特性

- **零依赖** — 纯 TypeScript，无原生插件
- **跨平台** — Node.js 和浏览器使用相同 API
- **ZIP + TAR** — 两种归档格式的统一 API
- **流式处理** — 真正的异步流式处理，支持背压
- **ZIP 编辑** — 对现有归档进行文件系统式编辑，高效直通传递
- **远程读取** — 通过 HTTP Range 请求读取 ZIP 文件（只下载需要的部分）
- **加密** — ZIP 传统加密和 AES-256 加密/解密
- **ZIP64** — 大文件支持（> 4GB 文件、> 65535 条目）
- **压缩** — DEFLATE、GZIP 和 Zlib，支持同步/异步/流式 API
- **PNG 编码** — `Archive.encodePng`，因为 PNG 就是带 CRC-32 校验 chunk 的 DEFLATE 流
- **进度和中止** — 内置进度回调和 AbortSignal 支持
- **文件系统集成** — Node.js 磁盘 I/O 便捷层

---

## 快速开始

### 创建 ZIP

```typescript
import { Archive } from "documonster/archive";

// 便捷函数
const archive = Archive.createZip();
archive.add("hello.txt", "Hello, World!");
archive.add("data.json", JSON.stringify({ key: "value" }));
const bytes = await archive.toBuffer();

// 使用选项
const archive = Archive.createZip({ level: 9, comment: "My archive" });
archive.add("file.txt", content, { modTime: new Date() });
archive.addDirectory("empty-dir/");
archive.addSymlink("link.txt", "hello.txt");
const bytes = await archive.toBuffer();
```

### 读取 ZIP

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

// 按路径随机访问
const data = await reader.bytes("hello.txt");
```

### 编辑 ZIP

```typescript
import { Archive } from "documonster/archive";

const editor = await Archive.ZipEditor.open(existingZipBytes, { preserve: "best-effort" });
editor.delete("old.txt");
editor.rename("a.txt", "renamed.txt");
editor.set("new.txt", "new content");
const output = await editor.toBuffer();
```

### 一次调用，无需构建器

所有条目都已在手时，一个函数生成归档，另一个函数把它读回为 路径 → `{ info, data }` 的 `Map`：

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

// 内存数据源可使用同步形式
const same = Archive.unpackZipSync(Archive.packZipSync(new Map([["a.txt", "A"]])));
```

### 各任务对应的函数

所有格式形状一致：需要逐条目选项、进度、流式输出或惰性读取时，使用构建器和读取器；不需要时，使用一次性函数。

| 任务                             | ZIP                        | TAR                         | tar.gz                          |
| -------------------------------- | -------------------------- | --------------------------- | ------------------------------- |
| 构建器                           | `Archive.createZip()`      | `Archive.createTar()`       | `Archive.createTarGz()`         |
| 读取器                           | `Archive.openZip(src)`     | `Archive.openTar(src)`      | `Archive.openTarGz(src)`        |
| 条目 → 字节                      | `Archive.packZip(entries)` | `Archive.packTar(entries)`  | `Archive.packTarGz(entries)`    |
| 条目 → 字节（同步）              | `Archive.packZipSync(…)`   | `Archive.packTarSync(…)`    | `Archive.packTarGzSync(…)`      |
| 字节 → `Map<path, {info, data}>` | `Archive.unpackZip(src)`   | `Archive.unpackTar(src)`    | `Archive.unpackTarGz(src)`      |
| 字节 → Map（同步）               | `Archive.unpackZipSync(…)` | `Archive.unpackTarSync(…)`  | `Archive.unpackTarGzSync(…)`    |
| 字节 → 条目列表                  | —（用读取器）              | `Archive.parseTar(bytes)`   | `Archive.parseTarGz(src)`       |
| 字节 → 条目流                    | —（用读取器）              | `Archive.parseTarStream(…)` | `Archive.parseTarGzStream(src)` |

每个动词只有一个含义，格式写在名字里：`create*` 返回构建器，`open*` 返回读取器，`pack*` 返回字节，`unpack*` 返回路径 Map，`parse*` 返回条目列表。`entries` 可以是 `Map<string, source>`，也可以是 `[name, source]` 二元组或 `{ name, source }` 对象的可迭代集合。同步形式只接受内存中的数据源，在 Node 和浏览器上都可用。构建器与读取器类（`ZipArchive`、`ZipReader`、`TarArchive`、`TarReader`、`TarGzArchive`）作为值公开，可用于 `instanceof` 判断和继承。

已弃用的 `Archive.zip` / `Archive.unzip`（按 `options.format` 选择格式）、`Archive.tar` / `tarSync` / `untar`、`createTarArchive` / `createTarReader` 仍然可用，行为不变，下一个次版本移除。

---

## ZIP API

### `Archive.createZip(options?)`

创建 `ZipArchive` 构建器。TAR 用 `Archive.createTar()`，tar.gz 用 `Archive.createTarGz()`。

```typescript
import { Archive } from "documonster/archive";

const archive = Archive.createZip(); // ZipArchive
const tarArchive = Archive.createTar(); // TarArchive
```

### `ZipArchive`

流式 ZIP 归档构建器，支持链式调用。

```typescript
import { Archive } from "documonster/archive";

const archive = Archive.createZip({ level: 6, reproducible: true });

// 添加条目（可链式调用）
archive
  .add("file.txt", "text content")
  .add("binary.dat", uint8Array)
  .add("from-stream.bin", readableStream)
  .addDirectory("empty-dir/")
  .addSymlink("link.txt", "file.txt");

// 输出选项：

// 1. 完整 Uint8Array
const bytes = await archive.toBuffer();

// 2. 同步（仅限内存数据源）
const bytes = archive.toBufferSync();

// 3. 异步流式
for await (const chunk of archive.stream()) {
  // 逐块处理
}

// 4. 管道到 WritableStream
await archive.pipeTo(writableStream);

// 5. 带进度和中止的操作
const op = archive.operation({
  onProgress: p => console.log(`${p.entriesDone}/${p.entriesTotal}`),
  signal: abortController.signal
});
for await (const chunk of op.iterable) {
  /* process chunk */
}
```

**`ZipOptions`：**

| 选项           | 类型                       | 默认值    | 描述                           |
| -------------- | -------------------------- | --------- | ------------------------------ |
| `level`        | `number`                   | `6`       | 压缩级别（0 = 存储，9 = 最大） |
| `comment`      | `string`                   | -         | 归档级注释                     |
| `reproducible` | `boolean`                  | `false`   | 稳定时间戳，用于可复现构建     |
| `smartStore`   | `boolean`                  | `true`    | 自动对不可压缩数据使用 STORE   |
| `zip64`        | `"auto" \| boolean`        | `"auto"`  | ZIP64 模式                     |
| `signal`       | `AbortSignal`              | -         | 中止信号                       |
| `onProgress`   | `(p: ZipProgress) => void` | -         | 进度回调                       |
| `encoding`     | `ZipStringEncoding`        | `"utf-8"` | 名称的字符串编码               |
| `timestamps`   | `ZipTimestampMode`         | -         | 时间戳编码模式                 |
| `path`         | `false \| ZipPathOptions`  | `false`   | 条目名规范化模式               |

条目名在添加时检查。绝对路径（`/etc/passwd`、`C:x`）或用 `..` 越出归档根目录的名称（`../evil.txt`）
会使 `add` 抛出 `ArchiveError`；`level` 超出 0–9、无效的 `Date`、头部无法容纳的 `mode` 同样如此，
`createTar().add` 采用相同的名称规则。传入 `path` 模式则改由该模式决定：`{ mode: "safe" }`
规范化并拒绝，`{ mode: "posix" }` 规范化并保留开头的 `..`。
不是 `ArchiveSource` 的条目数据同样在 `add` 时被拒绝，而不是在写出归档时。

时间戳：每个 ZIP 都带有的 DOS 日期和时间字段是本地挂钟时间，按 `modTime` 的本地字段写入、按本地时间读回，
与 Info-ZIP、7-Zip 和 Windows 一致；它们能表示 1980–2107 年。`timestamps: "dos+utc"` 会加上 Info-ZIP 的
UT 字段，它保存 1970–2106 年间精确的 UTC 时间，读取时优先使用。即使在默认的 `"dos"` 模式下，只要单靠 DOS 字段
读回的时间与 `modTime` 不同——1980 年之前的日期，或时区重复的挂钟时间（夏令时回拨的那一小时）——条目就会带上 UT 字段，
因此 1970–2106 年间的每个 `modTime` 都按写入时读回（精确到 DOS 字段的两秒）。只有 DOS 字段能表示的（2106-02-07 至 2107 年）则不带；
两者都无法表示的 `modTime` 会以 `ArchiveError` 拒绝。取自磁盘文件的时间则被限制到 DOS 范围内。

### `Archive.openZip(source, options?)`

打开 ZIP 归档进行读取。TAR 用 `Archive.openTar()`，tar.gz 用 `Archive.openTarGz()`。

```typescript
import { Archive } from "documonster/archive";

const reader = Archive.openZip(zipBytes);
const reader = Archive.openZip(zipBytes, { password: "secret" });
const reader = Archive.openTar(tarBytes);
```

### `ZipReader`

流式 ZIP 读取器，支持随机访问。

```typescript
import { Archive } from "documonster/archive";

const reader = Archive.openZip(zipBytes);

// 流式迭代
for await (const entry of reader.entries()) {
  const data = await entry.bytes();
  const text = await entry.text();
  // 或流式读取条目
  for await (const chunk of entry.stream()) {
    /* process chunk */
  }
  // 或管道
  await entry.pipeTo(writableStream);
}

// 随机访问（流式数据源会先被缓冲）
const entry = await reader.get("path/to/file.txt");
const bytes = await reader.bytes("path/to/file.txt");
const paths = await reader.list(); // 所有条目，按中央目录顺序

// 带进度和中止
const op = reader.operation({
  onProgress: p => console.log(`${p.entriesEmitted} 个条目，${p.bytesOut} 字节`),
  signal: controller.signal
});
```

### `UnzipEntry`

表示归档中的单个条目。

```typescript
entry.path; // "path/to/file.txt"
entry.type; // "file" | "directory" | "symlink"
entry.mode; // Unix 模式（如 0o644）
entry.linkTarget; // 符号链接目标（在 bytes() 之后可用）
entry.info; // 冻结的 ZipEntryRecord：大小、crc32、lastModified、注释……
//            （流式读取时来自本地文件头；见其文档）

await entry.bytes(); // Uint8Array
await entry.text(); // 字符串（UTF-8）
await entry.text("latin1"); // 字符串（自定义编码）
for await (const chunk of entry.stream()) {
  /* process chunk */
}
await entry.pipeTo(writable);
entry.readableStream(); // WHATWG ReadableStream
entry.discard(); // 跳过不读取
```

**`UnzipOptions`：**

| 选项           | 类型                         | 默认值    | 描述                         |
| -------------- | ---------------------------- | --------- | ---------------------------- |
| `password`     | `string \| Uint8Array`       | -         | 解密密码                     |
| `encoding`     | `ZipStringEncoding`          | `"utf-8"` | 条目名编码                   |
| `checkCrc32`   | `boolean`                    | `true`    | 校验每个条目的 CRC-32        |
| `maxEntrySize` | `number`                     | 512 MiB   | `bytes()` 读入内存的最大条目 |
| `maxEntries`   | `number`                     | -         | 归档最多可含的条目数         |
| `signal`       | `AbortSignal`                | -         | 中止信号                     |
| `onProgress`   | `(p: UnzipProgress) => void` | -         | 进度回调                     |

**限制。** 所有读取器使用同样的两个限制:`maxEntrySize`(读入内存的最大条目,超出抛 `ArchiveFileTooLargeError`),以及 ZIP 的 `maxEntries`(`ArchiveLimitError`)——`unpackZip`、`openZip`、`RemoteZipReader` 和 `ArchiveFile` 都一样;TAR 读取器也接受 `maxEntrySize`(默认 100 MiB;`maxFileSize` 是它已弃用的旧名)。读取不是自己写出的归档时请设置它们。每个读取器在调用时检查这些限制:不是从 0 起的整数(或 `Infinity`)的限制会抛 `ArchiveArgumentError`,其字段为 `options.<限制名>`。触及限制时抛 `ArchiveLimitError`,其 `limit` 是选项名,`allowed` 是它的值——`maxEntrySize` 除外:它抛 `ArchiveFileTooLargeError`,带有同样的 `limit`(`"maxEntrySize"`)和 `allowed` 字段,另带条目的 `path`。

**字符串。** 传给读取器的字符串是归档的内容,按 UTF-8 编码——绝不是文件路径或 URL。请先读出文件的字节,或在 Node 中用 `ArchiveFile.fromFile` 打开;ZIP 的二进制内容无法完整保存在 JavaScript 字符串中。

**名字。** 一个 ZIP 中每个名字只有一个条目:`ZipArchive.add` 与 `packZip` 拒绝已加入过的名字(`ArchiveArgumentError`)。其他程序写出的、同名出现两次的归档按解压工具的方式读取——`get`、`bytes`、`unpackZip` 与 `RemoteZipReader.extract` 返回后一个条目,与 TAR 的 `TarReader.get`、`unpackTar` 一致——未做修改的 `ZipEditor.open` 会原样写回每个条目。

**一次性数据源。** `ReadableStream`、Node 流或异步可迭代对象只能读一次。`openZip` / `openTar` 读取这类数据源时,要么由一次 `entries()` 读完,要么在 `list`、`get` 或 `bytes` 第一次需要时缓冲下来(之后的每次调用,包括 `entries()`,都使用缓冲);其他情况抛出 `ArchiveError`,说明流已被读完。`openTarGz` 会保留解压后的归档,可以任意多次读取。

**完整性。** 所有 ZIP 读取都会用归档记录的 CRC-32 校验每个条目，不一致时抛出
`ZipCrc32MismatchError`——包括 `Archive.openZip` / `ZipReader`（内存与流式数据源，`bytes()` 与
`stream()`）、`Archive.unpackZip` / `unpackZipSync`、`RemoteZipReader` 和 `ArchiveFile`。
校验和覆盖整个条目，因此 `stream()` 在产出最后一块之后才报告不一致。传入 `checkCrc32: false`
可以尽量读取已损坏的归档。AES 条目改由其 HMAC 认证，ZipCrypto 条目则始终校验。

校验针对**你读取的条目**，所有数据源一致：不一致会使该次 `bytes()` 或 `stream()` 失败——捕获后
迭代会继续到下一个条目；被跳过（直接越过或 `discard()`）的条目不做校验，也不会导致任何失败。
对流式数据源，读取器在字节经过时计算校验，但只把结果报告给读取该条目的人，与内存中的归档只在
解压条目时才校验完全一致。

**流式数据源。** 流只有在当前条目读完后才能前进到下一个条目，因此读取器会处理你没有读取的条目：

- 未读取就越过的条目，若较小（不超过 256 KiB）会保留在内存中，之后仍可读取；更大的条目会被
  跳过，之后读取会抛出 `ArchiveError`，而不是卡住。被 `discard()` 的条目读取时同样抛出。
- 离开循环——`break`、`return` 或循环体中的 `throw`——会立即结束。离开时持有的条目按同样规则
  仍可读取（因此较大的条目应在离开循环前读取），该条目处理完后数据源即被释放：Node 流被销毁，Web `ReadableStream` 被取消并解锁
  （与 Excel 的流式工作簿读取器一致），解析器及其解压器也随之释放。失败或中止则立即释放一切。

---

## ZIP 编辑器

对现有 ZIP 归档进行文件系统式编辑。未变更的条目高效直通传递（保留原始压缩字节）。

```typescript
import { Archive } from "documonster/archive";

// 从字节
const editor = await Archive.ZipEditor.open(zipBytes, {
  reproducible: true,
  preserve: "best-effort",
  onWarning: w => console.warn(w.code, w.entry, w.message)
});

// 从 URL（HTTP Range 请求）
const editor = await Archive.ZipEditor.openUrl("https://example.com/archive.zip");

// 操作
editor.has("file.txt"); // 检查存在
editor.set("new.txt", "content"); // 添加或替换
editor.delete("old.txt"); // 删除条目
editor.deleteDirectory("old-dir/"); // 递归删除目录
editor.rename("a.txt", "b.txt"); // 重命名条目
editor.setComment("Updated archive"); // 设置归档注释

// 可复用的编辑计划
import { Archive } from "documonster/archive";

const plan = new Archive.ZipEditPlan();
plan.set("config.json", newConfig);
plan.delete("temp/");
editor.apply(plan);

// 输出
const output = await editor.toBuffer();
await editor.pipeTo(writable);
for await (const chunk of editor.stream()) {
  /* process chunk */
}
```

---

## 远程 ZIP 读取器

通过 HTTP Range 请求读取 ZIP 文件 — 只下载你需要的条目。

```typescript
import { Archive } from "documonster/archive";

const reader = await Archive.RemoteZipReader.open("https://example.com/large.zip");

// 元数据（尚未下载内容）
const entries = reader.getEntries();
console.log(`${entries.length} 个条目`);

const files = reader.listFiles(); // 文件条目(不含目录);每个条目一项,与 getFileCount() 一致
const count = reader.getFileCount();

// 按模式查找条目
const jsFiles = reader.findEntries("**/*.js");

// 提取单个文件（只下载该条目）
const data = await reader.extract("path/to/file.txt");

// 提取多个（批量处理以提高效率）
const results = await reader.extractMultiple(["a.txt", "b.txt"]);

// 流式提取
await reader.extractToStream("large-file.bin", writableStream);

// 密码保护的归档
const reader = await Archive.RemoteZipReader.open(url);
const isValid = await reader.checkPassword("secret.txt", "mypassword");
const data = await reader.extract("secret.txt", { password: "mypassword" });

// 统计信息
const stats = reader.getStats();
console.log(`已下载 ${stats.http?.bytesDownloaded} 字节（${stats.http?.downloadedPercent}%）`);

await reader.close();
```

---

## 压缩 API

底层压缩工具。DEFLATE、GZIP 和 Zlib，支持同步、异步和流式变体。

### 一次性压缩

```typescript
import { Compression } from "documonster/archive";

// DEFLATE-RAW（ZIP 文件使用）
const compressed = await Compression.deflate(data, { level: 9 });
const original = await Compression.inflate(compressed);
const compressedSync = Compression.deflateSync(data);

// GZIP
const gzipped = await Compression.gzip(data);
const ungzipped = await Compression.gunzip(gzipped);

// Zlib（RFC 1950）
const zlibbed = await Compression.zlib(data);
const unzlibbed = await Compression.unzlib(zlibbed);

// 自动检测并解压
const format = Compression.detectFormat(data); // "gzip" | "zlib" | "deflate-raw"
const original = await Compression.decompressAuto(data);
```

### 流式压缩

```typescript
import { Compression } from "documonster/archive";

// DEFLATE-RAW 流式
const deflater = Compression.createDeflateStream({ level: 6 });
const inflater = Compression.createInflateStream();

// GZIP 流式
const gzipper = Compression.createGzipStream();
const gunzipper = Compression.createGunzipStream();

// Zlib 流式
const zlibber = Compression.createZlibStream();
const unzlibber = Compression.createUnzlibStream();

// 检查可用性
Compression.hasDeflateRaw(); // Node.js 中为 true，浏览器取决于 CompressionStream
```

### CRC32

```typescript
import { Archive } from "documonster/archive";

// 一次性
const checksum = Archive.crc32(data);

// 增量（流式）
let state = 0xffffffff;
state = Archive.crc32Update(state, chunk1);
state = Archive.crc32Update(state, chunk2);
const checksum = Archive.crc32Finalize(state);
```

### PNG 编码

PNG 就是一个 DEFLATE 流加上带 CRC-32 校验的 chunk —— 正是上面这两样原语 —— 所以编码器住在这里，
而不是住在某个"恰好需要图片"的模块里。它之所以存在，是因为绘图引擎刻意在像素处停下：
`documonster/draw` 的 `Draw.toRgba` 交回 RGBA，而编码需要 DEFLATE，后者位于 `draw` 之上一层。

```typescript
import { Archive } from "documonster/archive";
import { Draw } from "documonster/draw";

// 来自任意显示列表 —— 图表、Mermaid 图、你自己的生产者
const image = Draw.toRgba(list, { scale: 2 });
const png = Archive.encodePng(image.data, image.width, image.height, { dpi: 192 });

// 或者来自你手上已有的像素（8 位 RGBA，行优先）
const bytes = Archive.encodePng(rgba, width, height);
```

输入 8 位 RGBA，filter 0，单个 `IDAT`。给了 `dpi` 会写入 `pHYs` chunk；不给则不写。

---

## TAR 支持

与 ZIP 接口兼容的统一 API。

### 创建 TAR 归档

```typescript
import { Archive } from "documonster/archive";

// 便捷函数
const tarBytes = await Archive.packTar(
  new Map([
    ["file.txt", "content"],
    ["data.bin", uint8Array]
  ]),
  { modTime: new Date() }
);

// 构建器 API（与 ZipArchive 相同）
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

### 读取 TAR 归档

```typescript
import { Archive } from "documonster/archive";

// 通过统一 API
const reader = Archive.openTar(tarBytes);

// 直接使用
const reader = Archive.openTar(tarBytes);
for await (const entry of reader.entries()) {
  console.log(entry.path, entry.isDirectory);
  const data = await entry.bytes();
}

// 随机访问
const data = await reader.bytes("file.txt");
const paths = await reader.list();
```

### TAR + GZIP

```typescript
import { Archive } from "documonster/archive";

// 创建 .tar.gz
const tgzBytes = await Archive.packTarGz(new Map([["file.txt", "content"]]), { level: 9 });

// 构建器 API，使用流式 gzip
const archive = Archive.createTarGz({ level: 6 });
archive.add("file.txt", "content");
const bytes = await archive.toBuffer();
for await (const chunk of archive.stream()) {
  /* process chunk */
}

// 解析 .tar.gz
const entries = await Archive.parseTarGz(tgzBytes);
for (const entry of entries) {
  console.log(entry.info.path, await entry.bytes());
}

// 提取 .tar.gz 到 Map
const files = await Archive.unpackTarGz(tgzBytes);

// 大型 .tar.gz：增量解压并解析，内存中每次只保留一个条目
for await (const entry of Archive.parseTarGzStream(response.body!)) {
  console.log(entry.info.path, (await entry.bytes()).length);
}
```

`parseTarGz`、`unpackTarGz`、`unpackTarGzSync` 和 `Archive.openTarGz`
会在内存中解压整个归档（读取器会保留结果，因此即使是一次性流也可以反复迭代和查找）。
`Archive.openTarGz` 的并发调用共享一次解压，每个调用的 `signal` 只取消该调用本身——唯一例外是一次性流：
若在解压完成前等待读取的所有调用都已中止（包括 `close()` 之后），流会被取消，之后的调用抛出
`ArchiveError`（而不是 `AbortError`），因为它无法再次读取。已完成的读取永远不会被丢弃——此时中止只会拒绝该调用本身。内存中的数据源从不会被取消。一次性流的读取失败是终局的（之后的调用重新抛出同一错误）；
内存中的数据源会重试。
`parseTarGzStream`（与纯 TAR 的 `parseTarStream` 一样）不会，它每次只在内存中保留一个条目。它的错误在迭代途中（已产出前面的条目之后）才抛出。
所有 TAR 和 tar.gz 读取函数都会以 `ArchiveError` 拒绝在条目或头部内部被截断的归档，而不是返回不完整的条目。TAR 结束后它仍会读完剩余的 gzip 流，
因此被截断或损坏的 gzip 尾部同样会被报告。所有 tar.gz 读取函数都把损坏的 gzip 流报告为
`ArchiveError`（解码器自身的错误作为其 `cause`）。

在 Node 中，无论用哪种方式构建——`packTarGz`、`packTarGzSync`、`createTarGz().toBuffer()`、
`toBufferSync()` 或 `stream()`——得到的 `.tar.gz` 都逐字节相同。在浏览器中这只在 `level: 0`
时成立；其他级别下 `toBuffer()` 和 `stream()` 使用的原生 `CompressionStream` 是另一种编码器
（且不接受级别），这些输出解压后是相同的 TAR，但字节并不相同。

对内存中的数据源，
`unpackTar` / `unpackTarSync` / `unpackTarGz*` 返回的 `data` 是该缓冲区的**视图**（`subarray`），
而不是副本——修改任一方之前请先 `slice()`。

---

## 文件系统集成（Node.js）

高级 `ArchiveFile` 类，支持磁盘 I/O、glob 模式、目录遍历和提取。

```typescript
import { Archive } from "documonster/archive";

// 从零创建
const af = new Archive.ArchiveFile();
af.addFile("./src/index.ts");
af.addText("Hello", "hello.txt");
af.addBuffer(bytes, "data.bin");
af.addDirectory("./src", { recursive: true });
af.addGlob("**/*.ts", { cwd: "./src" });
await af.writeFile("output.zip");

// 打开现有归档
const af = await Archive.ArchiveFile.fromFile("archive.zip");
const entries = await af.getEntries();
const data = await af.readEntry("file.txt");
const text = await af.readAsText("file.txt");

// 提取
await af.extractTo("./output", {
  overwrite: "newer",
  filter: (entryPath, isDirectory) => isDirectory || entryPath.endsWith(".txt")
});

// 编辑（仅 ZIP）
af.set("new.txt", "content");
af.delete("old.txt");
af.rename("a.txt", "b.txt");
await af.writeFile("modified.zip");

// 流式输出
for await (const chunk of af.stream()) {
  /* process chunk */
}
await af.streamToFile("output.zip");

// TAR 格式
const tar = new Archive.ArchiveFile({ format: "tar" });
tar.addFile("./data.txt");
tar.addSymlink("latest.txt", "data.txt", { mode: 0o777 });
await tar.writeFile("archive.tar");
```

编辑在写出前一直挂起，而每个读取（`has`、`entryCount`、`getEntryNames`、`getEntries`、`readEntry`、`extractTo` 等）
回答的都是此刻写出会得到的归档：打开的归档的条目应用了编辑（添加、替换、删除、重命名、目录）之后的结果，每个条目的描述与写出时记录的一致。
每次写出——无论打开的是 ZIP 还是 TAR——写出的就是这个归档，之后它成为下一次写出的起点，因此写出两次得到的是同一个归档。
列出名称不读取任何条目的数据；挂起的条目在读取需要时从内存读取，`addFile` 的条目则从磁盘读取。挂起的流只能消费一次，
因此第一个需要它的读取或写出会把它的字节保存在内存中直到归档写出——内存开销为流的完整大小；同步读取无法消费流，
在异步读取缓冲它之前会抛出错误。按路径打开的 `.tgz` 写回时仍是 gzip 压缩。`writeFile` 默认拒绝覆盖已存在的文件，
除非 `overwrite` 另有指定，TAR 与 ZIP 相同。

其底层的文件系统辅助函数 — `Fs.glob`、`Fs.traverseDirectory`、`Fs.ensureDir`、
`Fs.readFileBytes` 等 — 发布在仅 Node.js 可用的 `Fs` 命名空间上。参数错误抛出 `ArchiveArgumentError`，
指明调用与参数名；文件系统随后的拒绝（`ENOENT`、`EACCES` 等）按平台错误原样传出。文件流请使用
`node:fs` 的 `createReadStream` / `createWriteStream`:

```typescript
import { Fs } from "documonster/archive";

for await (const entry of Fs.glob("**/*.json", { cwd: "./config" })) {
  console.log(entry.relativePath);
}
await Fs.ensureDir("./out");
```

---

## 加密

支持 ZIP 传统加密（旧版）和 AES-256 加密。

```typescript
import { Archive } from "documonster/archive";

// 加密条目（读取时自动处理）
const reader = Archive.openZip(encryptedZip, { password: "secret" });
for await (const entry of reader.entries()) {
  const data = await entry.bytes(); // 自动解密
}

// 远程加密 ZIP
const remote = await Archive.RemoteZipReader.open(url);
const data = await remote.extract("secret.txt", { password: "mypassword" });

// 无需完整提取即可验证密码
const valid = await remote.checkPassword("file.txt", "mypassword");
```

---

## 错误类

```typescript
import {
  ArchiveError,
  ArchiveArgumentError, // 参数或选项错误；`.field` 指出是哪一个
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
  ArchiveLimitError,
  ArchiveWorkerStreamStartError,
  ArchiveWorkerUnsupportedError, // 在 Node 上调用 Compression.Workers;`.operation` 指出是哪一个
  ArchiveFileError, // ArchiveFile.fromFile / writeFile 被文件系统拒绝;`.cause` 是系统错误
  AbortError,
  isArchiveError // Type guard: (err: unknown) => err is ArchiveError
} from "documonster/archive";
```

所有归档错误都继承自 `ArchiveError`，因此 `isArchiveError` 可以识别其中任何一个。错误的参数或选项抛出
`ArchiveArgumentError`，其 `field` 是从参数名起的路径（`"options.level"`、`"source"`、`"entries[0]"` 等），消息以调用名开头（`Archive.packZip: …`）；归档数据本身的损坏仍是普通的
`ArchiveError` 或其数据类子类。被取消的操作抛出的是共享的
`AbortError`，它不是 `ArchiveError`（与 `documonster/stream` 导出的是同一个类）：

```typescript
import { Archive, isArchiveError, AbortError } from "documonster/archive";

try {
  await Archive.unpackZip(bytes);
} catch (err) {
  if (err instanceof AbortError) {
    // 通过 AbortSignal 取消
  } else if (isArchiveError(err)) {
    console.error(err.name, err.message, err.cause);
  }
}
```

---

## API 参考

### 高级函数

| 函数                                       | 描述                      |
| ------------------------------------------ | ------------------------- |
| `Archive.createZip(options?)`              | 创建 ZIP 构建器           |
| `Archive.createTar(options?)`              | 创建 TAR 构建器           |
| `Archive.createTarGz(options?)`            | 创建 tar.gz 构建器        |
| `Archive.openZip(source, options?)`        | 打开 ZIP 进行读取         |
| `Archive.openTar(source, options?)`        | 打开 TAR 进行读取         |
| `Archive.openTarGz(source, options?)`      | 打开 tar.gz 进行读取      |
| `Archive.ZipEditor.open(source, options?)` | 打开 ZIP 进行编辑         |
| `Archive.ZipEditor.openUrl(url, options?)` | 打开远程 ZIP 进行编辑     |
| `Archive.packZip(entries, options?)`       | 从条目创建 ZIP（异步）    |
| `Archive.packZipSync(entries, options?)`   | 从条目创建 ZIP（同步）    |
| `Archive.packTar(entries, options?)`       | 从条目创建 TAR（异步）    |
| `Archive.packTarSync(entries, options?)`   | 从条目创建 TAR（同步）    |
| `Archive.packTarGz(entries, options?)`     | 从条目创建 tar.gz（异步） |
| `Archive.packTarGzSync(entries, options?)` | 从条目创建 tar.gz（同步） |
| `Archive.unpackZip(source, options?)`      | ZIP → 路径 Map（异步）    |
| `Archive.unpackZipSync(bytes, options?)`   | ZIP → 路径 Map（同步）    |
| `Archive.unpackTar(source, options?)`      | TAR → 路径 Map（异步）    |
| `Archive.unpackTarSync(bytes, options?)`   | TAR → 路径 Map（同步）    |
| `Archive.unpackTarGz(source, options?)`    | tar.gz → 路径 Map（异步） |
| `Archive.unpackTarGzSync(bytes, options?)` | tar.gz → 路径 Map（同步） |

### 类

| 类                        | 描述                                                                                                                 |
| ------------------------- | -------------------------------------------------------------------------------------------------------------------- |
| `Archive.ZipArchive`      | 流式 ZIP 构建器                                                                                                      |
| `Archive.ZipReader`       | 流式 ZIP 读取器                                                                                                      |
| `Archive.UnzipEntry`      | 单个归档条目                                                                                                         |
| `Archive.ZipEditor`       | 编辑现有 ZIP 归档                                                                                                    |
| `Archive.TarArchive`      | 流式 TAR 构建器（`Archive.createTar()`）                                                                             |
| `Archive.TarReader`       | 流式 TAR 读取器（`Archive.openTar(source)`）                                                                         |
| `Archive.TarGzArchive`    | TAR + GZIP 构建器（`Archive.createTarGz()`）                                                                         |
| `Archive.RemoteZipReader` | 随机访问 ZIP 读取器（`RemoteZipReader.open(url)`、`RemoteZipReader.fromReader(reader)`）                             |
| `Archive.HttpRangeReader` | HTTP Range 随机访问读取器（`HttpRangeReader.open(url)`）                                                             |
| `Archive.ArchiveFile`     | 磁盘归档（Node.js；`new ArchiveFile(options)`、`ArchiveFile.fromFile(path)` / `fromFileSync` / `fromBuffer(bytes)`） |

### 压缩

| 函数                                                                   | 描述                    |
| ---------------------------------------------------------------------- | ----------------------- |
| `Compression.deflate` / `deflateSync`                                  | DEFLATE-RAW 压缩        |
| `Compression.inflate` / `inflateSync`                                  | DEFLATE-RAW 解压        |
| `Compression.gzip` / `gzipSync`                                        | GZIP 压缩               |
| `Compression.gunzip` / `gunzipSync`                                    | GZIP 解压               |
| `Compression.zlib` / `zlibSync`                                        | Zlib 压缩               |
| `Compression.unzlib` / `unzlibSync`                                    | Zlib 解压               |
| `Compression.decompressAuto`                                           | 自动检测并解压          |
| `Compression.createDeflateStream`                                      | 流式 DEFLATE 压缩器     |
| `Compression.createInflateStream`                                      | 流式 DEFLATE 解压器     |
| `Compression.createGzipStream`                                         | 流式 GZIP 压缩器        |
| `Compression.createGunzipStream`                                       | 流式 GZIP 解压器        |
| `Archive.crc32`                                                        | CRC32 校验和            |
| `Archive.encodePng`                                                    | RGBA 像素转 PNG 字节    |
| `Compression.Workers.deflate` / `inflate`（及 `…Batch`、`WorkerPool`） | Web Worker 池（浏览器） |

底层 ZipCrypto / AES 原语及其常量（`encryptZipCrypto`、`decryptAes`、`AES_KEY_LENGTH` 等）位于
`Archive.Crypto`；普通的加密读写只需要 `password` 选项。
