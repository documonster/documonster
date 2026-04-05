# Archive 模块（ZIP / TAR）

[English](README.md)

零依赖、跨平台的归档工具包，用于创建、读取和编辑 ZIP 与 TAR 归档。

## 功能特性

- **ZIP 和 TAR** — 创建、读取、流式处理
- **ZIP 编辑** — `ZipEditor` 支持 set、delete、rename、deleteDirectory
- **远程 ZIP** — 通过 HTTP Range 请求读取远程 ZIP 文件
- **流式处理** — 支持背压的异步流
- **加密** — 传统 ZipCrypto 和 AES-256
- **ZIP64** — 大文件支持（>4GB）
- **压缩** — DEFLATE、GZIP、Zlib（同步和异步）
- **进度和中止** — 进度回调和 AbortSignal 支持
- **可重现** — 确定性输出（`reproducible: true`）
- **智能存储** — 自动检测不可压缩数据
- **文件系统集成** — `ArchiveFile` 高级文件操作

## 快速开始

```typescript
import { zip, unzip, ZipArchive } from "@cj-tech-master/excelts/zip";

// 创建 ZIP
const archive = zip();
archive.add("hello.txt", "Hello, World!");
archive.add("data.json", JSON.stringify({ key: "value" }));
const bytes = await archive.bytes();

// 读取 ZIP
const reader = await unzip(bytes);
for await (const entry of reader.entries()) {
  console.log(entry.path, await entry.text());
}

// 编辑 ZIP
import { ZipEditor } from "@cj-tech-master/excelts/zip";
const editor = await ZipEditor.open(bytes);
editor.set("new.txt", "新文件");
editor.delete("old.txt");
const edited = await editor.bytes();

// 压缩工具
import { compress, decompress, gzip, gunzip } from "@cj-tech-master/excelts/zip";
const compressed = await compress(data, { level: 6 });
const decompressed = await decompress(compressed);
```

## 示例

查看 [examples 目录](examples/) 获取可运行代码。
