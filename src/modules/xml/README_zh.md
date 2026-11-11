# XML 模块

[English](README.md)

零依赖、跨平台的 XML 工具包，支持流式和缓冲两种模式的读写。

```typescript
import { Xml } from "documonster/xml";
```

## 功能特性

- **零依赖** — 纯 TypeScript，无原生插件
- **跨平台** — Node.js 和浏览器使用相同 API
- **双模式写入** — 缓冲式（`Xml.Writer`）和流式（`Xml.StreamWriter`）
- **双模式读取** — SAX 流式（`Xml.SaxParser`）和 DOM 树（`Xml.parse`）
- **共享接口** — `XmlSink` 让渲染代码可以透明地同时适配两种写入模式
- **XML 编码** — 快速实体编码/解码，支持特殊字符处理
- **命名空间支持** — 完整的 XML 命名空间：前缀解析、保留命名空间强制、未绑定前缀检测
- **查询引擎** — 简化路径表达式，用于查询 DOM 树
- **安全加固** — 实体扩展限制、嵌套深度限制、重复属性拒绝、名称注入防护、BOM 处理

---

## 快速开始

### 写入 XML（缓冲模式）

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

### 写入 XML（流式模式）

```text
import { Xml } from "documonster/xml";

const chunks: string[] = [];
const target = { write: (chunk: string) => chunks.push(chunk) };
const sw = new Xml.StreamWriter(target);

sw.openXml();
sw.openNode("root");
sw.leafNode("item", { id: "1" }, "hello");
sw.closeNode();
// 直接写入 target；只暂存后面还没有内容的开始标签
```

### 读取 XML（SAX — 流式）

```text
import { Xml } from "documonster/xml";

const parser = new Xml.SaxParser();
parser.on("opentag", tag => console.log("open:", tag.name, tag.attributes));
parser.on("text", text => console.log("text:", text));
parser.on("closetag", tag => console.log("close:", tag.name));
// 每个事件只有一个处理函数：再次调用 `on("text", …)` 会替换前一个。
parser.write('<root><item id="1">hello</item></root>');
parser.close();
```

### 读取 XML（DOM — 缓冲模式）

```text
import { Xml } from "documonster/xml";

const doc = Xml.parse('<root><item id="1">hello</item></root>');
const item = Xml.findChild(doc.root, "item");
console.log(Xml.getAttr(item!, "id")); // "1"
console.log(Xml.getText(item!)); // "hello"
```

### XML 转换为普通对象

将 XML 转换为普通 JavaScript 对象。

两种入口适用于不同场景：

```text
import { Xml } from "documonster/xml";

// 方式 1：已有 DOM 树
const doc = Xml.parse('<root attr="1"><item>a</item><item>b</item></root>');
const obj = Xml.toPlainObject(doc.root);
// { root: { "@_attr": "1", item: ["a", "b"] } }

// 方式 2：XML 字符串直接转换为普通对象（更快，单次 SAX 遍历）
const obj2 = Xml.parseToObject('<root attr="1"><item>a</item><item>b</item></root>');
// 输出相同，中大型 XML 约快 1.6 倍
```

**如何选择：**

- `Xml.toPlainObject(element)` — 当你已经通过 `Xml.parse()` 获得了 `XmlElement`
- `Xml.parseToObject(xml)` — 当你只需要普通对象（跳过 DOM 分配）

**默认转换规则：**

- 属性以 `@_` 为前缀
- 重复的同级元素合并为数组
- 纯文本元素折叠为字符串值
- 空元素变为 `""`
- 仅含空白的缩进文本默认丢弃

**局限性：** 普通对象转换是有损的 — 不保留元素顺序、注释或处理指令。如需精确的 XML 结构，请使用 `Xml.parse()` 并直接操作 DOM 树。

### 查询引擎

```typescript
import { Xml } from "documonster/xml";

const doc = Xml.parse("<root><a><b>1</b><b>2</b></a><a><b>3</b></a></root>");
const first = Xml.query(doc.root, "a/b"); // 第一个 <b> 元素
const all = Xml.queryAll(doc.root, "a/b"); // 所有 <b> 元素
const indexed = Xml.queryAll(doc.root, "a/b[0]"); // 每个 <a> 下的第一个 <b>
const filtered = Xml.query(doc.root, "a/b[@id='x']"); // id="x" 的 <b>
const deep = Xml.queryAll(doc.root, "a//b"); // <a> 下任意深度的 <b>
```

### 编码/解码

```typescript
import { Xml } from "documonster/xml";

Xml.encode('<tag attr="val">'); // "&lt;tag attr=&quot;val&quot;&gt;"
Xml.decode("&lt;hello&gt;"); // "<hello>"
```

---

## 架构

```
src/modules/xml/
├── types.ts              # 核心类型（XmlNode、XmlSink、XmlSaxTag 等）
├── errors.ts             # XmlError、XmlParseError、XmlWriteError
├── encode.ts             # Xml.encode、Xml.decode、Xml.validateName、Xml.encodeCData 等
├── writer.ts             # Xml.Writer（缓冲式，支持回滚）
├── stream-writer.ts      # Xml.StreamWriter（流式，写入 XmlWritableTarget）
├── sax.ts                # Xml.SaxParser（事件驱动）+ Xml.parseSaxStream（异步生成器）
├── dom.ts                # Xml.parse + DOM 查询辅助函数 + Xml.toPlainObject
├── to-object.ts          # Xml.parseToObject（SAX 直接转换，单次遍历）
├── to-object-shared.ts   # 共享转换逻辑（内部使用）
├── query.ts              # 简化路径查询引擎
├── index.ts              # 公开 API barrel
└── __tests__/            # 测试
```

### 写入路径

```
XmlSink（接口）
├── Xml.Writer        — 在内存中构建 XML 字符串
│                      支持回滚/事务（snapshot/commit/rollback）
│                      适用于：中小型 XML、试探性写入
│
└── Xml.StreamWriter  — 直接写入 XmlWritableTarget
                       不持有整篇文档：内存随嵌套深度和正在写的开始标签增长，
                       与文件大小无关
                       适用于：大型 XML（10万+ 行的工作表）
```

### 读取路径

```
Xml.SaxParser  — 事件驱动的流式解析器
│                      通过 write() 输入数据块，事件同步触发
│                      适用于：大型 XML、只需特定元素时
│
├── Xml.parse        — 构建 DOM 树（XmlDocument/XmlElement）
│   │                  基于 Xml.SaxParser 构建 — 无重复解析逻辑
│   │                  适用于：中小型 XML、需要树遍历时
│   │
│   └── Xml.toPlainObject — 将 XmlElement DOM 转换为普通 JS 对象
│                       适用于：已有 DOM 树时
│
├── Xml.parseToObject — SAX 直接转换为普通 JS 对象（单次遍历，无 DOM）
│                      比 Xml.parse + Xml.toPlainObject 快约 1.6 倍
│                      适用于：XML 字符串 → 普通对象 → JSON.stringify
│
└── Xml.parseSaxStream     — 包装 Xml.SaxParser 的异步生成器，用于流式迭代
                       适用于：异步管道（如从 zip 流中读取）
```

---

## API 参考

### Xml.Writer

由 `new Xml.Writer()` 创建。

| 方法                            | 描述                            |
| ------------------------------- | ------------------------------- |
| `openXml(attrs?)`               | 写入 `<?xml ...?>` 声明(须最先) |
| `openNode(name, attrs?)`        | 打开一个元素                    |
| `closeNode()`                   | 关闭最近打开的元素              |
| `leafNode(name, attrs?, text?)` | 一次调用写入一个完整元素        |
| `addAttribute(name, value)`     | 为当前打开的元素添加属性        |
| `addAttributes(attrs)`          | 添加多个属性                    |
| `writeText(text)`               | 写入转义后的文本内容            |
| `writeRaw(xml)`                 | 写入预转义的 XML(字符串)        |
| `writeCData(text)`              | 写入 `<![CDATA[...]]>` 节       |
| `writeComment(text)`            | 写入 `<!--...-->` 注释          |
| `closeAll()`                    | 关闭所有打开的元素              |
| `toString()`                    | 获取构建的 XML 字符串           |
| `snapshot()`                    | 保存回滚快照                    |
| `commit()`                      | 丢弃快照（保留更改）            |
| `rollback()`                    | 恢复到快照（丢弃更改）          |
| `reset()`                       | 清除所有内容                    |
| `depth`                         | 当前嵌套深度                    |
| `currentElement`                | 最内层打开的元素名              |
| `cursor`                        | 随写入增长,回滚时回退           |

### Xml.StreamWriter

由 `new Xml.StreamWriter(target)` 创建。与 `Xml.Writer` 方法相同（都实现 `XmlSink`），区别在于：

- 无 `toString()` — 内容已写入目标
- 无 `snapshot()` / `commit()` / `rollback()` — 流式写入不可逆
- 无 `cursor` — 流式模式不适用
- 无 `reset()` — 已写入目标的内容无法收回

它不保留文档副本，只保留打开的元素和后面还没有内容的开始标签（最内层的仍可追加属性），因此内存为
O(嵌套深度 + 待写开始标签)，不随文件增长。待写的开始标签会和紧随其后的内容一起，在一次 `write()` 调用中写出。

两种写入器在已写出任何内容后都拒绝 `openXml`（`XmlWriteError`）：XML 只允许声明出现在文档最开头。解析器同样拒绝位置不对的 `<?xml …?>`。（本包的 Excel 和 Word 读取器从包内部件的第一个 `<` 开始解析，所以生成方在部件声明之前写的空白在那里不算错误——这是这些读取器的策略，不是 XML 解析器的。）

两种写入器只写属性映射自身的属性——继承来的不写；并拒绝任何文档都无法容纳的命名空间声明（`xmlns:xml` 绑定到别处、`xmlns:xmlns`、前缀绑定到 `""`、保留命名空间绑定到其他前缀）。写入器看不到绑定的前缀不受限制，因为写出的内容可能嵌入在声明了它的文档中。目标的 `write()` 不能回头调用它所属的 `Xml.StreamWriter`（`XmlWriteError`）。

### Xml.SaxParser

| 方法 / 属性             | 描述                                                                     |
| ----------------------- | ------------------------------------------------------------------------ |
| `write(chunk)`          | 输入 XML 文本，或以 `Uint8Array` 给出的文档字节（UTF-8，可在任意处切分） |
| `close()`               | 信号输入结束                                                             |
| `stop()`                | 在处理函数中调用：不再读取任何内容，包括当前数据块的剩余部分             |
| `on(event, handler)`    | 注册事件处理器                                                           |
| `off(event)`            | 移除事件处理器                                                           |
| `line` / `column`       | 当前位置（启用位置追踪时）                                               |
| `startByte` / `endByte` | 在处理函数中：当前事件在文件中的字节范围（输入为字节时）                 |
| `byte`                  | 下一个待解析字节的位置（输入为字节时）                                   |
| `getLexicalState()`     | 一个可直接比较的值，表示解析器在语法中的位置（见下文）                   |
| `getPrologContext()`    | 根元素名和内部子集声明的实体，读完根开始标签后可用                       |
| `closed` / `stopped`    | 解析器是否已关闭 / 已停止                                                |

**事件：** `opentag`、`closetag`、`text`、`cdata`、`comment`、`pi`、`error`

DOCTYPE 内部子集中的处理指令同样会报告（`Xml.parse` 不把它放进 `prologue`，那里只有文档本身的节点）。处理函数调用 `stop()` 之后不再报告任何内容，包括收尾时发现的错误。处理函数不能对正在调用它的解析器调用 `write()` 或 `close()`：这会被拒绝并抛出 `XmlError`。

**选项：**

| 选项                       | 默认值     | 描述                                                                         |
| -------------------------- | ---------- | ---------------------------------------------------------------------------- |
| `position`                 | `true`     | 追踪行/列位置用于错误信息                                                    |
| `fragment`                 | `false`    | 允许多个根元素                                                               |
| `xmlns`                    | `false`    | 启用命名空间处理                                                             |
| `maxDepth`                 | `256`      | 最大元素嵌套深度                                                             |
| `maxEntityExpansions`      | `10000`    | 最大实体扩展次数（XML 炸弹防御）                                             |
| `maxEntityExpansionLength` | `10000000` | 实体展开文本的最大总长度（UTF-16 单位）                                      |
| `invalidCharHandling`      | `"error"`  | 如何处理无效 XML 字符（见下文）                                              |
| `fileName`                 | —          | 用于错误信息的文件名                                                         |
| `byteOffset`               | `0`        | 第一个输入字节在文件中的位置                                                 |
| `resume`                   | `false`    | 从根元素内部、深度未知处开始读                                               |
| `internalEntities`         | `false`    | 读取并展开内部子集声明的通用实体                                             |
| `maxContextBytes`          | `65536`    | 文件开头上下文（其 JSON 的 UTF-8 字节数）和单条标记声明（UTF-16 码元）的上限 |
| `context`                  | —          | 文件开头上下文，供 `resume` 解析器使用                                       |
| `skipLeadingText`          | `false`    | 跳过文档第一个 `<` 之前的内容（见下文）                                      |

**字节输入。** 传入 `Uint8Array` 时，解析器自行解码 UTF-8——跨数据块的字符会被完整读出——并给每个事件
附上字节范围（`startByte` / `endByte`），给每个错误附上 `byte`。不属于合法 UTF-8 序列的字节是无效字符，按
`invalidCharHandling` 处理，因此字节位置不会漂移。声明 UTF-8（或 ASCII）以外的编码会被拒绝；声明 ASCII 之后，
大于 `0x7F` 的字节不是该编码的字符，按同样方式处理。一个文档只能全部传文本或全部传字节。

**从中途开始读。** `resume: true` 假设起点位于某个元素的内容中：起点之前打开的元素的结束标签会被跳过，不要求
有根元素，起点落在多字节字符中间时丢弃该字符的剩余字节。起点之后打开、到结尾仍未闭合的元素会报错；提供
`context` 时，读到过标签却在根元素结束标签之前就结束，也会报错。

`getLexicalState()` 用来判断读取**同一份输入**的两个解析器（一个从头读，一个从后面某处起读）何时在同一位置
到达同一状态；从那以后两者报告相同的事件。它包含扫描状态、未闭合的引号或读到一半的引用、是否已越过文件开头
部分，以及当前标记或文本的起点；不包含此后已积累的内容（对两者来说是同一段输入），也不包含外层打开的元素。
读取不同输入的解析器可能有相同的词法状态。

### Xml.parseSaxStream（异步生成器）

```text
Xml.parseSaxStream(
  iterable: AsyncIterable<string | Uint8Array> | Iterable<string | Uint8Array>,
  options?: XmlSaxOptions
): AsyncGenerator<XmlSaxEvent[]>
```

每批是一个数据块产生的事件，类型为可判别联合 `XmlSaxEvent`——按 `eventType` 收窄。`Xml.feedSax(parser, source)`
接受同样的数据块，供用 `.on(...)` 注册处理器的解析器使用。`Uint8Array` 数据块以字节形式交给 `Xml.SaxParser`，
因此上文 **字节输入** 的说明都适用：`byteOffset`、字节位置，以及 `invalidCharHandling` 对非法 UTF-8 字节生效。
同一个文档只能全部传文本块或全部传字节块，不能混用。不接受 `ArrayBuffer`，请先包成 `Uint8Array`。

### Xml.readRecords / Xml.readRoot

只从文件字节中读取需要的部分。数据源是一个 `Uint8Array`，或由它们组成的可迭代 / 异步可迭代对象，可在任意处切分。

```ts
import { Xml } from "documonster/xml";

// 读根元素和文件开头的上下文，读到根元素的 `>` 就停止。
const root = await Xml.readRoot(firstChunks);
root.name; // "DSR"

// 一遍扫描，把每个 <Track> 读成对象，任何时候只保留一条记录。
await Xml.readRecords(source, {
  tags: ["Track"],
  onRecord: record => {
    record.object; // { Track: { "@_id": "1", Title: "…" } }，和 Xml.parseToObject 读它的文本一致（见下文）
    record.startByte; // 开始标签 `<` 的位置；endByte 是结束标签之后的位置
  }
});
```

记录是名字在 `tags` 中、且不在另一条记录内部的元素；记录内部的同名元素属于这条记录。`onRecord` 按文件顺序
调用，返回的 Promise 会被等待。设 `rawXml: true` 时还会给出每条记录的原文。`record.object` 与 `Xml.parseToObject` 解析该文本的结果一致，
前提是实体上下文相同：用到文件 DOCTYPE 所声明实体的记录已展开这些实体，单独解析 `record.xml` 时需要等价的
实体声明（以及 `internalEntities: true`）才能得到相同对象。

**按字节范围读。** 给出 `offset` 和 `size` 时，只读开始标签起点落在 `[offset, offset + size)` 内的记录——
因此一个文件可以切成多个范围，由不同的 worker 独立读取，按顺序拼起来就是整个文件的记录。`source` 提供从
`offset` 开始的字节；`offset` 不为 0 时不会读文件开头，需要传入 `context: root.context`。

从自己的起点读一个范围时，会假设起点位于元素内容中；这个假设会被验证而不是被信任：每个范围的读取还会从
下一个范围的起点再读一遍，直到两次读取到达同一状态。若无法做到——范围起点落在一条包含同名子元素的记录内，
或落在满是标签的注释或 CDATA 中——就抛出 `XmlChunkAmbiguityError`（带 `chunkStart`、`atByte`），此时请把这个
范围和前一个合并读取。第一个之后的范围如果起点在根元素开始标签之前（prolog 中，或 `skipLeadingText` 跳过的文本中），
也可能这样被拒绝，因为它会把 XML 声明或 DOCTYPE 当作内容来读；第一个范围无论在哪里结束都会把 prolog 读完，所以只要
范围比 prolog 长，其他范围就不会从那里开始。因此，**只要切分的每个范围都读取成功、且 `Xml.verifyPartition` 接受这些结果，它们的记录就恰好等于整个文件的记录。** 每次读取
只会越过范围末尾，读完最后一条记录并确定下一个范围的起点——在常见文件上大约 1 KB。

**只整批发布。** 一个范围读取成功，不代表其他范围也可信：每个范围读完自己的记录就停止，看不到文件在更后面
被截断，也看不到后面的范围无法确定起点。因此每个结果都带有 `structure`——该范围所见内容的结构摘要，内容不透明，
但可 JSON 序列化，可由 worker 传回。所有范围都成功后，把结果按 `offset` 顺序传给 `Xml.verifyPartition`：它检查
单个范围无法检查的内容——跨范围边界的标签是否匹配、根元素是否闭合（文件是否被截断）、是否只有一个根、根元素外
是否只有空白、注释和处理指令——不满足就抛错。

应对同一个固定版本的文件运行全部范围，把各自的记录先暂存。只有全部范围成功 **且** `Xml.verifyPartition` 通过
才发布；任何一步失败，就丢弃整批结果（包括已成功范围的结果），再重新读取（把失败的范围和前一个合并，或整读
文件）。`onRecord` 在读到记录时就会调用，因此失败的调用可能已经交出了部分记录。

```ts
import { Xml } from "documonster/xml";
import type { XmlReadRecordsResult } from "documonster/xml";

// readRangeIntoStaging 返回该范围的 Xml.readRecords 结果；ranges 按 offset 排序。
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
  await discardStaged(); // 丢弃所有范围的结果，包括已成功的
}
```

出错一律抛出，不会静默恢复。输入在记录内部结束时抛出带 `recordStartByte` 的 `XmlTruncatedError`；读到输入末尾
的范围会检查文档是否在此正常结束。从 offset 0 开始的范围无论在哪里结束，都会把 prolog 一直读到根元素的开始标签，
所以 prolog 会被完整检查。`Xml.verifyPartition` 在文件于某个元素内部结束时抛出 `XmlTruncatedError`，在各范围
拼不成一个文档时抛出 `XmlParseError`；如果某个范围从根元素之外的标记内部开始（根元素的结束标签、尾部注释），
把标记的尾部读成了无法判断的文本，则抛出 `XmlChunkAmbiguityError`：把这个范围和前一个范围合并重读。

`maxDepth`、`maxEntityExpansions` 和 `maxEntityExpansionLength` 限制的是单次调用消耗的资源，所以按单次调用计数：一个范围不知道起点外面有
几层元素，也不知道之前的范围展开过多少次实体。因此整读会超限的文件，分块读时可能不报错——但读出的记录相同，
因为上限只决定是否停止，从不影响读出的内容。

`Xml.readRoot` 读到根元素的 `>` 即停止并关闭数据源。根元素之前的任何错误都会带字节位置抛出；`maxStartByte`
（默认 1,000,000）和 `maxTagBytes`（默认 65,536）按字节位置限制根元素的起点和开始标签的长度——无论数据源怎样
分块，结果都相同。它会展开根元素属性中内部子集声明的实体，受 `maxEntityExpansions`（默认 10,000）和
`maxEntityExpansionLength`（默认 10,000,000 个 UTF-16 单位）限制，含义与同名解析器选项相同。

**文档之前的文本。** 有些文件在 XML 之前带有任何 XML 文档都不会有的文本，例如从 `unzip -c` 的输出保存下来的文件，
开头是 `Archive:  …` 和 `  inflating: …` 两行。设置 `skipLeadingText: true` 后，`Xml.SaxParser`、`Xml.readRoot`
和 `Xml.readRecords` 把第一个 `<` 作为文档的开始：之前的内容被跳过，不作为文本报告，其后可以是 XML 声明。字节位置
仍是原文件中的位置，结果中报告跳过了多少（`skippedLeadingBytes`；解析器上为 `skippedLeadingText`）。其余规则不变：
这段文本中的 `<` 就是文档的开始，根元素之后的文本仍会被拒绝。`readRoot` 和分区的每个范围都应传入同样的设置。默认关闭，
因为这样的文件不是 XML。

两个函数对读到的内容抛出本模块的错误。有两类失败不是 `XmlError`：`signal` 取消时抛出包内的 `AbortError`
（`isAbortError` 为 true，`isXmlError` 为 false）；`source` 自身或 `onRecord`、`isArray` 等回调抛出的异常原样传出。

### Xml.parse

```text
Xml.parse(xml: string, options?: XmlParseOptions): XmlDocument;
```

**选项：**

| 选项                       | 默认值     | 描述                                      |
| -------------------------- | ---------- | ----------------------------------------- |
| `comments`                 | `false`    | 在 DOM 树中包含注释节点                   |
| `processingInstructions`   | `false`    | 在 DOM 树中包含处理指令节点               |
| `cdataAsNodes`             | `false`    | 保留 CDATA 为显式节点（而非合并到文本中） |
| `fragment`                 | `false`    | 允许多个根元素                            |
| `xmlns`                    | `false`    | 启用命名空间处理                          |
| `maxDepth`                 | `256`      | 最大元素嵌套深度                          |
| `maxEntityExpansions`      | `10000`    | 最大实体扩展次数                          |
| `maxEntityExpansionLength` | `10000000` | 实体展开文本的最大总长度                  |
| `invalidCharHandling`      | `"error"`  | 如何处理无效 XML 字符（见下文）           |
| `internalEntities`         | `false`    | 读取并展开内部子集声明的实体              |

**返回值：** `XmlDocument`，包含：

| 字段          | 类型                                            | 描述                                          |
| ------------- | ----------------------------------------------- | --------------------------------------------- |
| `root`        | `XmlElement`                                    | 第一个（或唯一的）根元素                      |
| `roots`       | `XmlElement[]`                                  | 所有根级元素（fragment 模式下有用）           |
| `declaration` | `Record<string, string> \| undefined`           | XML 声明属性（version、encoding、standalone） |
| `prologue`    | `Array<XmlComment \| XmlProcessingInstruction>` | 根元素之前的顶级注释和处理指令（启用时）      |
| `epilogue`    | `Array<XmlComment \| XmlProcessingInstruction>` | 根元素之后的顶级注释和处理指令（启用时）      |

### DOM 辅助函数

| 函数                           | 描述                   |
| ------------------------------ | ---------------------- |
| `Xml.findChild(el, name)`      | 按名称查找第一个子元素 |
| `Xml.listChildren(el, name)`   | 按名称查找所有子元素   |
| `Xml.getText(node)`            | 递归获取文本内容       |
| `Xml.getAttr(el, name)`        | 获取属性值             |
| `Xml.eachElement(el, visitor)` | 深度优先遍历           |

### Xml.toPlainObject

```text
Xml.toPlainObject(
  element: XmlElement,
  options?: XmlToPlainObjectOptions
): Record<string, unknown>;
```

将 `XmlElement` DOM 树转换为普通 JavaScript 对象。

### Xml.parseToObject

```text
Xml.parseToObject(xml: string, options?: XmlParseToObjectOptions): Record<string, unknown>;
```

将 XML 字符串在单次 SAX 遍历中直接转换为普通 JavaScript 对象。中大型 XML 比 `Xml.parse()` + `Xml.toPlainObject()` 快约 1.6 倍。

**转换选项**（两个函数共用）：

| 选项                   | 默认值    | 描述                                                 |
| ---------------------- | --------- | ---------------------------------------------------- |
| `ignoreAttributes`     | `false`   | 完全丢弃所有属性                                     |
| `attributePrefix`      | `"@_"`    | 属性键前缀（`""` 表示不加前缀）                      |
| `textKey`              | `"#text"` | 混合内容元素中文本内容的键名                         |
| `alwaysArray`          | `false`   | 始终将子元素包装在数组中                             |
| `isArray`              | —         | 回调 `(name) => boolean`，按标签名决定是否包装为数组 |
| `preserveCData`        | `true`    | 在文本中包含 CDATA 值（与 `cdataAsNodes` 相关）      |
| `ignoreWhitespaceText` | `true`    | 丢弃含有子元素的元素中仅含空白的文本                 |

**解析器选项**（仅 `Xml.parseToObject`）：

| 选项                       | 默认值     | 描述                             |
| -------------------------- | ---------- | -------------------------------- |
| `fragment`                 | `false`    | 允许多个根元素                   |
| `maxDepth`                 | `256`      | 最大元素嵌套深度                 |
| `maxEntityExpansions`      | `10000`    | 最大实体扩展次数（XML 炸弹防御） |
| `maxEntityExpansionLength` | `10000000` | 实体展开文本的最大总长度         |
| `invalidCharHandling`      | `"error"`  | 如何处理无效 XML 字符（见下文）  |
| `internalEntities`         | `false`    | 读取并展开内部子集声明的实体     |

`textKey` 与同名属性（`attributePrefix: ""`）或子元素冲突时，各值合并为数组，而不是互相覆盖。

`preserveCData: false` 只在 CDATA 仍是独立节点时才丢弃它：`Xml.parseToObject` 能看到 CDATA，而 `Xml.parse`
默认会把 CDATA 并入相邻文本，除非指定 `cdataAsNodes: true`。因此两条路径只有在设置该选项时才对 CDATA 结果一致。

`Xml.toPlainObject`、`Xml.getText`、`Xml.eachElement` 和 `Xml.query*` 接受调用方自行构建的树：只读取元素自身的
属性；缺少所需字段的节点会被拒绝（`XmlArgumentError`）；元素出现在自己的后代中（形成环）也会被拒绝，否则遍历永远不会结束。
每个字段只读取一次，检查的值即使用的值。这些 helper 只在到达时检查自己读取的字段，并非对整个模型的完整验证：
它们没有到达的畸形节点（例如 `Xml.query` 第一个匹配之后的节点）不会被报告。若模型把同一个元素放在多个父元素之下，
就按其路径展开成的树来遍历：`Xml.query*` 沿每条路径到达该元素及其下的内容。按这种方式会到达超过 2^25 个元素的
单个步骤会被拒绝——内存中没有这么大的树，但每层都把下一层共享两次的小模型就有这么多条路径。

### 查询引擎

```typescript
import { Xml } from "documonster/xml";
```

| 语法           | 描述                               |
| -------------- | ---------------------------------- |
| `a/b/c`        | 匹配子元素 `a`，然后 `b`，然后 `c` |
| `a/b[@id='1']` | 匹配 `id` 属性等于 `"1"` 的 `b`    |
| `a/*/c`        | 通配符：该层级的任意元素名         |
| `a//c`         | 递归下降：`a` 下任意深度的 `c`     |
| `a/b[0]`       | 索引：每个父级 `a` 下的第一个 `b`  |
| `//c`          | 上下文元素下任意深度的 `c`         |

- `Xml.query(element, path)` — 第一个匹配项或 `undefined`
- `Xml.queryAll(element, path)` — 所有匹配项（可能为空），每个节点只出现一次，按文档顺序

其他写法——绝对路径（`/a`）、空步骤（`a/`）、`.`/`..`、或上表之外的谓词（`[@id]`、`[last()]`）——会抛出 `XmlError`，而不是静默地匹配不到。索引从 0 开始，与 XPath 不同。

索引过滤器使用**逐父级语义**：`a/b[0]` 返回每个 `a` 下的第一个 `b`，而非全局第一个 `b`。

### 编码工具

| 函数                            | 描述                                                 |
| ------------------------------- | ---------------------------------------------------- |
| `Xml.encode(text)`              | 编码 XML 内容中的文本（`<`、`>`、`&`、`"`、`'`）     |
| `Xml.decode(text)`              | 将 XML 实体解码回文本                                |
| `Xml.encodeAttr(value)`         | 编码属性值（另将 `\t\n\r` 输出为 `&#x9;&#xA;&#xD;`） |
| `Xml.validateName(name)`        | 验证 XML 元素/属性名                                 |
| `Xml.validateCommentText(text)` | 验证 XML 注释内容文本                                |
| `Xml.encodeCData(text)`         | 编码 CDATA 节的文本（拆分 `]]>`）                    |

### 错误类型

| 类                       | 父类            | 使用场景                                        |
| ------------------------ | --------------- | ----------------------------------------------- |
| `XmlError`               | `BaseError`     | 编码/验证                                       |
| `XmlParseError`          | `XmlError`      | SAX 解析器、DOM 解析器（输入为字节时带 `byte`） |
| `XmlLimitExceededError`  | `XmlParseError` | 超出配置的上限（`limit`、`allowed`）            |
| `XmlTruncatedError`      | `XmlParseError` | 输入在某个结构或记录内部结束                    |
| `XmlChunkAmbiguityError` | `XmlError`      | `Xml.readRecords`：无法确定起点的范围           |
| `XmlWriteError`          | `XmlError`      | 写入器（状态错误）                              |

本模块抛出的错误都继承自 `XmlError`，因此 `catch (e) { if (e instanceof XmlError) ... }` 可以捕获它们。但调用抛出的
不一定都是：`Xml.readRoot` / `Xml.readRecords` 被取消时抛出包内的 `AbortError`，数据源或回调抛出的异常原样传出。

---

## 命名空间支持

当启用 `xmlns: true` 时：

- **前缀解析** — 元素的 QName 解析为元素（`XmlElement`、`XmlSaxTag`）上的 `prefix`、`local`、`uri`。属性前缀同样
  会解析和检查，但属性仍是以 QName 为键的普通 `Record<string, string>`，没有逐个属性的元数据
- **预绑定前缀** — `xml` 预绑定到 `http://www.w3.org/XML/1998/namespace`，`xmlns` 预绑定到 `http://www.w3.org/2000/xmlns/`
- **保留命名空间强制** — 不能将 `xml` 重新绑定到不同的 URI，不能重新绑定 `xmlns`，不能将其他前缀绑定到保留 URI
- **未绑定前缀检测** — 使用未声明前缀的元素和属性会产生错误
- **扩展名称重复检测** — 两个属性具有不同前缀但相同 URI + 本地名时被拒绝
- **多冒号 QName 拒绝** — 命名空间模式下 `<a:b:c/>` 被拒绝
- **作用域管理** — 命名空间声明遵循 XML 作用域规则（被后代继承，可在子元素中覆盖）

注意：根据 XML 命名空间规范 §6.2，无前缀的属性**不**继承默认命名空间。

---

## 安全

- **实体扩展限制** — 防止 XML 炸弹攻击（可通过 `maxEntityExpansions` 和 `maxEntityExpansionLength` 配置）
- **不读取文档以外的任何内容** — 外部 DTD 子集、外部实体（`<!ENTITY e SYSTEM "…">`）和参数实体只被识别，
  从不获取或展开。引用外部实体会报 `undefined entity` 错误；参数实体引用之后声明的实体不会被处理。没有任何
  选项能改变这一点：不打开文件，也不发网络请求。
- **内部实体需显式开启** — 默认情况下只检查 DOCTYPE 内部子集是否良构（注释、处理指令、字面量和声明，因此其中
  的 `]`、`>` 或引号不会结束子集），但不读取，引用其中声明的实体会报告 `undefined entity: e (entities declared
in a DOCTYPE internal subset are not supported)`。设 `internalEntities: true`（`Xml.readRecords` /
  `Xml.readRoot` 总是开启）时，其通用实体按 XML 1.0 §4.4 展开；替换文本含标记（`<`）或 `]]>`、或递归的实体会被拒绝，
  每一次嵌套展开都会在构建文本之前计入 `maxEntityExpansions`，展开文本的长度计入 `maxEntityExpansionLength`，所以实体展开出的文本受 `maxEntityExpansionLength`（UTF-16 码元）约束，而不是输入要多少就给多少。这是对展开的限制，并非对解析器全部内存的限制：正在读取的文本、缓存以及写入目标缓冲的内容都不计入其中。五个预定义实体和字符引用始终可用；调用方可通过
  解析器的 `ENTITIES` 映射添加其他实体——只认它自身的字符串值，且引用必须是 XML Name，因此 `&constructor;` 是未定义实体，而不是
  `Object`，`&1;` 无论映射里有什么都会被拒绝。调用方自定义的这些实体，其展开文本同样计入 `maxEntityExpansionLength`。
- **使用属性列表声明** — 按 XML 1.0 §5.1 对读取内部子集的解析器的要求，无论 `internalEntities` 如何设置：元素
  省略了带默认值声明的属性时会得到该默认值；类型不是 CDATA 的属性值，其中的空格会被合并并去掉首尾。默认值引用
  内部子集声明的实体时，开启 `internalEntities` 会展开；未开启时，在用到该默认值的地方报 `undefined entity` 错误；展开得到的默认值在声明处计入
  `maxEntityExpansions` 与 `maxEntityExpansionLength` 一次：得到它的元素共用同一个字符串。
  声明了 `standalone="yes"` 的文档必须声明它用到的内容，未读取的参数实体引用之后的声明仍会被读取。`Xml.readRoot`
  的 `context` 携带这些声明（`attributes`），因此按范围读取得到的记录与整读一致。解析器为读取后文而保留的内部子集信息（实体绑定、属性声明）同样受
  `maxContextBytes` 限制（按字符计），无论是否开启 `internalEntities`。
- **嵌套深度限制** — 防止深度嵌套 XML 导致的栈溢出（可通过 `maxDepth` 配置）
- **重复属性拒绝** — XML 1.0 §3.1 WFC: Unique Att Spec（报告错误，以最后一个值恢复）
- **名称注入防护** — 写入器通过 `Xml.validateName()` 验证元素和属性名
- **注释/CDATA 安全** — `Xml.validateCommentText()` 拒绝 `--`，`Xml.encodeCData()` 拆分 `]]>`
- **BOM 处理** — 输入开头的 UTF-8 BOM 被静默去除
- **原型污染防护** — 属性映射和转换得到的对象都是空原型对象，因此 `__proto__`、`constructor`、`prototype` 等键
  作为普通数据保留，无法触及 `Object.prototype`
- **无效字符处理** — 写入器通过 `Xml.encode()` 去除无效 XML 1.0 字符；解析器行为可通过 `invalidCharHandling` 配置

---

## 无效字符处理

实际的 XML 数据（尤其来自第三方 XLSX 文件）可能包含 XML 1.0 规范中无效的字符 — 例如 `0x01`–`0x08`、`0x0B`、`0x0C`、`0x0E`–`0x1F`、孤立代理项，以及非字符 `U+FFFE`/`U+FFFF`。`0x7F`（DEL）也按同样方式处理：XML 1.0 本身允许 DEL，只是不推荐使用；拒绝它是本库的策略，因为它在本库读写的文档中没有用途，字符引用 `&#x7F;` 仍会被读取。

`invalidCharHandling` 选项控制解析器的响应方式：

| 值          | 行为                                      |
| ----------- | ----------------------------------------- |
| `"error"`   | 通过错误处理器报告或抛出异常 **（默认）** |
| `"skip"`    | 静默丢弃无效字符                          |
| `"replace"` | 替换为 U+FFFD（Unicode 替换字符）         |

### 示例

```typescript
import { Xml } from "documonster/xml";

// 默认：严格模式 — 遇到 0x7F 时抛出异常
Xml.parse("<root>hello\x7fworld</root>");
// => XmlParseError: invalid XML character: 0x7f

// skip 模式 — 无效字符被移除
const doc = Xml.parse("<root>hello\x7fworld</root>", { invalidCharHandling: "skip" });
// doc.root 文本内容: "helloworld"

// replace 模式 — 无效字符变为 U+FFFD
const doc2 = Xml.parse("<root>hello\x7fworld</root>", { invalidCharHandling: "replace" });
// doc.root 文本内容: "hello\uFFFDworld"

// SAX 解析器使用 skip 模式
const parser = new Xml.SaxParser({ invalidCharHandling: "skip" });
parser.on("text", text => console.log(text)); // "helloworld"
parser.write("<root>hello\x7fworld</root>");
parser.close();
```

### 如何选择

- **`"error"`（默认）** — 用于严格的 XML 验证、测试，或当你控制 XML 来源时。
- **`"skip"`** — 用于读取不可信/脏数据 XML（如第三方 XLSX 文件），需要静默丢弃无效字符时。这是 Excel XLSX 读取器内部使用的模式。
- **`"replace"`** — 用于需要保留无效字符*位置*的场景（如诊断或数据取证），同时不让解析器崩溃。

> **注意：** XML _写入器_（`Xml.Writer`、`Xml.StreamWriter`）始终通过 `Xml.encode()` 去除无效字符 — 此选项仅影响*解析器*。
