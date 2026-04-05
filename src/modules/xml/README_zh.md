# XML 模块

[English](README.md)

零依赖、跨平台的 XML 工具包，支持流式和缓冲两种模式的读写。

## 功能特性

- **双模式写入** — 缓冲式 `XmlWriter` 和流式 `XmlStreamWriter`
- **双模式读取** — SAX 事件驱动解析器 和 DOM 树解析器
- **查询引擎** — 简化路径表达式（`a/b[@id='1']`、`a//c`、`a/b[0]`）
- **完整命名空间支持** — 前缀解析、保留命名空间、未绑定前缀检测
- **安全加固** — 实体扩展限制、嵌套深度限制、重复属性拒绝、BOM 处理
- **编码工具** — `xmlEncode`、`xmlDecode`、`xmlEncodeAttr`、`encodeCData`
- **错误处理** — `XmlError`、`XmlParseError`、`XmlWriteError`

## 快速开始

```typescript
import { SaxParser, parseXml, XmlWriter, queryAll } from "@cj-tech-master/excelts/xml";

// SAX 流式解析
const parser = new SaxParser();
parser.on("opentag", tag => console.log(tag.name, tag.attributes));
parser.write('<root><item id="1">hello</item></root>');
parser.close();

// DOM 解析 + 查询
const doc = parseXml("<root><a><b>1</b><b>2</b></a></root>");
const items = queryAll(doc.root, "a/b"); // 所有 <b> 元素

// 写入 XML
const w = new XmlWriter();
w.openXml();
w.openNode("root");
w.leafNode("item", { id: "1" }, "hello");
w.closeNode();
console.log(w.xml);
```

## 示例

查看 [examples 目录](examples/) 获取可运行代码。
