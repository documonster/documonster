# Formula 模块

[English](README.md)

独立的 Excel 兼容公式引擎 — tokenizer、parser、compiler、evaluator、依赖图、动态数组 spill 物化器,433 个内置函数。零运行时依赖。

## 两种使用方式

本模块提供**两个互补的入口**,按你的集成方式选一个就行 — 没用到的那个会被 tree-shake 完全摇掉。

| 模式                   | 入口                     | 适用场景                                                                                  |
| ---------------------- | ------------------------ | ----------------------------------------------------------------------------------------- |
| **配合 `Workbook` 用** | `installFormulaEngine()` | 你使用 excel 模块的 `Workbook` 类,想要 `wb.calculateFormulas()` 和 PDF 自动重算。         |
| **单独使用 / 函数式**  | `calculateFormulas(wb)`  | 你操作的是 `WorkbookLike` 对象(自定义宿主、服务端重算、测试)— **完全不用 excel 运行时**。 |

两种模式跑的是同一套引擎代码,只是宿主胶水不一样。tree-shake 数字见
[为什么单独一个 subpath?](#为什么单独一个-subpath)。

## 特性

- **完整表达式流水线** — tokenizer → AST → binder → 编译产物 → evaluator → writeback 计划
- **依赖图** — 拓扑求值、循环引用检测、迭代计算
- **动态数组 + spill** — `FILTER`、`SORT`、`UNIQUE`、`SEQUENCE`,spill 冲突检测,幽灵单元格清理
- **高阶函数** — `LAMBDA`、`LET`、`MAP`、`REDUCE`、`SCAN`、`BYROW`、`BYCOL`
- **数组语义** — 隐式相交、广播、CSE 数组公式
- **结构化引用** — 表格、`[#This Row]`、`[@Column]`、`[#Totals]`
- **共享公式** — 读取、翻译、重算
- **定义名称** — scoped、公式型、区域并集
- **跨表引用** — `Sheet2!A1`、3D 范围 `Sheet1:Sheet3!A1`
- **R1C1 寻址** — 支持 A1 和 R1C1 两种模式(通过 `INDIRECT`)
- **433 个内置 Excel 函数**,分为 11 大类
- **零运行时依赖**

## 函数覆盖

| 类别       | 数量 | 代表函数                                                                                |
| ---------- | ---- | --------------------------------------------------------------------------------------- |
| 数学与三角 | ~70  | `SUM`、`PRODUCT`、`ROUND`、`CEILING`、`POWER`、`MMULT`、`MDETERM`、`SIN`、`ATAN2`       |
| 文本       | ~55  | `CONCAT`、`TEXTJOIN`、`TEXT`、`LEFT`、`MID`、`SUBSTITUTE`、`REGEXTEST`、`REGEXEXTRACT`  |
| 逻辑       | ~15  | `IF`、`IFS`、`AND`、`OR`、`SWITCH`、`IFERROR`、`XOR`                                    |
| 日期与时间 | ~30  | `TODAY`、`DATEDIF`、`EDATE`、`NETWORKDAYS`、`WEEKNUM`、`ISOWEEKNUM`、`WORKDAY.INTL`     |
| 查找与引用 | ~25  | `VLOOKUP`、`XLOOKUP`、`XMATCH`、`INDEX`、`OFFSET`、`INDIRECT`、`ADDRESS`、`CHOOSE`      |
| 统计       | ~90  | `AVERAGE`、`STDEV`、`NORM.DIST`、`T.TEST`、`PERCENTILE`、`QUARTILE`、`CORREL`、`LINEST` |
| 金融       | ~50  | `PMT`、`IRR`、`NPV`、`XIRR`、`RATE`、`PRICE`、`DURATION`、`COUPNUM`、`ACCRINT`          |
| 动态数组   | ~25  | `FILTER`、`SORT`、`UNIQUE`、`SEQUENCE`、`TAKE`、`DROP`、`VSTACK`、`TEXTSPLIT`、`LAMBDA` |
| 数据库     | ~12  | `DSUM`、`DCOUNT`、`DAVERAGE`、`DMAX`、`DMIN`、`DSTDEV`、`DPRODUCT`                      |
| 工程       | ~45  | `DEC2BIN`、`BITAND`、`COMPLEX`、`IMSUM`、`ERF`、`BESSELJ`                               |
| 信息       | ~20  | `ISNUMBER`、`ISBLANK`、`ISREF`、`N`、`TYPE`、`CELL`、`FORMULA`                          |

完整列表见 `functions/`,注册入口在 `runtime/function-registry.ts`。

## 快速开始

### 配合 `Workbook` 用(最常见)

```typescript
import { Workbook } from "documonster";
import { installFormulaEngine } from "documonster/formula";

// 启动时调用一次 — 让 Workbook.calculateFormulas()、PDF bridge 自动重算,
// 以及 XLSX 加载时的 defined-name 严格分类生效。
installFormulaEngine();

const wb = new Workbook();
const ws = wb.addWorksheet("Sheet1");
ws.getCell("A1").value = 10;
ws.getCell("A2").value = 20;
ws.getCell("A3").value = 30;
ws.getCell("A4").value = { formula: "SUM(A1:A3)" };

wb.calculateFormulas();
console.log(ws.getCell("A4").result); // 60
```

### 单独使用 / 函数式

引擎接受任何形状符合 `WorkbookLike` 的对象 — **完全不需要** excel 模块。只 import
`calculateFormulas` 的 bundle 里**零** excel 运行时代码。

```typescript
import { calculateFormulas, type WorkbookLike } from "documonster/formula";

// 你自己的数据结构 — 只要实现 WorkbookLike 就行。
// 不需要 Workbook 类,也不需要 installFormulaEngine()。
const wb: WorkbookLike = buildMyWorkbookLike();

calculateFormulas(wb); // 纯函数,零全局副作用
```

适用场景:

- 服务端对已缓存 XLSX 的重算
- 已有自己数据模型的自定义表格宿主
- 需要每实例独立行为的测试和 benchmark
- 并发求值多个 workbook,不污染进程全局状态

### 针对单个 Workbook 的 syntax probe

如果你用 `Workbook` 但只想让**某个实例**走严格的 defined-name 分类,而不污染全局状态,
显式注入 probe:

```typescript
import { Workbook } from "documonster";
import { createFormulaSyntaxProbe } from "documonster/formula";

const wb = new Workbook({ formulaSyntaxProbe: createFormulaSyntaxProbe() });
await wb.xlsx.load(buffer);
// defined names 用注入的 probe 严格分类;
// 不需要调 installFormulaEngine()。
```

### 不求值,只 tokenize / parse

```typescript
import { tokenize, parse } from "documonster/formula";

const tokens = tokenize("SUM(A1:B10) + VLOOKUP(key, table, 2, FALSE)");
const ast = parse(tokens); // 语法错误时抛异常
```

## 为什么单独一个 subpath?

公式引擎 minified 后约 200 KB。大多数 `documonster` 用户只读写 XLSX、让 Excel 自己重算 — 无条件把引擎打进这些 bundle 是一笔看不见的巨大成本。

subpath 给你三种 tree-shaking 结果:

| 导入方式                                                 | Excel 模块 | Formula 引擎 |
| -------------------------------------------------------- | ---------- | ------------ |
| 只从根路径 import `Workbook`                             | ✓          | ✗            |
| 从 `/formula` import `calculateFormulas`                 | ✗          | ✓            |
| import `Workbook` + `/formula` 的 `installFormulaEngine` | ✓          | ✓            |

函数式 `calculateFormulas` API 通过 `WorkbookLike` 结构化接口工作，**不引入任何 excel 运行时代码** — 只要对象形状符合 workbook 接口就能传进去。服务端对已缓存的 XLSX 做重算也可以：excel 的 import 留在 excel bundle 里。

> **IIFE 说明:** `<script>` 标签的 IIFE 产物
> (`dist/iife/documonster.iife.min.js`) 刻意不包含公式引擎,保持精简。
> 如果通过 `<script>` 标签使用时需要公式计算,请改用 ESM 并调用本
> subpath 的 `installFormulaEngine()`。

## 示例

可运行示例在 `src/modules/formula/examples/`:

| 文件                         | 演示内容                                                       |
| ---------------------------- | -------------------------------------------------------------- |
| `formula-math.ts`            | 算术、舍入、三角、矩阵、幂与对数                               |
| `formula-text.ts`            | 切片、查找/替换、拼接、格式化、正则                            |
| `formula-logical.ts`         | `IF`/`IFS`、布尔运算、`IFERROR`、`SWITCH`、`CHOOSE`            |
| `formula-date.ts`            | 日期构造、提取、时长、工作日、格式化                           |
| `formula-lookup.ts`          | `VLOOKUP`、`XLOOKUP`、`INDEX/MATCH`、`OFFSET`、`INDIRECT`      |
| `formula-statistical.ts`     | 描述统计、条件聚合、回归、概率分布                             |
| `formula-financial.ts`       | 贷款、时值计算、NPV/IRR、折旧                                  |
| `formula-dynamic-array.ts`   | `FILTER`/`SORT`/`UNIQUE`、spill、`SEQUENCE`、`LAMBDA`/`REDUCE` |
| `formula-database.ts`        | `DSUM`/`DCOUNT`/`DAVERAGE` + 条件区域                          |
| `formula-engineering.ts`     | 进制转换、位运算、复数、ERF/BESSELJ                            |
| `formula-standalone.ts`      | 函数式 API + 不求值的 `tokenize`/`parse`                       |
| `formula-pdf-integration.ts` | `excelToPdf()` 中的自动重算                                    |

运行任意示例:

```bash
npx tsx src/modules/formula/examples/formula-math.ts
npx tsx src/modules/formula/examples/formula-dynamic-array.ts
npx tsx src/modules/formula/examples/formula-pdf-integration.ts
# 输出: tmp/formula-examples/formula-pdf-integration.pdf
```

## 架构

引擎是一个六层流水线(见 `AGENTS.md` 的 Layer 3,说明它在模块依赖图里的位置):

```
┌─ syntax/        tokenizer → parser → AST
├─ compile/       binder、依赖分析、编译产物
├─ runtime/       evaluator、函数注册表、RuntimeValue
├─ functions/     433 个函数实现(分 11 个文件)
├─ materialize/   spill 引擎、幽灵单元格跟踪、writeback 计划
└─ integration/   workbook adapter、snapshot、calculate-formulas 入口
```

所有层只依赖 `materialize/types.ts` 里的**结构化接口**(`WorkbookLike`、`WorksheetLike`、`CellLike` 等),不依赖 excel 模块里的具体 `Workbook`/`Worksheet`/`Cell` 类。任何实现这些接口的宿主都能驱动引擎。

## API

### `installFormulaEngine(): void`

把 `calculateFormulas` 接到 `Workbook.calculateFormulas()` 和 `excelToPdf()` 的预渲染自动重算上,并安装默认 syntax probe 让 XLSX 的 defined names 走严格分类。可多次调用(最后一次注册生效)。仅在你使用 excel 模块的 `Workbook` 类时需要。

### `uninstallFormulaEngine(): void`

对称地重置 `installFormulaEngine()` 设置的两个槽位 — engine 和默认 probe。调用后 `Workbook.calculateFormulas()` 会抛,defined-name 分类回退到保守的 "opaque" 路径。主要用于测试里验证冷启动分类行为。

### `calculateFormulas(workbook: WorkbookLike): void`

函数式入口 — **单独使用**模式的核心。遍历 workbook 所有公式单元格,完整解析依赖后求值,把结果写回 `cell.result`,并将动态数组 spill 物化到幽灵单元格。就地修改 workbook。零全局副作用;对不同 workbook 的并发调用是安全的。接受任何 `WorkbookLike`,**不需要** excel 模块。

### `createFormulaSyntaxProbe(): SyntaxProbe`

构造一个独立的 tokenizer+parser probe — 返回一个函数,参数字符串能解析为公式表达式时返回 `true`。通过 `new Workbook({ formulaSyntaxProbe })` 或 `new DefinedNames(probe)` 注入,让**该实例**的 defined-name 分类走严格路径,**不**触碰进程全局状态。适合测试和多宿主场景。

### `tokenize(source: string): Token[]`

纯词法分析 — 接受公式字符串(带不带前导 `=` 都行),返回扁平 token 流。遇到非法字符抛异常。

### `parse(tokens: Token[]): AstNode`

Pratt parser — 从 token 流构建类型化 AST。结构错误抛异常。

### 结构化类型

`WorkbookLike`、`WorksheetLike`、`CellLike`、`RowLike`、`CellErrorValueLike`、`FormulaResultLike`、`DefinedNameEntry`、`DefinedNamesLike`、`DimensionsLike`、`SpillRegion`、`SyntaxProbe`。

## 兼容性说明

- **日期系统** — 尊重工作簿的 1900 / 1904 设置,包括 1900 年闰年 bug
- **错误传播** — 匹配 Excel 的优先级(`#N/A > #VALUE! > ...`)
- **隐式相交** — 在非动态数组位置应用,和 Excel 365 一致
- **迭代计算** — 默认关闭;通过 `workbook.calcProperties = { iterate: true, iterateCount: 100, iterateDelta: 0.001 }` 开启
- **外部引用** — `[book.xlsx]Sheet!A1` 被解析为 `#REF!`;不读取缓存值

## 许可证

MIT — 与主包一致。
