# Formula 模块

[English](README.md)

独立的 Excel 兼容公式引擎 — tokenizer、parser、compiler、evaluator、依赖图、动态数组 spill 物化器,433 个内置函数。零运行时依赖。

## 两种使用方式

本模块只暴露一个函数式入口 — `Formula.calculate(wb)` — 它能以两种互补的方式工作。不 import 引擎就不会为它付出任何代价。

| 模式                   | 用法                              | 适用场景                                                                                  |
| ---------------------- | --------------------------------- | ----------------------------------------------------------------------------------------- |
| **配合 `Workbook` 用** | `Formula.calculate(wb)`           | 你使用 excel 模块的 `Workbook`,想重算它的公式(以及 PDF 导出时重算)。                      |
| **单独使用 / 函数式**  | `Formula.calculate(workbookLike)` | 你操作的是 `WorkbookLike` 对象(自定义宿主、服务端重算、测试)— **完全不用 excel 运行时**。 |

两种模式跑的是同一套引擎代码,只是传进去的数据不同。**没有任何安装或注册步骤** — 直接调用 `Formula.calculate`。tree-shake 数字见[为什么单独一个 subpath?](#为什么单独一个-subpath)。

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
import { Workbook, Cell } from "documonster/excel";
import { Formula } from "documonster/formula";

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "Sheet1");
Cell.setValue(ws, "A1", 10);
Cell.setValue(ws, "A2", 20);
Cell.setValue(ws, "A3", 30);
Cell.setValue(ws, "A4", { formula: "SUM(A1:A3)" });

Formula.calculate(wb);
console.log(Cell.getResult(ws, "A4")); // 60
```

### 单独使用 / 函数式

引擎接受任何形状符合 `WorkbookLike` 的对象 — 你**完全不必**使用 excel 模块。只 import `Formula.calculate` 的 bundle 里**零** excel 运行时代码。

```typescript
import { Formula, type WorkbookLike } from "documonster/formula";

// 你自己的数据结构 — 只要实现 WorkbookLike 就行。
// 不需要 Workbook 类。
const wb: WorkbookLike = buildMyWorkbookLike();

Formula.calculate(wb); // 纯函数,零全局副作用
```

这种模式适合:

- 服务端对已缓存 XLSX 的重算
- 已有自己数据模型的自定义表格宿主
- 需要每实例独立行为的测试和 benchmark
- 并发求值多个 workbook,不污染进程全局状态

### 重算已加载的工作簿

用 excel 模块加载 XLSX,然后函数式地重算它的公式。没有任何安装或注册步骤。

```typescript
import { Workbook } from "documonster/excel";
import { Formula } from "documonster/formula";

const wb = Workbook.create();
await Workbook.read(wb, buffer);
Formula.calculate(wb); // defined names 完成分类,公式被重算
```

### 不求值,只 tokenize / parse

```typescript
import { Formula } from "documonster/formula";

const tokens = Formula.tokenize("SUM(A1:B10) + VLOOKUP(key, table, 2, FALSE)");
const ast = Formula.parse(tokens); // 语法错误时抛异常
```

## 为什么单独一个 subpath?

公式引擎 minified 后约 200 KB。大多数 `documonster` 用户只读写 XLSX、让 Excel 自己重算 — 无条件把引擎打进这些 bundle 是一笔看不见的巨大成本。

subpath 给你三种 tree-shaking 结果:

| 导入方式                                              | Excel 模块 | Formula 引擎 |
| ----------------------------------------------------- | ---------- | ------------ |
| 只从根路径 import `Workbook`                          | ✓          | ✗            |
| 从 `/formula` import `Formula.calculate`              | ✗          | ✓            |
| import `Workbook` + `/formula` 的 `Formula.calculate` | ✓          | ✓            |

函数式 `Formula.calculate` API 通过 `WorkbookLike` 结构化接口工作,**不引入任何 excel 运行时代码** — 只要对象形状符合 workbook 接口就能传进去。服务端对已缓存的 XLSX 做重算也可以:excel 的 import 留在 excel bundle 里。

> **IIFE 说明:** `<script>` 标签的 IIFE 产物
> (`dist/iife/documonster.iife.min.js`) 刻意不包含公式引擎,保持精简。
> 如果通过 `<script>` 标签使用时需要公式计算,请改用 ESM,从本 subpath
> import `Formula`,然后调用 `Formula.calculate(wb)`。

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
| `formula-pdf-integration.ts` | `Pdf.fromExcel()` 中的自动重算                                 |

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

### `Formula.calculate(workbook: WorkbookLike): void`

唯一的函数式求值入口。遍历 workbook 所有公式单元格,完整解析依赖后求值,把结果写回每个单元格的 `result` 属性,并将动态数组 spill 物化到幽灵单元格。就地修改 workbook。零全局副作用;对不同 workbook 的并发调用是安全的。**没有安装或注册步骤**,也**没有 `Workbook.calculateFormulas()` 方法** — 直接调用 `Formula.calculate(wb)`。接受任何 `WorkbookLike`;不需要 excel 模块,但 excel 模块创建的 `Workbook` 在结构上是兼容的,可直接传入。

### PDF 导出重算

`Pdf.fromExcel` 不依赖公式引擎。要在渲染前重算公式,通过 `recalculate` 选项注入 `Formula.calculate` — 只有主动选用的调用方才会把约 200 KB 的引擎打进 bundle。不传时使用缓存在 XLSX 里的结果(对 Excel 自己写出的文件而言是安全的默认行为)。

```typescript
import { Pdf } from "documonster/pdf";
import { Formula } from "documonster/formula";

const bytes = await Pdf.fromExcel(wb, { recalculate: Formula.calculate });
```

### 定义名称的语法分类

excel 模块加载 XLSX 时,会用一个内置 syntax probe 对定义名称分类,该 probe 复用本引擎的 `tokenize` + `parse`。这是自动的 — 无需任何设置,而且从不加载 XLSX 的 `Workbook` 永远不会把 tokenizer/parser 拉进来。要按实例覆盖分类行为(例如自定义宿主),把你自己的 probe —— 一个 `(text: string) => boolean` —— 传给 `Workbook.create({ formulaSyntaxProbe })`。你可以用本模块的原语构造一个:

```typescript
import { Workbook } from "documonster/excel";
import { Formula } from "documonster/formula";

const probe = (text: string): boolean => {
  try {
    Formula.parse(Formula.tokenize(text));
    return true;
  } catch {
    return false;
  }
};

const wb = Workbook.create({ formulaSyntaxProbe: probe });
```

### `Formula.tokenize(source: string): Token[]`

纯词法分析 — 接受公式字符串(带不带前导 `=` 都行),返回扁平 token 流。遇到非法字符抛异常。

### `Formula.parse(tokens: Token[]): AstNode`

Pratt parser — 从 token 流构建类型化 AST。结构错误抛异常。

### 错误

`FormulaError`(基类)、`FormulaParseError`(携带可选的 0-based `position`),以及 `isFormulaError` 类型守卫。

### 结构化类型

`WorkbookLike`、`WorksheetLike`、`CellLike`、`RowLike`、`CellErrorValueLike`、`FormulaResultLike`、`DefinedNameEntry`、`DefinedNamesLike`、`DimensionsLike`、`SpillRegion`。

## 兼容性说明

- **日期系统** — 尊重工作簿的 1900 / 1904 设置,包括 1900 年闰年 bug
- **错误传播** — 匹配 Excel 的优先级(`#N/A > #VALUE! > ...`)
- **隐式相交** — 在非动态数组位置应用,和 Excel 365 一致
- **迭代计算** — 默认关闭;通过 `workbook.calcProperties = { iterate: true, iterateCount: 100, iterateDelta: 0.001 }` 开启
- **外部引用** — `[book.xlsx]Sheet!A1` 被解析为 `#REF!`;不读取缓存值

## 许可证

MIT — 与主包一致。
