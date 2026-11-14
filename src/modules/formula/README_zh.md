# Formula 模块

[English](README.md)

Excel 兼容的 tokenizer、parser、compiler、evaluator、依赖图、动态数组 spill 物化器,467 个内置函数。零运行时依赖。

## 使用方式

workbook 重算由 `documonster/excel/formula` 暴露;语法检查由
`documonster/formula` 暴露。两者都是直接函数调用,**无需安装或注册**。

| 任务                  | 入口                                               |
| --------------------- | -------------------------------------------------- |
| 重算 Excel workbook   | `/excel/formula` 的 `ExcelFormula.calculate(wb)`   |
| tokenize / parse 语法 | `/formula` 的 `Formula.tokenize` / `Formula.parse` |

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
- **467 个内置 Excel 函数**,分为 11 大类
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
import { ExcelFormula } from "documonster/excel/formula";

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "Sheet1");
Cell.setValue(ws, "A1", 10);
Cell.setValue(ws, "A2", 20);
Cell.setValue(ws, "A3", 30);
Cell.setValue(ws, "A4", { formula: "SUM(A1:A3)" });

ExcelFormula.calculate(wb);
console.log(Cell.getResult(ws, "A4")); // 60
```

### 重算已加载的工作簿

用 excel 模块加载 XLSX,然后函数式地重算它的公式。没有任何安装或注册步骤。

```typescript
import { Workbook } from "documonster/excel";
import { ExcelFormula } from "documonster/excel/formula";

const wb = await Workbook.read(buffer);
ExcelFormula.calculate(wb); // defined names 完成分类,公式被重算
```

### 不求值,只 tokenize / parse

```typescript
import { Formula } from "documonster/formula";

const tokens = Formula.tokenize("SUM(A1:B10) + VLOOKUP(key, table, 2, FALSE)");
const ast = Formula.parse(tokens); // 语法错误时抛异常
```

## 为什么分成多个 subpath?

公式引擎 minified 后约 200 KB。大多数 `documonster` 用户只读写 XLSX、让 Excel 自己重算 — 无条件把引擎打进这些 bundle 是一笔看不见的巨大成本。

这些 subpath 给你如下 tree-shaking 结果:

| 导入方式                                            | Excel 模块 | Formula 引擎 |
| --------------------------------------------------- | ---------- | ------------ |
| 只从 `/excel` import `Workbook`                     | ✓          | ✗            |
| 从 `/formula` import `Formula.tokenize`             | ✗          | 仅语法       |
| 从 `/excel/formula` import `ExcelFormula.calculate` | ✓          | ✓            |

把重算功能放在独立 subpath(而不是做成 `Workbook` 命名空间的一个成员),才能让"不主动要就不含引擎"在**所有**产物格式下都成立,而不只是在支持成员级 DCE 的打包器下成立。实测:挂到 `Workbook` 上会让 `<script>` 版 excel IIFE 增加约 200 KB,所有 CDN 用户都要付这笔钱 — 因为 IIFE 必须保留入口的全部导出。当前拆分方式由 `pnpm verify:treeshake`(`scripts/treeshake-verify.ts`)在 rolldown 和 rspack 上断言。

> **IIFE 说明:** 没有任何 `<script>` 版 IIFE 产物包含计算引擎。
> `dist/iife/documonster.formula.iife.min.js` 只含 tokenizer 与 parser,
> `dist/iife/documonster.excel.iife.min.js` 不受本 subpath 影响。通过
> `<script>` 使用且需要重算 workbook 时,请改用 ESM/CJS,从
> `documonster/excel/formula` import `ExcelFormula.calculate`。

## 示例

可运行示例在 `src/modules/formula/examples/`:

| 文件                          | 演示内容                                                       |
| ----------------------------- | -------------------------------------------------------------- |
| `formula-math.ts`             | 算术、舍入、三角、矩阵、幂与对数                               |
| `formula-text.ts`             | 切片、查找/替换、拼接、格式化、正则                            |
| `formula-logical.ts`          | `IF`/`IFS`、布尔运算、`IFERROR`、`SWITCH`、`CHOOSE`            |
| `formula-date.ts`             | 日期构造、提取、时长、工作日、格式化                           |
| `formula-lookup.ts`           | `VLOOKUP`、`XLOOKUP`、`INDEX/MATCH`、`OFFSET`、`INDIRECT`      |
| `formula-statistical.ts`      | 描述统计、条件聚合、回归、概率分布                             |
| `formula-financial.ts`        | 贷款、时值计算、NPV/IRR、折旧                                  |
| `formula-dynamic-array.ts`    | `FILTER`/`SORT`/`UNIQUE`、spill、`SEQUENCE`、`LAMBDA`/`REDUCE` |
| `formula-database.ts`         | `DSUM`/`DCOUNT`/`DAVERAGE` + 条件区域                          |
| `formula-engineering.ts`      | 进制转换、位运算、复数、ERF/BESSELJ                            |
| `formula-information.ts`      | `ISNUMBER`/`ISBLANK`/`CELL`/`TYPE` 等信息函数                  |
| `formula-custom-functions.ts` | 注册、覆盖与注销自定义函数                                     |
| `formula-standalone.ts`       | 重算 + 不求值的 `tokenize`/`parse`                             |
| `formula-pdf-integration.ts`  | `Pdf.fromExcel()` 中的自动重算                                 |

运行任意示例:

```bash
pnpm example --filter formula-math
pnpm example --filter formula-dynamic-array
pnpm example --filter formula-pdf-integration
# 输出: tmp/formula-examples/formula-pdf-integration.pdf
```

## 架构

引擎是一个六层流水线(见 `AGENTS.md` 的 Layer 3,说明它在模块依赖图里的位置):

```
┌─ syntax/        tokenizer → parser → AST
├─ compile/       binder、依赖分析、编译产物
├─ runtime/       evaluator、函数注册表、RuntimeValue
├─ functions/     467 个函数实现(分 11 个文件)
├─ materialize/   spill 引擎、幽灵单元格跟踪、writeback 计划
└─ integration/   immutable snapshot、calculate pipeline
```

计算核心是纯 `WorkbookSnapshot + FormulaCalculationState → WritebackPlan` 变换,完全不 import Excel。Excel 侧 adapter 同时拥有两端副作用:计算前捕获 `Workbook.Handle`,计算后应用 plan。AST/spill 持久状态是显式对象,且只有全部写操作成功后才提交。

## API

### `ExcelFormula.calculate(workbook: Workbook.Handle): void`

来自 `documonster/excel/formula`。捕获不可变 snapshot,运行引擎,再把结果写回真实单元格。这是唯一公开的求值入口,也适用于从 XLSX 加载的 workbook。

### PDF 导出重算

`Pdf.fromExcel` 不依赖公式引擎。要在渲染前重算公式,通过 `recalculate` 选项注入 `ExcelFormula.calculate` — 只有主动选用的调用方才会把约 200 KB 的引擎打进 bundle。不传时使用缓存在 XLSX 里的结果(对 Excel 自己写出的文件而言是安全的默认行为)。

```typescript
import { ExcelFormula } from "documonster/excel/formula";
import { Pdf } from "documonster/pdf";

const bytes = await Pdf.fromExcel(wb, { recalculate: ExcelFormula.calculate });
```

### 自定义函数

自定义函数接收并返回带 tag 的 `FormulaValue`。类型与 tag 映射都由同一个 bridge 入口导出:

```typescript
import { Workbook } from "documonster/excel";
import { FormulaValueKind } from "documonster/excel/formula";
import type { FormulaValue } from "documonster/excel/formula";

const wb = Workbook.create();

Workbook.registerFunction(
  wb,
  "TAX",
  (args: FormulaValue[]): FormulaValue => {
    const [amount, rate] = args;
    if (amount?.kind !== FormulaValueKind.Number || rate?.kind !== FormulaValueKind.Number) {
      return { kind: FormulaValueKind.Error, code: "#VALUE!" };
    }
    return { kind: FormulaValueKind.Number, value: amount.value * rate.value };
  },
  { minArity: 2, maxArity: 2 }
);
```

函数内抛出的异常会转成 `#VALUE!`。`Workbook.unregisterFunction(wb, name)` 可移除覆盖并恢复被遮蔽的内建函数。

### 定义名称的语法分类

excel 模块加载 XLSX 时,会用一个内置 syntax probe 对定义名称分类,该 probe 复用本引擎的 `tokenize` + `parse`。这是自动的 — 无需任何设置,而且从不加载 XLSX 的 `Workbook` 永远不会把 tokenizer/parser 拉进来。要按实例覆盖分类行为(例如自定义宿主),把你自己的 probe —— 一个 `(text: string) => boolean` —— 传给 `Workbook.create({ formulaSyntaxProbe })`。你可以用本模块的原语构造一个:

```typescript
import { Workbook } from "documonster/excel";
import { Formula } from "documonster/formula";

const probe = (text: string): boolean => {
  try {
    Formula.parse(text);
    return true;
  } catch {
    return false;
  }
};

const wb = Workbook.create({ formulaSyntaxProbe: probe });
```

### `Formula.tokenize(source: string): FormulaToken[]`

纯词法分析 — 接受公式字符串(带不带前导 `=` 都行),返回扁平 token 流。遇到任何 token 都不以之开头的字符、或 Excel 没有的错误值(`#FOOBAR!`)时抛 `FormulaParseError`。溢出区域运算符 `A1#` 是值为 `"#"` 的 `Operator` token,只在紧跟单元格引用时识别。

### `Formula.parse(source: string | FormulaToken[]): FormulaAstNode`

Pratt 解析器 — 从公式文本(自动分词)或 token 流构建类型化 AST。结构错误时抛异常。`A1#` 解析为包住该单元格引用的 `UnaryOp`(`op: "#"`);文件里保存的形式 `_xlfn.ANCHORARRAY(A1)` 仍是函数调用。两者都按原样打印回去,计算引擎都按锚点单元格溢出的区域求值。

### `Formula.format(ast: FormulaAstNode): string`

`parse` 的逆操作:把 AST 打印回公式文本(不带前导 `=`),括号最少化。

### 错误

`FormulaError`(基类)、`FormulaParseError`(源为文本时携带 0-based 字符偏移 `position`;源为 token 时消息中给出 token 序号),以及 `isFormulaError` 类型守卫。参数类型错误、或 `Formula.format` 无法输出的语法树,均抛出 `FormulaError`。

### 导出的类型

`documonster/formula`,全部带 `Formula` 前缀:`FormulaToken`、`FormulaTokenType`、`FormulaAstNode`、`FormulaNodeType`,以及这两个联合类型的每个成员(`FormulaCellRefNode`、`FormulaFunctionCallNode`、`FormulaCellRefToken` 等)。

`documonster/excel/formula`:`FormulaFunction`、`FormulaValue`,以及构造自定义函数返回值所需的 `FormulaValueKind` tag 映射。

## 兼容性说明

- **日期系统** — 尊重工作簿的 1900 / 1904 设置,包括 1900 年闰年 bug
- **错误传播** — 匹配 Excel 的优先级(`#N/A > #VALUE! > ...`)
- **隐式相交** — 工作簿里的普通(非数组)公式按 Excel 的方式计算:Excel 会与公式所在行或列相交的区域
  照样相交,不在区域内时结果为 `#VALUE!`。通过 API 设置、这样计算会与 Excel 365 不同的公式,改为写成动态数组,按数组计算。
- **哪些公式会溢出** — 公式产生数组时溢出:调用返回数组的函数(`SEQUENCE`、`FILTER`、`TRANSPOSE`、`MAP` 等,取自函数表的
  返回类别)、多元素数组常量(`={1,2,3}`)、溢出运算符 `A1#`,或 `LET` / 被调用的 `LAMBDA` 返回的内容——可直接出现,也可经过运算符
  或值参数(`=SEQUENCE(3)*2`)。只有求值时才能确定的情况(`XLOOKUP` 返回一列还是两列)按数组处理:只有一个值的动态数组就是那个值。
  根部的引用(`=A1:A3`、`=INDEX(A1:C3,0,1)`)不溢出,按 Excel 处理普通公式的方式相交。宿主可按单元格用 `isDynamicArray`
  决定——`false` 让普通公式即使产生数组也保持普通,文件里普通的 `=SEQUENCE(3)` 由此得到左上角的值,与 Excel 一致。Excel 写入器
  恰好在引擎会溢出时把公式存成动态数组。宿主函数（`FormulaFunction`、`Workbook.registerFunction`）以 `returnsArray: true`
  声明返回数组；该声明优先于同名内置函数，未声明的视为返回单个值。
- **空结果为 0** — 结果是空单元格的公式,其值为 0,与 Excel 存储和显示的一致:空 A9 上的 `=A9`、`=INDEX(A1:A3,2)`,以及溢出数组中的每个空元素
  (A2 为空时 `=A1:A3` 溢出为 `1, 0, 3`)。这是单元格本身的值,所以其他公式读取该单元格时看到的也是 0:`ISBLANK` 为 FALSE,`=""` 为 FALSE,
  `COUNTBLANK` 不计入。表达式内部的空单元格仍是空:`A9=""`、`A9=0` 和 `ISBLANK(A9)` 为 TRUE,`LEN(A9)` 为 0,`A9&"x"` 为 `"x"`。
  返回空字符串 `""` 的公式保留空字符串。引擎在结果成为单元格的值时转换,因此回写计划中已是 0,任何宿主无需自行转换,即可得到 Excel 的值。
- **逐元素函数** — 接受单个值的函数收到数组时(`LEN(B1:B5)`、`ABS(A1:A5)`、`ROUND(A1:A5,{0,1})`)
  对每个元素分别计算,和 Excel 一致;哪些参数接受单个值,依据 Excel 的操作数类别。
- **迭代计算** — 默认关闭;通过 `workbook.calcProperties = { iterate: true, iterateCount: 100, iterateDelta: 0.001 }` 开启
- **外部引用** — `[book.xlsx]Sheet!A1` 被解析为 `#REF!` 错误节点,并在 `externalReference` 中保留引用文本,`Formula.format` 会原样写回;不读取缓存值
