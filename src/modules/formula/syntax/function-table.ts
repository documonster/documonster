/**
 * The BIFF12 function table: what an `iftab` names, how many arguments it takes, and what operand class each of
 * its arguments wants.
 *
 * Split out of the XLSB codec (`@excel/xlsb/formula/ptg`) because it is data — three transcriptions of one
 * specification table plus the operand classes, around a thousand lines that the codec only ever looks up. It lives
 * with the formula engine because the operand classes are not only an encoding matter: they say which arguments a
 * function takes as one value, and the evaluator applies a function element by element over an array handed to such
 * an argument (`LEN(B1:B5)`), as Excel does. One table serves both, so they cannot disagree.
 */

/**
 * Built-in function ids — the complete `Ftab` from MS-XLSB 2.5.98.10, transcribed.
 *
 * **This was a hand-maintained subset of 104 entries, and twelve of them named the wrong function.** Not
 * missing — *wrong*, which is worse, because a wrong pair is silent in both directions. `FIND` was written as
 * 82, and 82 is `SEARCH`; a workbook containing `=FIND("a",A1)` was saved as `=SEARCH("a",A1)`, which is
 * case-insensitive and accepts wildcards, so Excel opened the file without complaint and computed a different
 * answer. The full list of what the twelve actually encoded: `FIND`→SEARCH, `ISNUMBER`→CHAR, `FINDB`→FIND,
 * `TEXT`→INDIRECT, `ODD`→FISHER, `STDEVP`→DEVSQ, `GETPIVOTDATA`→STDEVA, `IFERROR`→ISEVEN, `COUNTIFS`→BESSELY,
 * `SUMIFS`→BESSELI, `AVERAGEIF`→XNPV, `AVERAGEIFS`→PRICEMAT.
 *
 * **Nothing here could have caught that**, and the reason is worth keeping: the encoder and the decoder read
 * the same table, so `FIND` went out as 82 and came back as `FIND`. Every round-trip test passed. Only Excel
 * disagreed, and only about the answer — never about whether the file was valid. This is the same shape as
 * the reference/value operand classes having been swapped: a table shared by both directions cannot be
 * validated by using it in both directions.
 *
 * So the subset is gone. All 475 numeric ids are here, generated from the specification's table rather than
 * curated, because a curated subset is a list that drifts from the thing it describes — and the drift is
 * invisible. `excel/xlsb/__tests__/function-table.test.ts` pins the count and the twelve corrected pairs.
 *
 * **Sixty-one of these are macro-sheet commands** (`SET.NAME`, `GET.CELL`, `ADD.BAR`) rather than worksheet
 * functions. They are real `iftab` values and belong in the table: the decoder needs them to read a macro
 * sheet, and the encoder reaches one only if a formula names it.
 *
 * A *future function* — `XLOOKUP`, `TEXTJOIN`, `IFS`, `CONCAT`, every `*.DIST` — has no `iftab` at all and is
 * not in this table. It is called through `PtgName` pointing at a hidden `_xlfn.` stub; see
 * `excel/xlsb/formula/ptg.ts`.
 */
export const FUNCTION_TABLE: readonly (readonly [number, string])[] = [
  [0, "COUNT"],
  [1, "IF"],
  [2, "ISNA"],
  [3, "ISERROR"],
  [4, "SUM"],
  [5, "AVERAGE"],
  [6, "MIN"],
  [7, "MAX"],
  [8, "ROW"],
  [9, "COLUMN"],
  [10, "NA"],
  [11, "NPV"],
  [12, "STDEV"],
  [13, "DOLLAR"],
  [14, "FIXED"],
  [15, "SIN"],
  [16, "COS"],
  [17, "TAN"],
  [18, "ATAN"],
  [19, "PI"],
  [20, "SQRT"],
  [21, "EXP"],
  [22, "LN"],
  [23, "LOG10"],
  [24, "ABS"],
  [25, "INT"],
  [26, "SIGN"],
  [27, "ROUND"],
  [28, "LOOKUP"],
  [29, "INDEX"],
  [30, "REPT"],
  [31, "MID"],
  [32, "LEN"],
  [33, "VALUE"],
  [34, "TRUE"],
  [35, "FALSE"],
  [36, "AND"],
  [37, "OR"],
  [38, "NOT"],
  [39, "MOD"],
  [40, "DCOUNT"],
  [41, "DSUM"],
  [42, "DAVERAGE"],
  [43, "DMIN"],
  [44, "DMAX"],
  [45, "DSTDEV"],
  [46, "VAR"],
  [47, "DVAR"],
  [48, "TEXT"],
  [49, "LINEST"],
  [50, "TREND"],
  [51, "LOGEST"],
  [52, "GROWTH"],
  [53, "GOTO"],
  [54, "HALT"],
  [55, "RETURN"],
  [56, "PV"],
  [57, "FV"],
  [58, "NPER"],
  [59, "PMT"],
  [60, "RATE"],
  [61, "MIRR"],
  [62, "IRR"],
  [63, "RAND"],
  [64, "MATCH"],
  [65, "DATE"],
  [66, "TIME"],
  [67, "DAY"],
  [68, "MONTH"],
  [69, "YEAR"],
  [70, "WEEKDAY"],
  [71, "HOUR"],
  [72, "MINUTE"],
  [73, "SECOND"],
  [74, "NOW"],
  [75, "AREAS"],
  [76, "ROWS"],
  [77, "COLUMNS"],
  [78, "OFFSET"],
  [79, "ABSREF"],
  [80, "RELREF"],
  [81, "ARGUMENT"],
  [82, "SEARCH"],
  [83, "TRANSPOSE"],
  [84, "ERROR"],
  [85, "STEP"],
  [86, "TYPE"],
  [87, "ECHO"],
  [88, "SET.NAME"],
  [89, "CALLER"],
  [90, "DEREF"],
  [91, "WINDOWS"],
  [93, "DOCUMENTS"],
  [94, "ACTIVE.CELL"],
  [95, "SELECTION"],
  [96, "RESULT"],
  [97, "ATAN2"],
  [98, "ASIN"],
  [99, "ACOS"],
  [100, "CHOOSE"],
  [101, "HLOOKUP"],
  [102, "VLOOKUP"],
  [103, "LINKS"],
  [104, "INPUT"],
  [105, "ISREF"],
  [106, "GET.FORMULA"],
  [107, "GET.NAME"],
  [108, "SET.VALUE"],
  [109, "LOG"],
  [110, "EXEC"],
  [111, "CHAR"],
  [112, "LOWER"],
  [113, "UPPER"],
  [114, "PROPER"],
  [115, "LEFT"],
  [116, "RIGHT"],
  [117, "EXACT"],
  [118, "TRIM"],
  [119, "REPLACE"],
  [120, "SUBSTITUTE"],
  [121, "CODE"],
  [122, "NAMES"],
  [123, "DIRECTORY"],
  [124, "FIND"],
  [125, "CELL"],
  [126, "ISERR"],
  [127, "ISTEXT"],
  [128, "ISNUMBER"],
  [129, "ISBLANK"],
  [130, "T"],
  [131, "N"],
  [132, "FOPEN"],
  [133, "FCLOSE"],
  [134, "FSIZE"],
  [135, "FREADLN"],
  [136, "FREAD"],
  [137, "FWRITELN"],
  [138, "FWRITE"],
  [139, "FPOS"],
  [140, "DATEVALUE"],
  [141, "TIMEVALUE"],
  [142, "SLN"],
  [143, "SYD"],
  [144, "DDB"],
  [145, "GET.DEF"],
  [146, "REFTEXT"],
  [147, "TEXTREF"],
  [148, "INDIRECT"],
  [149, "REGISTER"],
  [150, "CALL"],
  [151, "ADD.BAR"],
  [152, "ADD.MENU"],
  [153, "ADD.COMMAND"],
  [154, "ENABLE.COMMAND"],
  [155, "CHECK.COMMAND"],
  [156, "RENAME.COMMAND"],
  [157, "SHOW.BAR"],
  [158, "DELETE.MENU"],
  [159, "DELETE.COMMAND"],
  [160, "GET.CHART.ITEM"],
  [161, "DIALOG.BOX"],
  [162, "CLEAN"],
  [163, "MDETERM"],
  [164, "MINVERSE"],
  [165, "MMULT"],
  [166, "FILES"],
  [167, "IPMT"],
  [168, "PPMT"],
  [169, "COUNTA"],
  [170, "CANCEL.KEY"],
  [171, "FOR"],
  [172, "WHILE"],
  [173, "BREAK"],
  [174, "NEXT"],
  [175, "INITIATE"],
  [176, "REQUEST"],
  [177, "POKE"],
  [178, "EXECUTE"],
  [179, "TERMINATE"],
  [180, "RESTART"],
  [181, "HELP"],
  [182, "GET.BAR"],
  [183, "PRODUCT"],
  [184, "FACT"],
  [185, "GET.CELL"],
  [186, "GET.WORKSPACE"],
  [187, "GET.WINDOW"],
  [188, "GET.DOCUMENT"],
  [189, "DPRODUCT"],
  [190, "ISNONTEXT"],
  [191, "GET.NOTE"],
  [192, "NOTE"],
  [193, "STDEVP"],
  [194, "VARP"],
  [195, "DSTDEVP"],
  [196, "DVARP"],
  [197, "TRUNC"],
  [198, "ISLOGICAL"],
  [199, "DCOUNTA"],
  [200, "DELETE.BAR"],
  [201, "UNREGISTER"],
  [204, "USDOLLAR"],
  [205, "FINDB"],
  [206, "SEARCHB"],
  [207, "REPLACEB"],
  [208, "LEFTB"],
  [209, "RIGHTB"],
  [210, "MIDB"],
  [211, "LENB"],
  [212, "ROUNDUP"],
  [213, "ROUNDDOWN"],
  [214, "ASC"],
  [215, "DBCS"],
  [216, "RANK"],
  [219, "ADDRESS"],
  [220, "DAYS360"],
  [221, "TODAY"],
  [222, "VDB"],
  [223, "ELSE"],
  [224, "ELSE.IF"],
  [225, "END.IF"],
  [226, "FOR.CELL"],
  [227, "MEDIAN"],
  [228, "SUMPRODUCT"],
  [229, "SINH"],
  [230, "COSH"],
  [231, "TANH"],
  [232, "ASINH"],
  [233, "ACOSH"],
  [234, "ATANH"],
  [235, "DGET"],
  [236, "CREATE.OBJECT"],
  [237, "VOLATILE"],
  [238, "LAST.ERROR"],
  [239, "CUSTOM.UNDO"],
  [240, "CUSTOM.REPEAT"],
  [241, "FORMULA.CONVERT"],
  [242, "GET.LINK.INFO"],
  [243, "TEXT.BOX"],
  [244, "INFO"],
  [245, "GROUP"],
  [246, "GET.OBJECT"],
  [247, "DB"],
  [248, "PAUSE"],
  [251, "RESUME"],
  [252, "FREQUENCY"],
  [253, "ADD.TOOLBAR"],
  [254, "DELETE.TOOLBAR"],
  [256, "RESET.TOOLBAR"],
  [257, "EVALUATE"],
  [258, "GET.TOOLBAR"],
  [259, "GET.TOOL"],
  [260, "SPELLING.CHECK"],
  [261, "ERROR.TYPE"],
  [262, "APP.TITLE"],
  [263, "WINDOW.TITLE"],
  [264, "SAVE.TOOLBAR"],
  [265, "ENABLE.TOOL"],
  [266, "PRESS.TOOL"],
  [267, "REGISTER.ID"],
  [268, "GET.WORKBOOK"],
  [269, "AVEDEV"],
  [270, "BETADIST"],
  [271, "GAMMALN"],
  [272, "BETAINV"],
  [273, "BINOMDIST"],
  [274, "CHIDIST"],
  [275, "CHIINV"],
  [276, "COMBIN"],
  [277, "CONFIDENCE"],
  [278, "CRITBINOM"],
  [279, "EVEN"],
  [280, "EXPONDIST"],
  [281, "FDIST"],
  [282, "FINV"],
  [283, "FISHER"],
  [284, "FISHERINV"],
  [285, "FLOOR"],
  [286, "GAMMADIST"],
  [287, "GAMMAINV"],
  [288, "CEILING"],
  [289, "HYPGEOMDIST"],
  [290, "LOGNORMDIST"],
  [291, "LOGINV"],
  [292, "NEGBINOMDIST"],
  [293, "NORMDIST"],
  [294, "NORMSDIST"],
  [295, "NORMINV"],
  [296, "NORMSINV"],
  [297, "STANDARDIZE"],
  [298, "ODD"],
  [299, "PERMUT"],
  [300, "POISSON"],
  [301, "TDIST"],
  [302, "WEIBULL"],
  [303, "SUMXMY2"],
  [304, "SUMX2MY2"],
  [305, "SUMX2PY2"],
  [306, "CHITEST"],
  [307, "CORREL"],
  [308, "COVAR"],
  [309, "FORECAST"],
  [310, "FTEST"],
  [311, "INTERCEPT"],
  [312, "PEARSON"],
  [313, "RSQ"],
  [314, "STEYX"],
  [315, "SLOPE"],
  [316, "TTEST"],
  [317, "PROB"],
  [318, "DEVSQ"],
  [319, "GEOMEAN"],
  [320, "HARMEAN"],
  [321, "SUMSQ"],
  [322, "KURT"],
  [323, "SKEW"],
  [324, "ZTEST"],
  [325, "LARGE"],
  [326, "SMALL"],
  [327, "QUARTILE"],
  [328, "PERCENTILE"],
  [329, "PERCENTRANK"],
  [330, "MODE"],
  [331, "TRIMMEAN"],
  [332, "TINV"],
  [334, "MOVIE.COMMAND"],
  [335, "GET.MOVIE"],
  [336, "CONCATENATE"],
  [337, "POWER"],
  [338, "PIVOT.ADD.DATA"],
  [339, "GET.PIVOT.TABLE"],
  [340, "GET.PIVOT.FIELD"],
  [341, "GET.PIVOT.ITEM"],
  [342, "RADIANS"],
  [343, "DEGREES"],
  [344, "SUBTOTAL"],
  [345, "SUMIF"],
  [346, "COUNTIF"],
  [347, "COUNTBLANK"],
  [348, "SCENARIO.GET"],
  [349, "OPTIONS.LISTS.GET"],
  [350, "ISPMT"],
  [351, "DATEDIF"],
  [352, "DATESTRING"],
  [353, "NUMBERSTRING"],
  [354, "ROMAN"],
  [355, "OPEN.DIALOG"],
  [356, "SAVE.DIALOG"],
  [357, "VIEW.GET"],
  [358, "GETPIVOTDATA"],
  [359, "HYPERLINK"],
  [360, "PHONETIC"],
  [361, "AVERAGEA"],
  [362, "MAXA"],
  [363, "MINA"],
  [364, "STDEVPA"],
  [365, "VARPA"],
  [366, "STDEVA"],
  [367, "VARA"],
  [368, "BAHTTEXT"],
  [369, "THAIDAYOFWEEK"],
  [370, "THAIDIGIT"],
  [371, "THAIMONTHOFYEAR"],
  [372, "THAINUMSOUND"],
  [373, "THAINUMSTRING"],
  [374, "THAISTRINGLENGTH"],
  [375, "ISTHAIDIGIT"],
  [376, "ROUNDBAHTDOWN"],
  [377, "ROUNDBAHTUP"],
  [378, "THAIYEAR"],
  [379, "RTD"],
  [380, "CUBEVALUE"],
  [381, "CUBEMEMBER"],
  [382, "CUBEMEMBERPROPERTY"],
  [383, "CUBERANKEDMEMBER"],
  [384, "HEX2BIN"],
  [385, "HEX2DEC"],
  [386, "HEX2OCT"],
  [387, "DEC2BIN"],
  [388, "DEC2HEX"],
  [389, "DEC2OCT"],
  [390, "OCT2BIN"],
  [391, "OCT2HEX"],
  [392, "OCT2DEC"],
  [393, "BIN2DEC"],
  [394, "BIN2OCT"],
  [395, "BIN2HEX"],
  [396, "IMSUB"],
  [397, "IMDIV"],
  [398, "IMPOWER"],
  [399, "IMABS"],
  [400, "IMSQRT"],
  [401, "IMLN"],
  [402, "IMLOG2"],
  [403, "IMLOG10"],
  [404, "IMSIN"],
  [405, "IMCOS"],
  [406, "IMEXP"],
  [407, "IMARGUMENT"],
  [408, "IMCONJUGATE"],
  [409, "IMAGINARY"],
  [410, "IMREAL"],
  [411, "COMPLEX"],
  [412, "IMSUM"],
  [413, "IMPRODUCT"],
  [414, "SERIESSUM"],
  [415, "FACTDOUBLE"],
  [416, "SQRTPI"],
  [417, "QUOTIENT"],
  [418, "DELTA"],
  [419, "GESTEP"],
  [420, "ISEVEN"],
  [421, "ISODD"],
  [422, "MROUND"],
  [423, "ERF"],
  [424, "ERFC"],
  [425, "BESSELJ"],
  [426, "BESSELK"],
  [427, "BESSELY"],
  [428, "BESSELI"],
  [429, "XIRR"],
  [430, "XNPV"],
  [431, "PRICEMAT"],
  [432, "YIELDMAT"],
  [433, "INTRATE"],
  [434, "RECEIVED"],
  [435, "DISC"],
  [436, "PRICEDISC"],
  [437, "YIELDDISC"],
  [438, "TBILLEQ"],
  [439, "TBILLPRICE"],
  [440, "TBILLYIELD"],
  [441, "PRICE"],
  [442, "YIELD"],
  [443, "DOLLARDE"],
  [444, "DOLLARFR"],
  [445, "NOMINAL"],
  [446, "EFFECT"],
  [447, "CUMPRINC"],
  [448, "CUMIPMT"],
  [449, "EDATE"],
  [450, "EOMONTH"],
  [451, "YEARFRAC"],
  [452, "COUPDAYBS"],
  [453, "COUPDAYS"],
  [454, "COUPDAYSNC"],
  [455, "COUPNCD"],
  [456, "COUPNUM"],
  [457, "COUPPCD"],
  [458, "DURATION"],
  [459, "MDURATION"],
  [460, "ODDLPRICE"],
  [461, "ODDLYIELD"],
  [462, "ODDFPRICE"],
  [463, "ODDFYIELD"],
  [464, "RANDBETWEEN"],
  [465, "WEEKNUM"],
  [466, "AMORDEGRC"],
  [467, "AMORLINC"],
  [469, "ACCRINT"],
  [470, "ACCRINTM"],
  [471, "WORKDAY"],
  [472, "NETWORKDAYS"],
  [473, "GCD"],
  [474, "MULTINOMIAL"],
  [475, "LCM"],
  [476, "FVSCHEDULE"],
  [477, "CUBEKPIMEMBER"],
  [478, "CUBESET"],
  [479, "CUBESETCOUNT"],
  [480, "IFERROR"],
  [481, "COUNTIFS"],
  [482, "SUMIFS"],
  [483, "AVERAGEIF"],
  [484, "AVERAGEIFS"]
];

/**
 * The arity of every fixed-arity function, by `iftab`.
 *
 * **`PtgFunc` and `PtgFuncVar` are not interchangeable, and this codec used only the second one.** A function
 * with a fixed signature is called through `PtgFunc`, which carries no argument count because the count is
 * implied; a variadic one is called through `PtgFuncVar`, which carries it. Emitting `PtgFuncVar` for `MOD` —
 * exactly two arguments — produced a formula Excel repaired: `Repaired Records: Conditional formatting from
 * /xl/worksheets/sheet1.bin`, in six sheets of one workbook. Excel's own bytes for the same rule are
 * `41 27 00` where this wrote `42 02 27 00`.
 *
 * The reader had the mirror of the same gap and was louder about it: `PtgFunc` threw
 * `not supported yet`, so **no formula Excel wrote for any fixed-arity function could be read at all**.
 *
 * Derived from the `Parameters` column of MS-XLSB 2.5.98.10 rather than curated: a signature is fixed when its
 * grammar has no repetition (`*`), no optional group (`[`) and no alternation, and the arity is then the number
 * of terms — then checked against the token Excel itself writes, which corrected two rows. 257 of the 475 are fixed. Hand-listing that many would drift, and the drift would be
 * invisible — the two tokens differ by one byte and both decode through this same codec.
 */
const FIXED_ARITY: readonly (readonly [number, number])[] = [
  [2, 1],
  [3, 1],
  [10, 0],
  [15, 1],
  [16, 1],
  [17, 1],
  [18, 1],
  [19, 0],
  [20, 1],
  [21, 1],
  [22, 1],
  [23, 1],
  [24, 1],
  [25, 1],
  [26, 1],
  [27, 2],
  [30, 2],
  [31, 3],
  [32, 1],
  [33, 1],
  [34, 0],
  [35, 0],
  [38, 1],
  [39, 2],
  [40, 3],
  [41, 3],
  [42, 3],
  [43, 3],
  [44, 3],
  [45, 3],
  [47, 3],
  [48, 2],
  [53, 1],
  [61, 3],
  [63, 0],
  [65, 3],
  [66, 3],
  [67, 1],
  [68, 1],
  [69, 1],
  [71, 1],
  [72, 1],
  [73, 1],
  [74, 0],
  [75, 1],
  [76, 1],
  [77, 1],
  [79, 2],
  [80, 2],
  [83, 1],
  [85, 0],
  [86, 1],
  [89, 0],
  [90, 1],
  [94, 0],
  [95, 0],
  [97, 2],
  [98, 1],
  [99, 1],
  [105, 1],
  [106, 1],
  [108, 2],
  [111, 1],
  [112, 1],
  [113, 1],
  [114, 1],
  // `EXACT` was missing and `LENB` was given `MIDB`'s three: Excel writes `EXACT(D1,D1)` as `… 41 75 00` and
  // `LENB("ab")` as `… 41 d3 00` — both `PtgFunc`, two and one argument. Reading either threw.
  [117, 2],
  [118, 1],
  [119, 4],
  [121, 1],
  [126, 1],
  [127, 1],
  [128, 1],
  [129, 1],
  [130, 1],
  [131, 1],
  [133, 1],
  [134, 1],
  [135, 1],
  [136, 2],
  [137, 2],
  [138, 2],
  [140, 1],
  [141, 1],
  [142, 3],
  [143, 4],
  [161, 1],
  [162, 1],
  [163, 1],
  [164, 1],
  [165, 2],
  [172, 1],
  [173, 0],
  [174, 0],
  [175, 2],
  [176, 2],
  [177, 3],
  [178, 2],
  [179, 1],
  [184, 1],
  [186, 1],
  [189, 3],
  [190, 1],
  [195, 3],
  [196, 3],
  [198, 1],
  [199, 3],
  [200, 1],
  [201, 1],
  [207, 4],
  [210, 3],
  [211, 1],
  [212, 2],
  [213, 2],
  [214, 1],
  [215, 1],
  [221, 0],
  [223, 0],
  [224, 1],
  [225, 0],
  [229, 1],
  [230, 1],
  [231, 1],
  [232, 1],
  [233, 1],
  [234, 1],
  [235, 3],
  [238, 0],
  [244, 1],
  [245, 0],
  [252, 2],
  [254, 1],
  [256, 1],
  [257, 1],
  [261, 1],
  [265, 3],
  [266, 3],
  [271, 1],
  [273, 4],
  [274, 2],
  [275, 2],
  [276, 2],
  [277, 3],
  [278, 3],
  [279, 1],
  [280, 3],
  [281, 3],
  [282, 3],
  [283, 1],
  [284, 1],
  [285, 2],
  [286, 4],
  [287, 3],
  [288, 2],
  [289, 4],
  [290, 3],
  [291, 3],
  [292, 3],
  [293, 4],
  [294, 1],
  [295, 3],
  [296, 1],
  [297, 3],
  [298, 1],
  [299, 2],
  [300, 3],
  [301, 3],
  [302, 4],
  [303, 2],
  [304, 2],
  [305, 2],
  [306, 2],
  [307, 2],
  [308, 2],
  [309, 3],
  [310, 2],
  [311, 2],
  [312, 2],
  [313, 2],
  [314, 2],
  [315, 2],
  [316, 4],
  [325, 2],
  [326, 2],
  [327, 2],
  [328, 2],
  [331, 2],
  [332, 2],
  [337, 2],
  [342, 1],
  [343, 1],
  [346, 2],
  [347, 1],
  [349, 1],
  [350, 4],
  [351, 3],
  [352, 1],
  [353, 2],
  [360, 1],
  [368, 1],
  [369, 1],
  [370, 1],
  [371, 1],
  [372, 1],
  [373, 1],
  [374, 1],
  [375, 1],
  [376, 1],
  [377, 1],
  [378, 1],
  [382, 3],
  [385, 1],
  [392, 1],
  [393, 1],
  [396, 2],
  [397, 2],
  [398, 2],
  [399, 1],
  [400, 1],
  [401, 1],
  [402, 1],
  [403, 1],
  [404, 1],
  [405, 1],
  [406, 1],
  [407, 1],
  [408, 1],
  [409, 1],
  [410, 1],
  [414, 4],
  [415, 1],
  [416, 1],
  [417, 2],
  [420, 1],
  [421, 1],
  [422, 2],
  [424, 1],
  [425, 2],
  [426, 2],
  [427, 2],
  [428, 2],
  [430, 3],
  [438, 3],
  [439, 3],
  [440, 3],
  [443, 2],
  [444, 2],
  [445, 2],
  [446, 2],
  [447, 6],
  [448, 6],
  [449, 2],
  [450, 2],
  [464, 2],
  [476, 2],
  [479, 1],
  [480, 2]
];

/** `iftab` → argument count, for the fixed-arity functions. Absent means variadic. */
export const ARITY_BY_ID: ReadonlyMap<number, number> = new Map(FIXED_ARITY);

export const FUNCTION_NAME_BY_ID: ReadonlyMap<number, string> = new Map(FUNCTION_TABLE);
export const FUNCTION_ID_BY_NAME: ReadonlyMap<string, number> = new Map(
  FUNCTION_TABLE.map(([id, name]) => [name, id])
);

// =============================================================================
// Operand classes
// =============================================================================

/**
 * A function's return class and the class each of its parameters wants — `R`eference, `V`alue or `A`rray, and a
 * conversion rule — by `iftab`.
 *
 * **Why a table at all.** A reference token's class says what its *consumer* does with it: `SUM(A1:A3)` hands `SUM`
 * a reference (`PtgArea` 0x25), `ABS(A1-B1)` hands `-` two values (`PtgRef` 0x44), and `SUMPRODUCT(A1:A3*B1:B3)`
 * hands `*` two arrays (0x65). Written in the wrong class the formula still opens — and Excel evaluates `=A1+B1`
 * with two reference-class operands as `#VALUE!`. Which class a function wants for each argument is a property of
 * the function, so it has to be looked up.
 *
 * **Where it comes from.** The class letters are LibreOffice's (`sc/source/filter/oox/formulabase.cxx`), whose
 * export has been matched against Excel's own token classes for two decades; the specification documents the
 * classes per token but not per function parameter. The two-letter codes are: first letter `R`/`V` — whether a
 * reference argument is wanted as a reference or as a value (`mbValType`); second letter the conversion — `O`
 * original, `V` value, `A` array, `R` repeat the parent's, `X` repeat the parent's for references only. The last entry
 * repeats for every further argument; with the `p` flag (parameter pairs, `COUNTIFS`) the last *two* repeat. Where
 * LibreOffice's table has a row per file-format version the latest is taken: its BIFF2 row for `VLOOKUP` stops at
 * three parameters, and Excel writes the fourth, `range_lookup`, as a value (`44`) as its BIFF5 row says.
 *
 * Verified, not trusted: `excel/xlsb/__tests__/operand-classes.node.test.ts` re-encodes every formula of workbooks
 * Excel itself saved as XLSB and asserts the encoder reproduces Excel's tokens byte for byte.
 *
 * Where Excel's bytes disagree with LibreOffice's letters, Excel's win: `ROWS` and `COLUMNS` are `RO` there, but Excel
 * writes an expression argument in the array class (`ROWS(A1:A5+0)` is `65 … 03 41 4c 00`) and counts all five rows,
 * so they are `RA` here — `excel/xlsb/__tests__/data/formula-classes.excel.xlsb`.
 *
 * An `iftab` absent from here — the macro-sheet commands — is treated as returning a value from value arguments.
 * Flags: `v` volatile (the formula then starts with `PtgAttrSemi`), `p` parameter pairs.
 *
 * **`!` marks a parameter that only a reference can fill** (`RO!`). The class letters cannot say this: `RO` is also
 * `IF`'s branches, `CHOOSE`'s values, `VLOOKUP`'s table and `IFERROR`'s fallback, which take any value. Excel refuses
 * to accept a constant or a computed value in a reference-only parameter — `RANK(2,{2,2,1},0)` cannot be typed, and a
 * file holding it is one Excel will not open — so the writers refuse it too (`nonReferenceArgument` in the XLSB
 * operand-class assignment). Marked from Microsoft's function reference:
 *
 * - `RANK`, `RANK.EQ`, `RANK.AVG` `ref`: "A reference to a list of numbers."
 * - `OFFSET` `reference`: "Reference must refer to a cell or range of adjacent cells".
 * - `AREAS` `reference`: "A reference to a cell or range of cells and can refer to multiple areas."
 * - `ROW`, `COLUMN` `reference`: "The cell or range of cells for which you want the row number."
 * - `CELL` `reference`: "The cell that you want information about."
 * - `SUBTOTAL` `ref1`…: "The first named range or reference for which you want the subtotal."
 * - `FORMULATEXT` `reference`: "A reference to a cell or range of cells."
 * - `COUNTBLANK` `range`: "The range from which you want to count the blank cells."
 * - `COUNTIF`, `SUMIF`, `AVERAGEIF` `range` (and `sum_range`/`average_range`), `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`
 *   `criteria_range`s and `sum_range`/`average_range`: "The group of cells you want to count", "The range of cells that
 *   you want evaluated by criteria". These pages also say a range "can contain … arrays" — that is about what the cells
 *   hold; Excel refuses an array constant as the argument itself.
 *
 * Deliberately not marked: `AGGREGATE`, whose `ref1` is a reference in the reference form and "an array, an array
 * formula, or a reference" in the array form, chosen by the value of `function_num` — a per-parameter mark cannot
 * express that; the database functions, `SHEET`, `SHEETS` and `PHONETIC`, whose documentation admits a value or whose
 * behaviour is not established.
 */
const OPERAND_CLASSES: readonly (readonly [number, "R" | "V" | "A", string, string?])[] = [
  [0, "V", "RX"], // COUNT
  [1, "R", "VO RO"], // IF
  [2, "V", "VR"], // ISNA
  [3, "V", "VR"], // ISERROR
  [4, "V", "RX"], // SUM
  [5, "V", "RX"], // AVERAGE
  [6, "V", "RX"], // MIN
  [7, "V", "RX"], // MAX
  [8, "V", "RO!"], // ROW
  [9, "V", "RO!"], // COLUMN
  [10, "V", ""], // NA
  [11, "V", "VR RX"], // NPV
  [12, "V", "RX"], // STDEV
  [13, "V", "VR"], // DOLLAR
  [14, "V", "VR"], // FIXED
  [15, "V", "VR"], // SIN
  [16, "V", "VR"], // COS
  [17, "V", "VR"], // TAN
  [18, "V", "VR"], // ATAN
  [19, "V", ""], // PI
  [20, "V", "VR"], // SQRT
  [21, "V", "VR"], // EXP
  [22, "V", "VR"], // LN
  [23, "V", "VR"], // LOG10
  [24, "V", "VR"], // ABS
  [25, "V", "VR"], // INT
  [26, "V", "VR"], // SIGN
  [27, "V", "VR"], // ROUND
  [28, "V", "VR RA"], // LOOKUP
  [29, "R", "RA VV"], // INDEX
  [30, "V", "VR"], // REPT
  [31, "V", "VR"], // MID
  [32, "V", "VR"], // LEN
  [33, "V", "VR"], // VALUE
  [34, "V", ""], // TRUE
  [35, "V", ""], // FALSE
  [36, "V", "RX"], // AND
  [37, "V", "RX"], // OR
  [38, "V", "VR"], // NOT
  [39, "V", "VR"], // MOD
  [40, "V", "RO RR"], // DCOUNT
  [41, "V", "RO RR"], // DSUM
  [42, "V", "RO RR"], // DAVERAGE
  [43, "V", "RO RR"], // DMIN
  [44, "V", "RO RR"], // DMAX
  [45, "V", "RO RR"], // DSTDEV
  [46, "V", "RX"], // VAR
  [47, "V", "RO RR"], // DVAR
  [48, "V", "VR"], // TEXT
  // Excel writes the `const`/`stats` arguments as references (`LINEST(A1:A3,B1:B3,A1,B1)` is `… 24 … 24 … 42 04 31`),
  // where LibreOffice's table has values.
  [49, "A", "RA RA RA"], // LINEST
  [50, "A", "RA RA RA VV"], // TREND
  [51, "A", "RA RA RA"], // LOGEST
  [52, "A", "RA RA RA VV"], // GROWTH
  [55, "R", "RO"], // RETURN
  [56, "V", "VR"], // PV
  [57, "V", "VR"], // FV
  [58, "V", "VR"], // NPER
  [59, "V", "VR"], // PMT
  [60, "V", "VR"], // RATE
  [61, "V", "RA VR"], // MIRR
  [62, "V", "RA VR"], // IRR
  [63, "V", "", "v"], // RAND
  [64, "V", "VR RX RR"], // MATCH
  [65, "V", "VR"], // DATE
  [66, "V", "VR"], // TIME
  [67, "V", "VR"], // DAY
  [68, "V", "VR"], // MONTH
  [69, "V", "VR"], // YEAR
  [70, "V", "VR"], // WEEKDAY
  [71, "V", "VR"], // HOUR
  [72, "V", "VR"], // MINUTE
  [73, "V", "VR"], // SECOND
  [74, "V", "", "v"], // NOW
  [75, "V", "RO!"], // AREAS
  [76, "V", "RA"], // ROWS
  [77, "V", "RA"], // COLUMNS
  [78, "R", "RO! VR", "v"], // OFFSET
  [79, "R", "VR RO"], // ABSREF
  [82, "V", "VR"], // SEARCH
  [83, "A", "VO"], // TRANSPOSE
  [86, "V", "VX"], // TYPE
  [94, "R", ""], // ACTIVE.CELL
  [97, "V", "VR"], // ATAN2
  [98, "V", "VR"], // ASIN
  [99, "V", "VR"], // ACOS
  [100, "R", "VO RO"], // CHOOSE
  [101, "V", "VV RO RO VV"], // HLOOKUP
  [102, "V", "VV RO RO VV"], // VLOOKUP
  [105, "V", "RX"], // ISREF
  [109, "V", "VR"], // LOG
  [111, "V", "VR"], // CHAR
  [112, "V", "VR"], // LOWER
  [113, "V", "VR"], // UPPER
  [114, "V", "VR"], // PROPER
  [115, "V", "VR"], // LEFT
  [116, "V", "VR"], // RIGHT
  [117, "V", "VR"], // EXACT
  [118, "V", "VR"], // TRIM
  [119, "V", "VR"], // REPLACE
  [120, "V", "VR"], // SUBSTITUTE
  [121, "V", "VR"], // CODE
  [124, "V", "VR"], // FIND
  [125, "V", "VV RO!", "v"], // CELL
  [126, "V", "VR"], // ISERR
  [127, "V", "VR"], // ISTEXT
  [128, "V", "VR"], // ISNUMBER
  [129, "V", "VR"], // ISBLANK
  [130, "V", "RO"], // T
  [131, "V", "RO"], // N
  [140, "V", "VR"], // DATEVALUE
  [141, "V", "VR"], // TIMEVALUE
  [142, "V", "VR"], // SLN
  [143, "V", "VR"], // SYD
  [144, "V", "VR"], // DDB
  [148, "R", "VR", "v"], // INDIRECT
  [151, "V", "VR"], // ADD.BAR
  [152, "V", "VR RO RO VR"], // ADD.MENU
  [153, "V", "VR RO RO RO VR"], // ADD.COMMAND
  [162, "V", "VR"], // CLEAN
  [163, "V", "VA"], // MDETERM
  [164, "A", "VA"], // MINVERSE
  [165, "A", "VA"], // MMULT
  [167, "V", "VR"], // IPMT
  [168, "V", "VR"], // PPMT
  [169, "V", "RX"], // COUNTA
  [183, "V", "RX"], // PRODUCT
  [184, "V", "VR"], // FACT
  [189, "V", "RO RR"], // DPRODUCT
  [190, "V", "VR"], // ISNONTEXT
  [193, "V", "RX"], // STDEVP
  [194, "V", "RX"], // VARP
  [195, "V", "RO RR"], // DSTDEVP
  [196, "V", "RO RR"], // DVARP
  [197, "V", "VR"], // TRUNC
  [198, "V", "VR"], // ISLOGICAL
  [199, "V", "RO RR"], // DCOUNTA
  [204, "V", "VR"], // USDOLLAR
  [205, "V", "VR"], // FINDB
  [206, "V", "VR"], // SEARCHB
  [207, "V", "VR"], // REPLACEB
  [208, "V", "VR"], // LEFTB
  [209, "V", "VR"], // RIGHTB
  [210, "V", "VR"], // MIDB
  [211, "V", "VR"], // LENB
  [212, "V", "VR"], // ROUNDUP
  [213, "V", "VR"], // ROUNDDOWN
  [214, "V", "VR"], // ASC
  [215, "V", "VR"], // DBCS
  [216, "V", "VR RO! VR"], // RANK
  [219, "V", "VR"], // ADDRESS
  [220, "V", "VR"], // DAYS360
  [221, "V", "", "v"], // TODAY
  [222, "V", "VR"], // VDB
  [227, "V", "RX"], // MEDIAN
  [228, "V", "VA"], // SUMPRODUCT
  [229, "V", "VR"], // SINH
  [230, "V", "VR"], // COSH
  [231, "V", "VR"], // TANH
  [232, "V", "VR"], // ASINH
  [233, "V", "VR"], // ACOSH
  [234, "V", "VR"], // ATANH
  [235, "V", "RO RR"], // DGET
  [244, "V", "VR", "v"], // INFO
  [247, "V", "VR"], // DB
  [252, "A", "RA"], // FREQUENCY
  [261, "V", "VR"], // ERROR.TYPE
  [269, "V", "RX"], // AVEDEV
  [270, "V", "VR"], // BETADIST
  [271, "V", "VR"], // GAMMALN
  [272, "V", "VR"], // BETAINV
  [273, "V", "VR"], // BINOMDIST
  [274, "V", "VR"], // CHIDIST
  [275, "V", "VR"], // CHIINV
  [276, "V", "VR"], // COMBIN
  [277, "V", "VR"], // CONFIDENCE
  [278, "V", "VR"], // CRITBINOM
  [279, "V", "VR"], // EVEN
  [280, "V", "VR"], // EXPONDIST
  [281, "V", "VR"], // FDIST
  [282, "V", "VR"], // FINV
  [283, "V", "VR"], // FISHER
  [284, "V", "VR"], // FISHERINV
  [285, "V", "VR"], // FLOOR
  [286, "V", "VR"], // GAMMADIST
  [287, "V", "VR"], // GAMMAINV
  [288, "V", "VR"], // CEILING
  [289, "V", "VR"], // HYPGEOMDIST
  [290, "V", "VR"], // LOGNORMDIST
  [291, "V", "VR"], // LOGINV
  [292, "V", "VR"], // NEGBINOMDIST
  [293, "V", "VR"], // NORMDIST
  [294, "V", "VR"], // NORMSDIST
  [295, "V", "VR"], // NORMINV
  [296, "V", "VR"], // NORMSINV
  [297, "V", "VR"], // STANDARDIZE
  [298, "V", "VR"], // ODD
  [299, "V", "VR"], // PERMUT
  [300, "V", "VR"], // POISSON
  [301, "V", "VR"], // TDIST
  [302, "V", "VR"], // WEIBULL
  [303, "V", "VA"], // SUMXMY2
  [304, "V", "VA"], // SUMX2MY2
  [305, "V", "VA"], // SUMX2PY2
  [306, "V", "VA"], // CHITEST
  [307, "V", "VA"], // CORREL
  [308, "V", "VA"], // COVAR
  [309, "V", "VR VA"], // FORECAST
  [310, "V", "VA"], // FTEST
  [311, "V", "VA"], // INTERCEPT
  [312, "V", "VA"], // PEARSON
  [313, "V", "VA"], // RSQ
  [314, "V", "VA"], // STEYX
  [315, "V", "VA"], // SLOPE
  [316, "V", "VA VA VR"], // TTEST
  [317, "V", "VA VA VR"], // PROB
  [318, "V", "RX"], // DEVSQ
  [319, "V", "RX"], // GEOMEAN
  [320, "V", "RX"], // HARMEAN
  [321, "V", "RX"], // SUMSQ
  [322, "V", "RX"], // KURT
  [323, "V", "RX"], // SKEW
  [324, "V", "RX VR"], // ZTEST
  [325, "V", "RX VR"], // LARGE
  [326, "V", "RX VR"], // SMALL
  [327, "V", "RX VR"], // QUARTILE
  [328, "V", "RX VR"], // PERCENTILE
  [329, "V", "RX VR VR_E"], // PERCENTRANK
  [330, "V", "VA"], // MODE
  [331, "V", "RX VR"], // TRIMMEAN
  [332, "V", "VR"], // TINV
  [336, "V", "VR"], // CONCATENATE
  [337, "V", "VR"], // POWER
  [342, "V", "VR"], // RADIANS
  [343, "V", "VR"], // DEGREES
  [344, "V", "VR RO!"], // SUBTOTAL
  [345, "V", "RO! VR RO!"], // SUMIF
  [346, "V", "RO! VR"], // COUNTIF
  [347, "V", "RO!"], // COUNTBLANK
  [350, "V", "VR"], // ISPMT
  [351, "V", "VR"], // DATEDIF
  [352, "V", "VR"], // DATESTRING
  [353, "V", "VR"], // NUMBERSTRING
  [354, "V", "VR"], // ROMAN
  [358, "V", "RR RR VR VR", "p"], // GETPIVOTDATA
  [359, "V", "VV VO"], // HYPERLINK
  [360, "V", "RO"], // PHONETIC
  [361, "V", "RX"], // AVERAGEA
  [362, "V", "RX"], // MAXA
  [363, "V", "RX"], // MINA
  [364, "V", "RX"], // STDEVPA
  [365, "V", "RX"], // VARPA
  [366, "V", "RX"], // STDEVA
  [367, "V", "RX"], // VARA
  [368, "V", "VR"], // BAHTTEXT
  [369, "V", "VR"], // THAIDAYOFWEEK
  [370, "V", "VR"], // THAIDIGIT
  [371, "V", "VR"], // THAIMONTHOFYEAR
  [372, "V", "VR"], // THAINUMSOUND
  [373, "V", "VR"], // THAINUMSTRING
  [374, "V", "VR"], // THAISTRINGLENGTH
  [375, "V", "VR"], // ISTHAIDIGIT
  [376, "V", "VR"], // ROUNDBAHTDOWN
  [377, "V", "VR"], // ROUNDBAHTUP
  [378, "V", "VR"], // THAIYEAR
  [379, "A", "VR VR RO"], // RTD
  [380, "V", "VR RX"], // CUBEVALUE
  [381, "V", "VR RX VR"], // CUBEMEMBER
  [382, "V", "VR"], // CUBEMEMBERPROPERTY
  [383, "V", "VR"], // CUBERANKEDMEMBER
  [384, "V", "RR"], // HEX2BIN
  [385, "V", "RR"], // HEX2DEC
  [386, "V", "RR"], // HEX2OCT
  [387, "V", "RR"], // DEC2BIN
  [388, "V", "RR"], // DEC2HEX
  [389, "V", "RR"], // DEC2OCT
  [390, "V", "RR"], // OCT2BIN
  [391, "V", "RR"], // OCT2HEX
  [392, "V", "RR"], // OCT2DEC
  [393, "V", "RR"], // BIN2DEC
  [394, "V", "RR"], // BIN2OCT
  [395, "V", "RR"], // BIN2HEX
  [396, "V", "RR"], // IMSUB
  [397, "V", "RR"], // IMDIV
  [398, "V", "RR"], // IMPOWER
  [399, "V", "RR"], // IMABS
  [400, "V", "RR"], // IMSQRT
  [401, "V", "RR"], // IMLN
  [402, "V", "RR"], // IMLOG2
  [403, "V", "RR"], // IMLOG10
  [404, "V", "RR"], // IMSIN
  [405, "V", "RR"], // IMCOS
  [406, "V", "RR"], // IMEXP
  [407, "V", "RR"], // IMARGUMENT
  [408, "V", "RR"], // IMCONJUGATE
  [409, "V", "RR"], // IMAGINARY
  [410, "V", "RR"], // IMREAL
  [411, "V", "RR"], // COMPLEX
  [412, "V", "RX"], // IMSUM
  [413, "V", "RX"], // IMPRODUCT
  [414, "V", "RR RR RR RX"], // SERIESSUM
  [415, "V", "RR"], // FACTDOUBLE
  [416, "V", "RR"], // SQRTPI
  [417, "V", "RR"], // QUOTIENT
  [418, "V", "RR"], // DELTA
  [419, "V", "RR"], // GESTEP
  [420, "V", "RR"], // ISEVEN
  [421, "V", "RR"], // ISODD
  [422, "V", "RR"], // MROUND
  [423, "V", "RR"], // ERF
  [424, "V", "RR"], // ERFC
  [425, "V", "RR"], // BESSELJ
  [426, "V", "RR"], // BESSELK
  [427, "V", "RR"], // BESSELY
  [428, "V", "RR"], // BESSELI
  [429, "V", "RX RX RR"], // XIRR
  [430, "V", "RR RX RX"], // XNPV
  [431, "V", "RR"], // PRICEMAT
  [432, "V", "RR"], // YIELDMAT
  [433, "V", "RR"], // INTRATE
  [434, "V", "RR"], // RECEIVED
  [435, "V", "RR"], // DISC
  [436, "V", "RR"], // PRICEDISC
  [437, "V", "RR"], // YIELDDISC
  [438, "V", "RR"], // TBILLEQ
  [439, "V", "RR"], // TBILLPRICE
  [440, "V", "RR"], // TBILLYIELD
  [441, "V", "RR"], // PRICE
  [442, "V", "RR"], // YIELD
  [443, "V", "RR"], // DOLLARDE
  [444, "V", "RR"], // DOLLARFR
  [445, "V", "RR"], // NOMINAL
  [446, "V", "RR"], // EFFECT
  [447, "V", "RR"], // CUMPRINC
  [448, "V", "RR"], // CUMIPMT
  [449, "V", "RR"], // EDATE
  [450, "V", "RR"], // EOMONTH
  [451, "V", "RR"], // YEARFRAC
  [452, "V", "RR"], // COUPDAYBS
  [453, "V", "RR"], // COUPDAYS
  [454, "V", "RR"], // COUPDAYSNC
  [455, "V", "RR"], // COUPNCD
  [456, "V", "RR"], // COUPNUM
  [457, "V", "RR"], // COUPPCD
  [458, "V", "RR"], // DURATION
  [459, "V", "RR"], // MDURATION
  [460, "V", "RR"], // ODDLPRICE
  [461, "V", "RR"], // ODDLYIELD
  [462, "V", "RR"], // ODDFPRICE
  [463, "V", "RR"], // ODDFYIELD
  [464, "V", "RR", "v"], // RANDBETWEEN
  [465, "V", "RR"], // WEEKNUM
  [466, "V", "RR"], // AMORDEGRC
  [467, "V", "RR"], // AMORLINC
  [469, "V", "RR"], // ACCRINT
  [470, "V", "RR"], // ACCRINTM
  [471, "V", "RR RR RX C"], // WORKDAY
  [472, "V", "RR RR RX C"], // NETWORKDAYS
  [473, "V", "RX"], // GCD
  [474, "V", "RX"], // MULTINOMIAL
  [475, "V", "RX"], // LCM
  [476, "V", "RR RX"], // FVSCHEDULE
  [477, "V", "VR"], // CUBEKPIMEMBER
  [478, "V", "VR RX VR"], // CUBESET
  [479, "V", "VR"], // CUBESETCOUNT
  [480, "V", "VO RO"], // IFERROR
  [481, "V", "RO! VR", "p"], // COUNTIFS
  [482, "V", "RO! RO! VR", "p"], // SUMIFS
  [483, "V", "RO! VR RO!"], // AVERAGEIF
  [484, "V", "RO! RO! VR", "p"] // AVERAGEIFS
];

/** One parameter's class rule; see {@link OPERAND_CLASSES}. */
export interface ParamClass {
  /** A reference argument is wanted as a value (`V…`) rather than as a reference (`R…`). */
  readonly valueType: boolean;
  readonly conversion: "original" | "value" | "array" | "repeat" | "repeatReference";
  /** Only a reference can fill this parameter — `!` in {@link OPERAND_CLASSES}. */
  readonly referenceOnly?: true;
}

/** A function's operand-class description. */
export interface FunctionClasses {
  readonly returns: "R" | "V" | "A";
  readonly params: readonly ParamClass[];
  readonly volatile: boolean;
  readonly pairs: boolean;
}

const CONVERSION_BY_LETTER = {
  O: "original",
  V: "value",
  A: "array",
  R: "repeat",
  X: "repeatReference"
} as const;

function paramClass(code: string): ParamClass {
  return {
    valueType: code[0] === "V",
    conversion: CONVERSION_BY_LETTER[code[1] as keyof typeof CONVERSION_BY_LETTER],
    ...(code[2] === "!" ? { referenceOnly: true as const } : {})
  };
}

const CLASSES_BY_ID: ReadonlyMap<number, FunctionClasses> = new Map(
  OPERAND_CLASSES.map(([id, returns, params, flags]) => [
    id,
    {
      returns,
      params: params === "" ? [] : params.split(" ").map(paramClass),
      volatile: flags?.includes("v") ?? false,
      pairs: flags?.includes("p") ?? false
    }
  ])
);

/** What a value-returning function of value arguments looks like — the default for an `iftab` with no entry. */
const DEFAULT_CLASSES: FunctionClasses = {
  returns: "V",
  params: [{ valueType: true, conversion: "repeat" }],
  volatile: false,
  pairs: false
};

/**
 * A future function's classes when nothing more is known: every argument a reference, as Excel's own
 * `_xlfn.XLOOKUP(2,A1:A3,B1:B3)`, `_xlfn.CONCAT(A1,B1)`, `_xlfn.MAXIFS(…)` and `_xlfn.VSTACK(…)` carry them
 * (0x25, 0x24), and a value back. A reference argument where a value is wanted computes the same thing; the
 * reverse is what fails.
 */
export const FUTURE_FUNCTION_CLASSES: FunctionClasses = {
  returns: "V",
  params: [{ valueType: false, conversion: "original" }],
  volatile: false,
  pairs: false
};

/**
 * Future functions whose classes Excel was observed to write differently from {@link FUTURE_FUNCTION_CLASSES}, by
 * name without the `_xlfn.`/`_xlws.` prefix — each line read off Excel's own XLSB of the call (`fx10` probes, and the
 * `formula-edges` fixture for `SORT`). No published table gives these, and LibreOffice's disagrees with Excel's
 * bytes for most of them, so only what was observed is listed.
 *
 * - value arguments: `_xlfn.SEQUENCE(A1)`, `_xlfn.DAYS(A1,B1)`, `_xlfn.CEILING.MATH(A1)` are `44 …`;
 * - a value first, references after: `_xlfn.IFNA(A1,0)`, `_xlfn.RANK.EQ(A1,A1:A3)` (`44 … 25`), and `_xlfn.IFS`
 *   in condition/value pairs (`44 … 24`);
 * - array arguments: `_xlfn._xlws.SORT(A1:A3)`, `_xlfn.UNIQUE`, `_xlfn.SORTBY`, `_xlfn.MODE.SNGL`,
 *   `_xlfn.FORECAST.LINEAR` are `65 …`, and `_xlfn._xlws.FILTER(A1:A3,B1:B3>2)` is `25 … 65`;
 * - volatile: `_xlfn.SHEET(A1)` and `_xlfn.FORMULATEXT(A1)` start with `PtgAttrSemi`;
 * - from Excel's XLSB of `future-functions.xlsx` (`excel/xlsb/__tests__/future-function-tokens.node.test.ts`): array
 *   arguments for `_xlfn.CHISQ.TEST`, `_xlfn.COVARIANCE.P`/`.S`, `_xlfn.F.TEST`, `_xlfn.T.TEST` (`7b … 7b`), and for the first
 *   argument of `_xlfn.EXPAND`, `_xlfn.WRAPCOLS` and `_xlfn.WRAPROWS`; `_xlfn.RANDARRAY` starts with `PtgAttrSemi`.
 *   The scalar arguments after them are values, as LibreOffice lists them; Excel's samples pass constants there.
 * - `_xlfn.GROUPBY` and `_xlfn.PIVOTBY`: the values argument (second and third) and the function after it are arrays —
 *   `65 …` in `eta-lambda.excel.xlsb` (`excel/xlsb/__tests__/eta-lambda.node.test.ts`), and `7b …` then a defined
 *   name as `63 …` in Excel's XLSB of `future-functions.xlsx`; the fields before them are references, `25 …`. A function
 *   passed by name is always `43`, whatever its argument's class — see `etaStubName` in the XLSB codec.
 */
const FUTURE_FUNCTION_OVERRIDES: ReadonlyMap<string, readonly [string, string?]> = new Map([
  ["SEQUENCE", ["VR"]],
  ["DAYS", ["VR"]],
  ["CEILING.MATH", ["VR"]],
  ["FLOOR.MATH", ["VR"]],
  ["NORM.S.DIST", ["VR"]],
  ["IFNA", ["VR RO"]],
  ["RANK.EQ", ["VR RO!"]],
  ["RANK.AVG", ["VR RO!"]],
  ["SWITCH", ["VR RO"]],
  ["IFS", ["VR RO", "p"]],
  ["SORT", ["VA"]],
  ["UNIQUE", ["VA"]],
  ["SORTBY", ["VA"]],
  ["MODE.SNGL", ["VA"]],
  ["MODE.MULT", ["VA"]],
  ["FORECAST.LINEAR", ["VA"]],
  ["FILTER", ["RO VA"]],
  ["CHISQ.TEST", ["VA"]],
  ["COVARIANCE.P", ["VA"]],
  ["COVARIANCE.S", ["VA"]],
  ["F.TEST", ["VA"]],
  ["T.TEST", ["VA VA VR"]],
  ["EXPAND", ["VA VR"]],
  ["WRAPCOLS", ["VA VR"]],
  ["WRAPROWS", ["VA VR"]],
  ["RANDARRAY", ["VR", "v"]],
  ["SHEET", ["RO", "v"]],
  ["SHEETS", ["RO", "v"]],
  ["FORMULATEXT", ["RO!", "v"]],
  ["GROUPBY", ["RO VA VA RO"]],
  ["PIVOTBY", ["RO RO VA VA RO"]]
]);

/**
 * Future functions that return something other than a value, by bare name — LibreOffice's return classes for them
 * (`formulabase.cxx`), plus `SORTBY` and `MODE.MULT`, which return arrays as well and which LibreOffice lists as
 * values.
 *
 * **Why it matters.** The function token's class is what its consumer gets. Written in the value class, a call that
 * returns an array is intersected down to one cell where it is an argument: Excel reads
 * `SUM(_xlfn._xlws.FILTER(B2:B5,B2:B5>1))` as `SUM(@FILTER(…))` and computes 2 where the answer is 9, and
 * `ROWS(_xlfn.SEQUENCE(5))` as 1. In the array class it is left whole. As the whole formula the root's value
 * conversion still takes it to the value class, which is what Excel writes there.
 *
 * The reference-returning ones are read off Excel's own XLSB
 * (`excel/xlsb/__tests__/future-function-tokens.node.test.ts`):
 * `SUM(_xlfn.TRIMRANGE(D!A1:A20))`, `SUM(_xlfn.ANCHORARRAY(D!H1))`, `SUM(_xlfn.TAKE(…))` and `SUM(_xlfn.DROP(…))`
 * are `22 02 ff 00`. In the value class Excel shows `SUM(@TRIMRANGE(…))` and sums one cell.
 */
const FUTURE_FUNCTION_RETURNS: ReadonlyMap<string, "R" | "A"> = new Map([
  ...(
    [
      "BYCOL",
      "BYROW",
      "CHOOSECOLS",
      "CHOOSEROWS",
      "EXPAND",
      "FILTER",
      "GROUPBY",
      "HSTACK",
      "MAKEARRAY",
      "MAP",
      "MODE.MULT",
      "MUNIT",
      "PIVOTBY",
      "RANDARRAY",
      "SCAN",
      "SEQUENCE",
      "SORT",
      "SORTBY",
      "TEXTAFTER",
      "TEXTBEFORE",
      "TEXTSPLIT",
      "TOCOL",
      "TOROW",
      "UNIQUE",
      "VSTACK",
      "WRAPCOLS",
      "WRAPROWS"
    ] as const
  ).map(name => [name, "A"] as const),
  // `REDUCE` returns whatever its lambda accumulates, which may be a reference — Excel's own XLSB of
  // `_xlfn.LET(_xlpm.sq,…,_xlfn.REDUCE(…))` (`data/let-oracle.excel.xlsb`) calls it `22 04 ff 00` as `LET`'s result.
  // Written as `42`, Excel showed `@REDUCE(…)` and saved it as `_xlfn.SINGLE(_xlfn.REDUCE(…))`. As the whole formula
  // the root's value conversion still makes it `42`, which is what Excel writes there.
  // `SINGLE`, the implicit-intersection operator `@` as a call, intersects a reference to a reference: Excel's own XLSB
  // calls it `22 02 ff 00` as `SUM`'s or `INDEX`'s reference argument and `62` as `SUMPRODUCT`'s array one
  // (`data/single-cells.excel.xlsb`).
  ...(["ANCHORARRAY", "DROP", "REDUCE", "SINGLE", "TAKE", "TRIMRANGE", "XLOOKUP"] as const).map(
    name => [name, "R"] as const
  )
]);

/**
 * A call through a `LET`/`LAMBDA` parameter, `_xlpm.f(1,2)`: what it returns is whatever the lambda bound to it does,
 * so it is a reference where one is wanted. Excel's own XLSB of `_xlfn.LET(_xlpm.f,…,_xlpm.f(1,2))` calls it with
 * `22 03 ff 00`; in the value class `SUM(f(3))` shows as `SUM(@f(3))` and sums the first element only.
 */
const LAMBDA_PARAMETER_CLASSES: FunctionClasses = { ...FUTURE_FUNCTION_CLASSES, returns: "R" };

/** `name` without any `_xlfn.` / `_xlws.` prefix, upper-cased. */
function bareFunctionName(name: string): string {
  let upper = name.toUpperCase();
  for (const prefix of ["_XLFN.", "_XLWS."]) {
    if (upper.startsWith(prefix)) {
      upper = upper.slice(prefix.length);
    }
  }
  return upper;
}

/** The classes for a call through a future-function stub, by the call's name, with or without its prefixes. */
export function futureFunctionClasses(name: string): FunctionClasses {
  // A call through a `LET` / `LAMBDA` parameter as a file spells it, `_xlpm.f(…)`.
  if (name.toLowerCase().startsWith("_xlpm.")) {
    return LAMBDA_PARAMETER_CLASSES;
  }
  const bare = bareFunctionName(name);
  const returns = FUTURE_FUNCTION_RETURNS.get(bare) ?? "V";
  const override = FUTURE_FUNCTION_OVERRIDES.get(bare);
  if (override === undefined) {
    return returns === "V" ? FUTURE_FUNCTION_CLASSES : { ...FUTURE_FUNCTION_CLASSES, returns };
  }
  const [params, flags] = override;
  return {
    returns,
    params: params.split(" ").map(paramClass),
    volatile: flags?.includes("v") ?? false,
    pairs: flags?.includes("p") ?? false
  };
}

/**
 * The bare upper-case names of the functions with a reference-only parameter, so a writer can skip parsing a formula
 * that calls none of them.
 */
export const REFERENCE_ONLY_FUNCTIONS: ReadonlySet<string> = new Set([
  ...FUNCTION_TABLE.filter(([id]) =>
    CLASSES_BY_ID.get(id)?.params.some(param => param.referenceOnly)
  ).map(([, name]) => name),
  ...[...FUTURE_FUNCTION_OVERRIDES]
    .filter(([, [params]]) => params.includes("!"))
    .map(([name]) => name)
]);

/** The operand classes for `iftab`. */
export function functionClasses(id: number): FunctionClasses {
  return CLASSES_BY_ID.get(id) ?? DEFAULT_CLASSES;
}

/** The class rule for argument `index` of a call with `classes`: the last entry repeats, or the last two for pairs. */
export function paramClassAt(classes: FunctionClasses, index: number): ParamClass {
  const { params } = classes;
  if (params.length === 0) {
    return DEFAULT_CLASSES.params[0]!;
  }
  if (index < params.length) {
    return params[index]!;
  }
  if (classes.pairs && params.length >= 2) {
    return params[params.length - 2 + ((index - params.length) % 2)]!;
  }
  return params[params.length - 1]!;
}

/**
 * The two function tables, for the cross-checks that keep them from drifting apart.
 *
 * **They are two transcriptions of one specification table, related only by a bare number.** `FUNCTION_TABLE` maps an
 * `iftab` to a name and `FIXED_ARITY` maps the same `iftab` to a parameter count; the first came from the specification's
 * name column and the second from a syntax judgement on its `Parameters` column. Nothing connected them, and both the
 * decoder and the encoder look each up separately, so an id present in one and absent from the other is a silent
 * degradation rather than an error.
 *
 * That matters more here than duplication usually does, because of how an arity mistake fails: `PtgFunc` has nowhere to
 * record the real argument count, so a wrong arity produces a file Excel opens without complaint and evaluates
 * differently — and a round trip cannot see it, since the encoder and decoder share the same wrong entry. It is the exact
 * shape of the twelve `iftab` defects this table's own history records.
 *
 * Exported so `excel/xlsb/__tests__/function-table.test.ts` can assert what no type can: that every id in one appears in
 * the other,
 * that neither has a duplicate, and that the counts are what they are.
 */
export const FUNCTION_TABLES = {
  names: FUNCTION_TABLE,
  arities: FIXED_ARITY,
  operandClasses: OPERAND_CLASSES
} as const;
