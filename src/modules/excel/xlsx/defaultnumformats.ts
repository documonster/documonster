interface NumFormat {
  f?: string;
  "zh-tw"?: string;
  "zh-cn"?: string;
  "ja-jp"?: string;
  "ko-kr"?: string;
  "th-th"?: string;
}

const defaultNumFormats: { [key: number]: NumFormat } = {
  0: { f: "General" },
  1: { f: "0" },
  2: { f: "0.00" },
  3: { f: "#,##0" },
  4: { f: "#,##0.00" },
  9: { f: "0%" },
  10: { f: "0.00%" },
  11: { f: "0.00E+00" },
  12: { f: "# ?/?" },
  13: { f: "# ??/??" },
  14: { f: "mm-dd-yy" },
  15: { f: "d-mmm-yy" },
  16: { f: "d-mmm" },
  17: { f: "mmm-yy" },
  18: { f: "h:mm AM/PM" },
  19: { f: "h:mm:ss AM/PM" },
  20: { f: "h:mm" },
  21: { f: "h:mm:ss" },
  22: { f: "m/d/yy h:mm" },

  27: {
    "zh-tw": "[$-404]e/m/d",
    "zh-cn": 'yyyy"年"m"月"',
    "ja-jp": "[$-411]ge.m.d",
    "ko-kr": 'yyyy"年" mm"月" dd"日"'
  },
  28: {
    "zh-tw": '[$-404]e"年"m"月"d"日"',
    "zh-cn": 'm"月"d"日"',
    "ja-jp": '[$-411]ggge"年"m"月"d"日"',
    "ko-kr": "mm-dd"
  },
  29: {
    "zh-tw": '[$-404]e"年"m"月"d"日"',
    "zh-cn": 'm"月"d"日"',
    "ja-jp": '[$-411]ggge"年"m"月"d"日"',
    "ko-kr": "mm-dd"
  },
  30: { "zh-tw": "m/d/yy ", "zh-cn": "m-d-yy", "ja-jp": "m/d/yy", "ko-kr": "mm-dd-yy" },
  31: {
    "zh-tw": 'yyyy"年"m"月"d"日"',
    "zh-cn": 'yyyy"年"m"月"d"日"',
    "ja-jp": 'yyyy"年"m"月"d"日"',
    "ko-kr": 'yyyy"년" mm"월" dd"일"'
  },
  32: {
    "zh-tw": 'hh"時"mm"分"',
    "zh-cn": 'h"时"mm"分"',
    "ja-jp": 'h"時"mm"分"',
    "ko-kr": 'h"시" mm"분"'
  },
  33: {
    "zh-tw": 'hh"時"mm"分"ss"秒"',
    "zh-cn": 'h"时"mm"分"ss"秒"',
    "ja-jp": 'h"時"mm"分"ss"秒"',
    "ko-kr": 'h"시" mm"분" ss"초"'
  },
  34: {
    "zh-tw": '上午/下午 hh"時"mm"分"',
    "zh-cn": '上午/下午 h"时"mm"分"',
    "ja-jp": 'yyyy"年"m"月"',
    "ko-kr": "yyyy-mm-dd"
  },
  35: {
    "zh-tw": '上午/下午 hh"時"mm"分"ss"秒"',
    "zh-cn": '上午/下午 h"时"mm"分"ss"秒"',
    "ja-jp": 'm"月"d"日"',
    "ko-kr": "yyyy-mm-dd"
  },
  36: {
    "zh-tw": "[$-404]e/m/d",
    "zh-cn": 'yyyy"年"m"月"',
    "ja-jp": "[$-411]ge.m.d",
    "ko-kr": 'yyyy"年" mm"月" dd"日"'
  },

  37: { f: "#,##0 ;(#,##0)" },
  38: { f: "#,##0 ;[Red](#,##0)" },
  39: { f: "#,##0.00;(#,##0.00)" },
  40: { f: "#,##0.00;[Red](#,##0.00)" },
  // **41–44, the accounting formats, are the locale's**, as the currencies 5–8 are: ECMA-376 leaves them out of its
  // table of built-in codes, and Excel shows them in the form of the reader's locale — `$` in an English one, `¥` in
  // Chinese. Excel writes each with a `<numFmt>` declaring that form, which a read takes as it is; the forms a display
  // gives them are in `display-locale.ts`.
  //
  // They once sat here as the codes a Chinese Excel wrote, under `f`. That made a bare id 42 read as a `￥` format
  // whatever the file's locale, and a `￥` code set by a caller written as id 42, which an English Excel shows in
  // dollars. Without an `f` the bare id is kept instead (`kept-num-fmt.ts`) and a code is written as the code.
  45: { f: "mm:ss" },
  46: { f: "[h]:mm:ss" },
  // ECMA-376's table prints `mmss.0`, a misprint: Excel shows and writes id 47 as `mm:ss.0` (measured), and
  // `mmss.0` would show 45000.75 as `0000.0` where Excel shows `00:00.0`.
  47: { f: "mm:ss.0" },
  48: { f: "##0.0E+0" },
  49: { f: "@" },

  50: {
    "zh-tw": "[$-404]e/m/d",
    "zh-cn": 'yyyy"年"m"月"',
    "ja-jp": "[$-411]ge.m.d",
    "ko-kr": 'yyyy"年" mm"月" dd"日"'
  },
  51: {
    "zh-tw": '[$-404]e"年"m"月"d"日"',
    "zh-cn": 'm"月"d"日"',
    "ja-jp": '[$-411]ggge"年"m"月"d"日"',
    "ko-kr": "mm-dd"
  },
  52: {
    "zh-tw": '上午/下午 hh"時"mm"分"',
    "zh-cn": 'yyyy"年"m"月"',
    "ja-jp": 'yyyy"年"m"月"',
    "ko-kr": "yyyy-mm-dd"
  },
  53: {
    "zh-tw": '上午/下午 hh"時"mm"分"ss"秒"',
    "zh-cn": 'm"月"d"日"',
    "ja-jp": 'm"月"d"日"',
    "ko-kr": "yyyy-mm-dd"
  },
  54: {
    "zh-tw": '[$-404]e"年"m"月"d"日"',
    "zh-cn": 'm"月"d"日"',
    "ja-jp": '[$-411]ggge"年"m"月"d"日"',
    "ko-kr": "mm-dd"
  },
  55: {
    "zh-tw": '上午/下午 hh"時"mm"分"',
    "zh-cn": '上午/下午 h"时"mm"分"',
    "ja-jp": 'yyyy"年"m"月"',
    "ko-kr": "yyyy-mm-dd"
  },
  56: {
    "zh-tw": '上午/下午 hh"時"mm"分"ss"秒"',
    "zh-cn": '上午/下午 h"时"mm"分"ss"秒"',
    "ja-jp": 'm"月"d"日"',
    "ko-kr": "yyyy-mm-dd"
  },
  57: {
    "zh-tw": "[$-404]e/m/d",
    "zh-cn": 'yyyy"年"m"月"',
    "ja-jp": "[$-411]ge.m.d",
    "ko-kr": 'yyyy"年" mm"月" dd"日"'
  },
  58: {
    "zh-tw": '[$-404]e"年"m"月"d"日"',
    "zh-cn": 'm"月"d"日"',
    "ja-jp": '[$-411]ggge"年"m"月"d"日"',
    "ko-kr": "mm-dd"
  },

  59: { "th-th": "t0" },
  60: { "th-th": "t0.00" },
  61: { "th-th": "t#,##0" },
  62: { "th-th": "t#,##0.00" },
  67: { "th-th": "t0%" },
  68: { "th-th": "t0.00%" },
  69: { "th-th": "t# ?/?" },
  70: { "th-th": "t# ??/??" },

  71: { "th-th": "ว/ด/ปปปป" },
  72: { "th-th": "ว-ดดด-ปป" },
  73: { "th-th": "ว-ดดด" },
  74: { "th-th": "ดดด-ปป" },
  75: { "th-th": "ช:นน" },
  76: { "th-th": "ช:นน:ทท" },
  77: { "th-th": "ว/ด/ปปปป ช:นน" },
  78: { "th-th": "นน:ทท" },
  79: { "th-th": "[ช]:นน:ทท" },
  80: { "th-th": "นน:ทท.0" },
  81: { "th-th": "d/m/bb" }
};

export { defaultNumFormats };
