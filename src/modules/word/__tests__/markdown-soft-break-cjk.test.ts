/**
 * What a soft line break becomes, per script.
 *
 * CommonMark turns a soft break into a space. That is right for a script that marks
 * word boundaries with one and wrong for a script that does not: a Chinese paragraph
 * hard-wrapped in the source — which is how nearly every hand-written CJK Markdown
 * file looks, this repository's own `README_zh.md` included — came out with a space at
 * every wrap point, in the middle of words. Pandoc addresses the same problem with its
 * `east_asian_line_breaks` extension.
 *
 * Found by converting this repository's own translated docs to DOCX and PDF and
 * reading the result, which is the only way this class of defect surfaces: every
 * structural assertion passed, and the text was simply wrong.
 */

import { isSpacelessScript } from "@utils/cjk";
import { markdownToDocx } from "@word/convert/markdown/markdown-import";
import { describe, expect, it } from "vitest";

/** The visible text of the first paragraph, with breaks marked. */
async function firstParagraph(markdown: string): Promise<string> {
  const doc = await markdownToDocx(markdown);
  const paragraph = doc.body[0] as {
    children?: { content?: { type: string; text?: string }[] }[];
  };
  return (paragraph.children ?? [])
    .flatMap(child => child.content ?? [])
    .map(node => (node.type === "text" ? (node.text ?? "") : `<${node.type}>`))
    .join("");
}

describe("soft line breaks across scripts", () => {
  it("keeps the space for a space-separated script", async () => {
    expect(await firstParagraph("the quick brown\nfox jumps")).toBe("the quick brown fox jumps");
  });

  it("drops the space between Chinese characters", async () => {
    expect(await firstParagraph("要理解这个\n系统读本文")).toBe("要理解这个系统读本文");
  });

  it("drops the space between Japanese characters", async () => {
    expect(await firstParagraph("システムを理解\nするには")).toBe("システムを理解するには");
  });

  it("keeps the space between Korean words", async () => {
    // Korean *does* separate words with spaces, so dropping this one would run two
    // words together. This is the case that rules out reusing `isCjkBreakable`, which
    // includes Hangul because the question it answers is a different one.
    expect(await firstParagraph("한국어를\n이해하려면")).toBe("한국어를 이해하려면");
  });

  it("keeps the space where the scripts meet", async () => {
    // Only a break with a spaceless script on *both* sides loses its space.
    expect(await firstParagraph("中文\nEnglish")).toBe("中文 English");
    expect(await firstParagraph("English\n中文")).toBe("English 中文");
  });

  it("drops the space after full-width punctuation", async () => {
    expect(await firstParagraph("结束。\n下一句")).toBe("结束。下一句");
    expect(await firstParagraph("（注）\n继续")).toBe("（注）继续");
  });

  it("leaves an explicit hard break alone", async () => {
    // Two trailing spaces mean a real line break, which is not this rule's business.
    expect(await firstParagraph("中文  \n下一行")).toBe("中文<break>下一行");
    expect(await firstParagraph("中文\\\n下一行")).toBe("中文<break>下一行");
  });

  it("keeps the space when the break follows markup", async () => {
    // After emphasis the character beside the break is markup, not prose, so there is
    // nothing reliable to judge and the space is kept — a spurious space is a far
    // smaller defect than two words run together.
    expect(await firstParagraph("**粗体**\n中文")).toBe("粗体 中文");
  });

  it("classifies scripts by whether they separate words with a space", () => {
    // `isSpacelessScript` is derived from `isCjkBreakable` minus Hangul rather than
    // restating its ranges. Written out separately it drifted immediately: the copy was
    // missing Yi and the halfwidth CJK punctuation. These pin the derivation, including
    // the Hangul blocks that are the whole reason the two predicates differ.
    for (const [label, cp] of [
      ["Han", 0x4e2d],
      ["Hiragana", 0x3072],
      ["Katakana", 0x30e9],
      ["ideographic full stop", 0x3002],
      ["fullwidth parenthesis", 0xff08],
      ["halfwidth ideographic full stop", 0xff61],
      ["halfwidth katakana", 0xff71],
      ["Yi syllable", 0xa000],
      ["Han beyond the BMP", 0x20000]
    ] as [string, number][]) {
      expect(isSpacelessScript(cp), label).toBe(true);
    }
    for (const [label, cp] of [
      ["Hangul syllable", 0xd55c],
      ["Hangul Jamo", 0x1100],
      ["Hangul Jamo Extended-A", 0xa960],
      ["Hangul Jamo Extended-B", 0xd7b0],
      ["Hangul compatibility Jamo", 0x3131],
      ["halfwidth Hangul Jamo", 0xffa0],
      ["Latin", 0x41],
      ["space", 0x20]
    ] as [string, number][]) {
      expect(isSpacelessScript(cp), label).toBe(false);
    }
  });

  it("joins a whole hard-wrapped Chinese paragraph", async () => {
    const wrapped =
      "这个约束是贯穿全局的设计压力。它排除了遇到难处就找个库这个选项，\n所以代码要么自己把活干完，要么这个功能就不存在。";
    expect(await firstParagraph(wrapped)).toBe(
      "这个约束是贯穿全局的设计压力。它排除了遇到难处就找个库这个选项，所以代码要么自己把活干完，要么这个功能就不存在。"
    );
  });
});
