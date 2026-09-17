/**
 * The rasteriser must use the same font discovery the PDF embedder does.
 *
 * This is the direct regression test for the original defect, and it is separate
 * from `font-fallback.test.ts` for a reason: that file injects synthetic fonts and
 * switches the host's off, which proves the *chain* works but cannot prove the
 * chain is ever populated. The bug was precisely that it was not — the rasteriser
 * had a hardcoded list of Latin filenames and never called discovery at all, so a
 * machine with `Songti SC` installed still drew Chinese as blank space.
 *
 * `.node.test.ts` because it reads the host's font directories, which is also why
 * every assertion below is conditional on the host actually having a CJK face:
 * a CI container with none must not fail a test about this repository's code. The
 * oracle is `findSystemFontForCodePoints` — the function the PDF pipeline uses — so
 * a passing run means the two backends agree about what is installed, which is the
 * property that was broken.
 */

import { BasicRasterCanvas } from "@draw/raster/canvas";
import type { RasterFont } from "@draw/raster/glyph-outline";
import { fontHasGlyph } from "@draw/raster/glyph-outline";
import { resetDiscoveredFonts, resolveFontChain } from "@draw/raster/system-raster-font";
import { buildCoverageFont } from "@test/ttf-fixture";
import { detectCjkLanguage } from "@utils/cjk";
import type * as FontDiscovery from "@utils/font-discovery";
import {
  _setCandidatesForTest,
  findSystemFontForCodePoints,
  resetFontDiscoveryCache
} from "@utils/font-discovery";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

// Wraps the real function rather than replacing it, so the suite still exercises
// genuine discovery and can additionally count how often it is asked.
vi.mock("@utils/font-discovery", async importOriginal => {
  const actual = await importOriginal<typeof FontDiscovery>();
  return { ...actual, findSystemFontForCodePoints: vi.fn(actual.findSystemFontForCodePoints) };
});

const HAN_ZHONG = 0x4e2d; // 中

/** Whether this host has any face that can draw 中. */
function hostHasCjkFace(): boolean {
  return findSystemFontForCodePoints(new Set([HAN_ZHONG])) !== null;
}

describe("raster font discovery", () => {
  beforeEach(() => {
    resetDiscoveredFonts();
  });
  afterEach(() => {
    resetDiscoveredFonts();
  });

  it("puts a CJK face in the chain when the host has one", () => {
    if (!hostHasCjkFace()) {
      return; // No CJK font installed; nothing to assert about this host.
    }
    const chain = resolveFontChain("中文", [], true);
    // The bug: the chain was [Arial] and this was false.
    expect(chain.some(font => fontHasGlyph(font, HAN_ZHONG))).toBe(true);
  });

  it("orders a Latin face ahead of the discovered one", () => {
    if (!hostHasCjkFace()) {
      return;
    }
    const chain = resolveFontChain("Mixed 混合", [], true);
    const latinIndex = chain.findIndex(font => fontHasGlyph(font, 0x41));
    const cjkIndex = chain.findIndex(font => fontHasGlyph(font, HAN_ZHONG));
    expect(latinIndex).toBeGreaterThanOrEqual(0);
    expect(cjkIndex).toBeGreaterThanOrEqual(0);
    // ASCII must come from the Latin face, so mixed text is not set in a CJK
    // font's serif hand while the SVG of the same drawing uses Arial.
    if (latinIndex !== cjkIndex) {
      expect(latinIndex).toBeLessThan(cjkIndex);
    }
  });

  it("asks for nothing when the text is plain ASCII", () => {
    const chain = resolveFontChain("ABC 123", [], true);
    // A Latin face answers this, so discovery must not be consulted — and must not
    // append a second face for a string that needed none.
    expect(chain.length).toBeLessThanOrEqual(1);
  });

  it("builds one face for many labels of the same script", () => {
    if (!hostHasCjkFace()) {
      return;
    }
    // Deliberately mixes labels that `detectCjkLanguage` answers differently for:
    // 开始处理 is shaped identically in Simplified and Traditional so it yields
    // `undefined`, while 数据校验 carries simplified-only forms and yields
    // `zh-Hans`. Both resolve to the same installed font, and both must reuse the
    // same parsed face.
    const labels = [
      "开始处理",
      "数据校验",
      "写入数据库",
      "记录错误日志",
      "发送通知",
      "生成报表",
      "用户认证",
      "权限检查",
      "缓存更新",
      "消息队列",
      "定时任务",
      "日志归档",
      "备份恢复",
      "监控告警",
      "配置中心",
      "服务注册",
      "负载均衡",
      "熔断降级"
    ];
    // The premise of the test: these really do take different cache keys.
    const keys = new Set(labels.map(l => String(detectCjkLanguage(l))));
    expect(keys.size).toBeGreaterThan(1);

    resetDiscoveredFonts();
    const faces = new Set<RasterFont>();
    for (const label of labels) {
      for (const face of resolveFontChain(label, [], true)) {
        faces.add(face);
      }
    }

    // Exactly a Latin face plus a CJK one.
    // was its own key and this was 19 — 19 parses of the same font, 19 copies of a
    // 43,000-entry cmap, and a glyph cache that never hit.
    expect(faces.size).toBeLessThanOrEqual(2);
    expect(faces.size).toBeLessThan(labels.length);
  });

  it("searches once per script, not once per label", () => {
    if (!hostHasCjkFace()) {
      return;
    }
    // Face identity keeps the *parse* down to one either way, so this is the other
    // half and it is not free: keyed per label the searches alone cost 293 ms for
    // these 18 labels against 17 ms keyed per script, because each one re-runs
    // candidate enumeration and the coverage check.
    const labels = [
      "开始处理",
      "数据校验",
      "写入数据库",
      "记录错误日志",
      "发送通知",
      "生成报表",
      "用户认证",
      "权限检查",
      "缓存更新",
      "消息队列"
    ];
    const scripts = new Set(labels.map(l => String(detectCjkLanguage(l))));

    resetDiscoveredFonts();
    resolveFontChain("预热", [], true); // pay the one-time path-index cost outside the count
    const spy = vi.mocked(findSystemFontForCodePoints);
    spy.mockClear();

    for (const label of labels) {
      resolveFontChain(label, [], true);
    }

    // At most one search per distinct script key already seen, plus a little slack
    // for a label carrying a character the script's face lacks.
    expect(spy.mock.calls.length).toBeLessThanOrEqual(scripts.size + 1);
    expect(spy.mock.calls.length).toBeLessThan(labels.length);
  });

  it("keeps asking until a partially-covering face chain covers the text", () => {
    // Discovery may return a face covering only *part* of what was asked for — that
    // is deliberate, since most of the text beats none of it. Taking the first
    // answer and stopping dropped the rest: `αא` loaded the Greek face and lost the
    // Hebrew even with a Hebrew font installed.
    //
    // Private-use code points, so no real installed font can quietly satisfy this
    // and the synthetic faces are the only way to draw the string.
    //
    // Three of them, deliberately: the per-script search supplies one face and a
    // single follow-up search supplies a second, so a two-character string is
    // covered even without a loop. Only a third character proves the search repeats.
    const wanted = [0xe000, 0xe001, 0xe002];
    const text = wanted.map(cp => String.fromCodePoint(cp)).join("");

    resetFontDiscoveryCache();
    _setCandidatesForTest(wanted.map((cp, i) => buildCoverageFont([cp], `Only PUA-${i}`)));
    resetDiscoveredFonts();

    const chain = resolveFontChain(text, [], true);
    for (const cp of wanted) {
      expect(chain.some(f => fontHasGlyph(f, cp))).toBe(true);
    }

    const canvas = new BasicRasterCanvas(200, 60);
    canvas.drawText(10, 40, text, 24, "#000000", "start");
    expect([...canvas.uncoveredCodePoints]).toEqual([]);

    resetFontDiscoveryCache();
    _setCandidatesForTest([]);
  });

  it("does not grow the chain when the text can never be fully covered", () => {
    // A face that helps with nothing must not be appended again on every label. The
    // search cannot succeed, so without recording the failure there is nothing to
    // stop it repeating — and each repeat left another face in the chain, which is
    // then walked per character for the rest of the process.
    const drawable = 0xe000;
    const impossible = 0xe500;
    const text = String.fromCodePoint(drawable) + String.fromCodePoint(impossible);

    resetFontDiscoveryCache();
    _setCandidatesForTest([buildCoverageFont([drawable], "Only PUA-1")]);
    resetDiscoveredFonts();

    const lengths = new Set<number>();
    for (let i = 0; i < 50; i++) {
      lengths.add(resolveFontChain(text, [], true).length);
    }
    expect(lengths.size).toBe(1);

    const canvas = new BasicRasterCanvas(200, 60);
    canvas.drawText(10, 40, text, 24, "#000000", "start");
    // The one it cannot draw is reported; the one it can is not.
    expect([...canvas.uncoveredCodePoints]).toEqual([impossible]);

    resetFontDiscoveryCache();
    _setCandidatesForTest([]);
  });

  it("caches a discovery rather than repeating the directory walk", () => {
    if (!hostHasCjkFace()) {
      return;
    }
    const first = resolveFontChain("中文", [], true);
    const second = resolveFontChain("中文", [], true);
    // Same face objects, so the second call resolved from the cache. A fresh parse
    // would also break the per-outline glyph cache, which is keyed by identity.
    expect(second).toEqual(first);
    expect(second[second.length - 1]).toBe(first[first.length - 1]);
  });
});
