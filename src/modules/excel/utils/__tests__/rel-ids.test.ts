import { extractAll } from "@archive/unzip/extract";
import { requireEntryText } from "@excel/__tests__/helpers/zip-text";
import { Cell, Workbook } from "@excel/index";
import { createFirstFreeRelIds, createMaxRelIds } from "@excel/utils/rel-ids";
import { describe, expect, it } from "vitest";

/** A relationship list that counts how many entries were read, to pin the allocators' cost. */
function counted(initial: { Id?: string }[] = []) {
  const reads = { n: 0 };
  const rels = new Proxy(initial, {
    get(target, key, receiver) {
      if (typeof key === "string" && /^\d+$/.test(key)) {
        reads.n++;
      }
      return Reflect.get(target, key, receiver);
    }
  });
  return { rels, reads };
}

describe("createMaxRelIds", () => {
  it("allocates above the highest id in use, gaps and foreign ids included", () => {
    const next = createMaxRelIds();
    const rels: { Id?: string }[] = [{ Id: "rId7" }, { Id: "rIdFoo" }, {}, { Id: "rId2" }];
    expect(next(rels)).toBe("rId8");
    rels.push({ Id: "rId8" });
    expect(next(rels)).toBe("rId9");
    rels.push({ Id: "rId20" });
    expect(next(rels)).toBe("rId21");
  });

  it("tracks each list separately", () => {
    const next = createMaxRelIds();
    const a = [{ Id: "rId3" }];
    const b: { Id?: string }[] = [];
    expect(next(a)).toBe("rId4");
    expect(next(b)).toBe("rId1");
  });

  it("rescans a list that shrank", () => {
    const next = createMaxRelIds();
    const rels = [{ Id: "rId1" }, { Id: "rId2" }];
    expect(next(rels)).toBe("rId3");
    rels.length = 0;
    rels.push({ Id: "rId1" });
    expect(next(rels)).toBe("rId2");
  });

  it("reads each entry once however many ids are allocated", () => {
    const next = createMaxRelIds();
    const { rels, reads } = counted();
    for (let i = 0; i < 2000; i++) {
      rels.push({ Id: next(rels) });
    }
    expect(rels[1999].Id).toBe("rId2000");
    // 2000 for the final check above plus one read per pushed entry — not 2000²/2.
    expect(reads.n).toBeLessThan(5000);
  });
});

describe("createFirstFreeRelIds", () => {
  it("allocates the lowest free id from 1 and records it as taken", () => {
    const taken = new Set(["rId1", "rId3"]);
    const next = createFirstFreeRelIds(taken);
    expect(next()).toBe("rId2");
    expect(next()).toBe("rId4");
    taken.add("rId5");
    expect(next()).toBe("rId6");
    expect([...taken].sort()).toEqual(["rId1", "rId2", "rId3", "rId4", "rId5", "rId6"]);
  });

  it("probes a constant number of candidates per id", () => {
    let probes = 0;
    const taken = new Set<string>();
    const has = taken.has.bind(taken);
    taken.has = (id: string) => {
      probes++;
      return has(id);
    };
    const next = createFirstFreeRelIds(taken);
    for (let i = 0; i < 2000; i++) {
      next();
    }
    expect(probes).toBeLessThan(5000);
  });
});

describe("hyperlink relationships", () => {
  for (const format of ["xlsx", "xlsb"] as const) {
    it(`numbers ${format} hyperlink relationships rId1..rIdN`, async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      const n = 300;
      for (let i = 1; i <= n; i++) {
        Cell.setValue(ws, i, 1, { text: `l${i}`, hyperlink: `https://example.com/${i}` });
      }
      const entries = await extractAll(await Workbook.toBuffer(wb, { format }));
      const rels = requireEntryText(
        entries,
        `xl/worksheets/_rels/sheet1.${format === "xlsx" ? "xml" : "bin"}.rels`
      );
      const ids = [...rels.matchAll(/Id="(rId\d+)"[^>]*Target="https:\/\/example\.com\/(\d+)"/g)];
      expect(ids).toHaveLength(n);
      expect(new Set(ids.map(m => m[1])).size).toBe(n);
      expect(ids.map(m => m[1]).sort()).toEqual(
        Array.from({ length: n }, (_, i) => `rId${i + 1}`).sort()
      );
    });
  }
});
