import {
  cellFont,
  cellGetStyle,
  cellOwnStyle,
  mergeCellStyle,
  setFacetShared
} from "@excel/core/cell";
import { getCell, getColumn } from "@excel/core/worksheet";
import { Cell, Column, Row, Table, Workbook } from "@excel/index";
import { describe, it, expect } from "vitest";

/**
 * Cells covered by a styled row or column point at **one shared, frozen snapshot**
 * of its facets rather than each holding a deep copy — a 200k-cell sheet carrying one
 * logical style went from +155 MB to no measurable increase.
 *
 * The contract that makes that safe has three halves, and each is pinned here:
 *
 *  1. `Cell.get*` gives the cell its own copy, so mutating what it returns cannot
 *     reach a sibling ("style isolation").
 *  2. The shared snapshot is frozen, so any path that skips (1) throws instead of
 *     silently rewriting every cell that shares it.
 *  3. The snapshot is a *copy* of the owner's facets, so mutating the row/column
 *     itself does not reach cells it has already styled.
 */
describe("style sharing", () => {
  function sheet() {
    const wb = Workbook.create();
    return Workbook.addWorksheet(wb, "S");
  }

  describe("cells share one snapshot", () => {
    it("points every cell in a styled column at the same facet object", () => {
      const ws = sheet();
      Cell.setValue(ws, "A1", 1);
      Cell.setValue(ws, "A2", 2);
      Column.setStyle(ws, 1, { font: { bold: true, size: 12 } });

      // Read without materialising — the internal view the writers use.
      const a1 = cellGetStyle(getCell(ws, 1, 1)).font;
      const a2 = cellGetStyle(getCell(ws, 2, 1)).font;
      expect(a1).toBe(a2);
      expect(Object.isFrozen(a1)).toBe(true);
    });

    it("shares with cells created after the style was set", () => {
      const ws = sheet();
      Column.setStyle(ws, 1, { font: { bold: true } });
      Cell.setValue(ws, "A1", 1);
      Cell.setValue(ws, "A2", 2);
      expect(cellGetStyle(getCell(ws, 1, 1)).font).toBe(cellGetStyle(getCell(ws, 2, 1)).font);
    });

    it("shares a nested facet too, not just the top object", () => {
      const ws = sheet();
      Cell.setValue(ws, "A1", 1);
      Cell.setValue(ws, "A2", 2);
      Column.setStyle(ws, 1, { border: { top: { style: "thin", color: { argb: "FF010203" } } } });
      const a1 = cellGetStyle(getCell(ws, 1, 1)).border!;
      const a2 = cellGetStyle(getCell(ws, 2, 1)).border!;
      expect(a1.top).toBe(a2.top);
      expect(Object.isFrozen(a1.top!.color)).toBe(true);
    });

    it("does not hand cells the owner's own facet object", () => {
      const ws = sheet();
      Cell.setValue(ws, "A1", 1);
      const font = { bold: true };
      Column.setStyle(ws, 1, { font });
      // The column keeps the caller's object; the cell gets the snapshot.
      expect(getColumn(ws, 1).style.font).toBe(font);
      expect(cellGetStyle(getCell(ws, 1, 1)).font).not.toBe(font);
      expect(Object.isFrozen(font)).toBe(false);
    });
  });

  describe("Cell.get* separates a cell again", () => {
    it("isolates a column-propagated facet", () => {
      const ws = sheet();
      Cell.setValue(ws, "A1", 1);
      Cell.setValue(ws, "A2", 2);
      Column.setStyle(ws, 1, { font: { bold: true } });

      Cell.getStyle(ws, "A1").font!.bold = false;

      expect(Cell.getStyle(ws, "A1").font!.bold).toBe(false);
      expect(Cell.getStyle(ws, "A2").font!.bold).toBe(true);
    });

    it("isolates through each facet reader", () => {
      const ws = sheet();
      Cell.setValue(ws, "A1", 1);
      Cell.setValue(ws, "A2", 2);
      Column.setStyle(ws, 1, {
        font: { bold: true },
        alignment: { horizontal: "left" },
        border: { top: { style: "thin" } },
        fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } },
        protection: { locked: true }
      });

      Cell.getFont(ws, "A1")!.bold = false;
      Cell.getAlignment(ws, "A1")!.horizontal = "right";
      Cell.getBorder(ws, "A1")!.top = { style: "thick" };
      Cell.getProtection(ws, "A1")!.locked = false;

      expect(Cell.getFont(ws, "A2")!.bold).toBe(true);
      expect(Cell.getAlignment(ws, "A2")!.horizontal).toBe("left");
      expect(Cell.getBorder(ws, "A2")!.top).toEqual({ style: "thin" });
      expect(Cell.getProtection(ws, "A2")!.locked).toBe(true);
    });

    it("is idempotent — a second read does not re-copy", () => {
      const ws = sheet();
      Cell.setValue(ws, "A1", 1);
      Column.setStyle(ws, 1, { font: { bold: true } });
      const first = Cell.getStyle(ws, "A1").font;
      const second = Cell.getStyle(ws, "A1").font;
      expect(second).toBe(first);
      expect(Object.isFrozen(first)).toBe(false);
    });

    it("separates the aliasing that one style object used to create", () => {
      const ws = sheet();
      Cell.setValue(ws, "A1", 1);
      Cell.setValue(ws, "A2", 2);
      // Setters store the caller's facet by reference. Before cells could be
      // separated again, this aliased A1 and A2 permanently — mutating A1's font
      // changed A2 *and* the caller's own object.
      const shared = { font: { bold: true } };
      Cell.setStyle(ws, "A1", shared);
      Cell.setStyle(ws, "A2", shared);

      Cell.getStyle(ws, "A1").font!.bold = false;

      expect(Cell.getStyle(ws, "A1").font!.bold).toBe(false);
      expect(Cell.getStyle(ws, "A2").font!.bold).toBe(true);
      expect(shared.font.bold).toBe(true);
    });
  });

  describe("the snapshot is frozen", () => {
    it("throws rather than rewriting every cell that shares it", () => {
      const ws = sheet();
      Cell.setValue(ws, "A1", 1);
      Cell.setValue(ws, "A2", 2);
      Column.setStyle(ws, 1, { font: { bold: true } });
      const shared = cellFont(getCell(ws, 1, 1))!;
      expect(() => {
        shared.bold = false;
      }).toThrow(TypeError);
      expect(Cell.getStyle(ws, "A2").font!.bold).toBe(true);
    });
  });

  describe("mutating the owner does not reach cells it already styled", () => {
    it("holds for a column", () => {
      const ws = sheet();
      Cell.setValue(ws, "A1", 1);
      Column.setStyle(ws, 1, { font: { bold: true } });

      Column.getStyle(ws, 1).font!.bold = false;

      expect(Cell.getStyle(ws, "A1").font!.bold).toBe(true);
    });

    it("holds for a row", () => {
      const ws = sheet();
      Cell.setValue(ws, "A1", 1);
      Row.setStyle(ws, 1, { font: { bold: true } });

      Row.getStyle(ws, 1).font!.bold = false;

      expect(Cell.getStyle(ws, "A1").font!.bold).toBe(true);
    });

    it("but a cell created afterwards sees the owner's current style", () => {
      const ws = sheet();
      Cell.setValue(ws, "A1", 1);
      Column.setStyle(ws, 1, { font: { bold: true } });

      // Handing the column's style to a caller drops the snapshot, because the
      // caller may mutate it in place — which is exactly what happens next.
      Column.getStyle(ws, 1).font!.bold = false;
      Cell.setValue(ws, "A2", 2);

      expect(Cell.getStyle(ws, "A1").font!.bold).toBe(true);
      expect(Cell.getStyle(ws, "A2").font!.bold).toBe(false);
    });

    it("re-snapshots after the owner's style is replaced wholesale", () => {
      const ws = sheet();
      Cell.setValue(ws, "A1", 1);
      Column.setStyle(ws, 1, { font: { bold: true } });
      const before = cellGetStyle(getCell(ws, 1, 1)).font;

      // A wholesale replacement invalidates by construction: the snapshot is keyed
      // by the style object, so a new object cannot find a stale entry.
      Column.setStyle(ws, 1, { font: { italic: true } });

      const after = cellGetStyle(getCell(ws, 1, 1)).font;
      expect(after).not.toBe(before);
      expect(after).toEqual({ italic: true });
    });
  });

  describe("row wins over column, as before", () => {
    it("takes the row's facet when both carry one", () => {
      const ws = sheet();
      Column.setStyle(ws, 1, { font: { bold: true, size: 10 } });
      Row.setStyle(ws, 1, { font: { italic: true, size: 20 } });
      Cell.setValue(ws, "A1", 1);
      expect(Cell.getStyle(ws, "A1").font).toEqual({ italic: true, size: 20 });
    });
  });

  describe("primitives are not snapshotted", () => {
    it("passes numFmt and styleName straight through", () => {
      const ws = sheet();
      Cell.setValue(ws, "A1", 1);
      Column.setNumFmt(ws, 1, "0.00%");
      expect(Cell.getStyle(ws, "A1").numFmt).toBe("0.00%");
      // A primitive cannot be aliased, so it needs no copy and no freeze.
      const style = cellGetStyle(getCell(ws, 1, 1));
      expect(Object.isFrozen(style)).toBe(false);
    });
  });

  describe("the internal seam", () => {
    it("cellOwnStyle is a no-op for a cell that already owns its facets", () => {
      const ws = sheet();
      Cell.setValue(ws, "A1", 1);
      const cell = getCell(ws, 1, 1);
      const first = cellOwnStyle(cell);
      expect(cellOwnStyle(cell)).toBe(first);
    });

    it("setFacetShared marks the cell, mergeCellStyle reads the snapshot", () => {
      const ws = sheet();
      Cell.setValue(ws, "A1", 1);
      const cell = getCell(ws, 1, 1);
      const owner = { font: { bold: true } };
      // `mergeCellStyle` resolves a new cell's facets out of its row/column.
      const merged = mergeCellStyle(owner, {}, {});
      expect(Object.isFrozen(merged.font)).toBe(true);

      setFacetShared(cell, "font", merged.font);
      expect(cellGetStyle(cell).font).toBe(merged.font);
      // …and the public reader hands back a private, mutable copy.
      const own = Cell.getStyle(ws, "A1").font!;
      expect(own).toEqual({ bold: true });
      expect(Object.isFrozen(own)).toBe(false);
    });
  });

  describe("a reference the caller is still holding stays live", () => {
    it("survives an update that touches only a primitive facet", () => {
      const ws = sheet();
      Cell.setValue(ws, "A1", 1);
      Cell.setStyle(ws, "A1", { font: { bold: true } });

      // The documented way to tweak a facet — the isolation tests do exactly this.
      const font = Cell.getStyle(ws, "A1").font!;

      // Marking the cell as sharing on *any* set, even a numFmt-only one, made the
      // next read re-copy facets the cell already owned, silently detaching this
      // reference and losing the write below.
      Cell.setStyle(ws, "A1", { numFmt: "0.00" });
      Cell.getStyle(ws, "A1");

      font.bold = false;
      expect(Cell.getStyle(ws, "A1").font!.bold).toBe(false);
    });

    it("survives Cell.setNumFmt", () => {
      const ws = sheet();
      Cell.setValue(ws, "A1", 1);
      Cell.setStyle(ws, "A1", { font: { bold: true } });
      const font = Cell.getStyle(ws, "A1").font!;
      Cell.setNumFmt(ws, "A1", "0.00");
      Cell.getStyle(ws, "A1");
      font.bold = false;
      expect(Cell.getStyle(ws, "A1").font!.bold).toBe(false);
    });

    it("is correctly replaced when the facet itself is set again", () => {
      const ws = sheet();
      Cell.setValue(ws, "A1", 1);
      Cell.setStyle(ws, "A1", { font: { bold: true } });
      const stale = Cell.getStyle(ws, "A1").font!;

      Cell.setStyle(ws, "A1", { font: { italic: true } });

      // The old facet was replaced, so writing to it must *not* resurrect it.
      stale.bold = false;
      expect(Cell.getStyle(ws, "A1").font).toEqual({ italic: true });
    });
  });

  describe("table columns", () => {
    it("does not alias the cells a table column styled", () => {
      const ws = sheet();
      // Materialise first, so nothing but the table's own marking can save this.
      Cell.setValue(ws, "A1", 0);
      Cell.setValue(ws, "A2", 0);
      Cell.getStyle(ws, "A1");
      Cell.getStyle(ws, "A2");

      Table.add(ws, {
        name: "t",
        ref: "A1",
        headerRow: true,
        style: { showRowStripes: false },
        columns: [{ name: "A", style: { font: { bold: true } } }],
        rows: [[1], [2]]
      });

      Cell.getStyle(ws, "A1").font!.bold = false;

      expect(Cell.getStyle(ws, "A1").font!.bold).toBe(false);
      expect(Cell.getStyle(ws, "A2").font!.bold).toBe(true);
    });

    it("shares one snapshot across the column's cells", () => {
      const ws = sheet();
      Table.add(ws, {
        name: "t",
        ref: "A1",
        headerRow: true,
        style: { showRowStripes: false },
        columns: [{ name: "A", style: { font: { bold: true } } }],
        rows: [[1], [2]]
      });
      const header = cellGetStyle(getCell(ws, 1, 1)).font;
      const body = cellGetStyle(getCell(ws, 2, 1)).font;
      expect(header).toBe(body);
      expect(Object.isFrozen(header)).toBe(true);
    });

    it("re-snapshots after Table.columnStyle is mutated in place", () => {
      const ws = sheet();
      const table = Table.add(ws, {
        name: "t",
        ref: "A1",
        headerRow: true,
        style: { showRowStripes: false },
        columns: [{ name: "A", style: { font: { bold: true } } }],
        rows: [[1]]
      });
      expect(Cell.getStyle(ws, "A2").font!.bold).toBe(true);

      // Handing the style out drops the snapshot, so a later store() applies what
      // the column says now.
      Table.columnStyle(Table.column(table, 0))!.font!.bold = false;
      Table.addRow(table, [2]);
      Table.commit(table);

      expect(Cell.getStyle(ws, "A3").font!.bold).toBe(false);
    });
  });

  describe("Cell.view hands out the shared facet", () => {
    it("is frozen, and rejects a nested write at compile time", () => {
      const ws = sheet();
      Cell.setValue(ws, "A1", 1);
      Cell.setValue(ws, "A2", 2);
      Column.setStyle(ws, 1, { font: { bold: true } });

      const view = Cell.view(getCell(ws, 1, 1));
      // Deliberately not materialised: `Cell.view` is a read-only projection, and
      // copying here would make iterating a sheet allocate per cell.
      expect(view.font).toBe(cellGetStyle(getCell(ws, 2, 1)).font);
      expect(Object.isFrozen(view.font)).toBe(true);

      // `CellView.font` is `DeepReadonly`, so this is a type error rather than only a
      // runtime throw — a shallow `readonly` let the nested write compile and the
      // type then promised something the frozen object did not honour.
      // @ts-expect-error the projection is deeply read-only
      expect(() => (view.font!.bold = false)).toThrow(TypeError);

      // Reading it out mutably is what `Cell.get*` is for.
      expect(Cell.getFont(ws, "A1")).toEqual({ bold: true });
      expect(Object.isFrozen(Cell.getFont(ws, "A1"))).toBe(false);
    });
  });

  describe("the read path shares too", () => {
    it("gives cells that loaded with the same styleId one snapshot", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      for (let r = 1; r <= 4; r++) {
        Cell.setValue(ws, r, 1, r);
      }
      Column.setStyle(ws, 1, { font: { bold: true, size: 13 } });
      const bytes = new Uint8Array(await Workbook.toBuffer(wb));

      const loaded = Workbook.create();
      await Workbook.read(loaded, bytes);
      const ws2 = Workbook.getWorksheet(loaded, "S")!;

      // The reader caches one style model per styleId, so every cell carrying that
      // id shares a single snapshot instead of deep-copying the facets per cell.
      const a1 = cellGetStyle(getCell(ws2, 1, 1)).font;
      const a2 = cellGetStyle(getCell(ws2, 2, 1)).font;
      expect(a1).toEqual({ bold: true, size: 13 });
      expect(a1).toBe(a2);
      expect(Object.isFrozen(a1)).toBe(true);
    });

    it("still isolates a loaded cell through the public reader", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      Cell.setValue(ws, "A1", 1);
      Cell.setValue(ws, "A2", 2);
      Column.setStyle(ws, 1, { font: { bold: true } });
      const bytes = new Uint8Array(await Workbook.toBuffer(wb));

      const loaded = Workbook.create();
      await Workbook.read(loaded, bytes);
      const ws2 = Workbook.getWorksheet(loaded, "S")!;

      Cell.getStyle(ws2, "A1").font!.bold = false;

      expect(Cell.getStyle(ws2, "A1").font!.bold).toBe(false);
      expect(Cell.getStyle(ws2, "A2").font!.bold).toBe(true);
    });

    it("preserves the xf-level fields the reader carries alongside the facets", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      Cell.setValue(ws, "A1", 1);
      Cell.setStyle(ws, "A1", { numFmt: "0.00%", font: { bold: true } });
      const bytes = new Uint8Array(await Workbook.toBuffer(wb));

      const loaded = Workbook.create();
      await Workbook.read(loaded, bytes);
      const ws2 = Workbook.getWorksheet(loaded, "S")!;

      // Replacing only the object facets with snapshots must not drop the
      // primitives sitting beside them on the model.
      expect(Cell.getStyle(ws2, "A1").numFmt).toBe("0.00%");
      expect(Cell.getStyle(ws2, "A1").font!.bold).toBe(true);
    });
  });

  describe("the file is unaffected", () => {
    it("writes the same styles whether a cell was separated or not", async () => {
      const build = (separate: boolean) => {
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "S");
        for (let r = 1; r <= 4; r++) {
          Cell.setValue(ws, r, 1, r);
        }
        Column.setStyle(ws, 1, { font: { bold: true }, numFmt: "0.00" });
        if (separate) {
          // Materialise every cell; the content is identical, so the written
          // stylesheet must be too.
          for (let r = 1; r <= 4; r++) {
            Cell.getStyle(ws, r, 1);
          }
        }
        return wb;
      };
      const shared = new Uint8Array(await Workbook.toBuffer(build(false)));
      const separated = new Uint8Array(await Workbook.toBuffer(build(true)));
      expect(separated.length).toBe(shared.length);
    });
  });
});
