import { extractAll } from "@archive/unzip/extract";
import { Cell, Chart, Sparkline, Table, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

/**
 * What the Open XML SDK's validator — Office's own schema — refused in workbooks this library wrote, each checked
 * against the file Excel writes when it saves the same workbook. What Excel dropped or reordered on saving was ours
 * to fix; what Excel writes itself (a fallback shape's `id="0"`, `<cx:binCount val>`) stays.
 */

async function parts(wb: Workbook.Handle): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  for (const [path, entry] of await extractAll(await Workbook.toBuffer(wb))) {
    if (path.endsWith(".xml")) {
      out.set(path, new TextDecoder().decode(entry.data));
    }
  }
  return out;
}

function sheetWithData(): { wb: Workbook.Handle; ws: ReturnType<typeof Workbook.addWorksheet> } {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "Sales");
  ["Month", "Revenue"].forEach((v, i) => Cell.setValue(ws, { row: 1, col: i + 1 }, v));
  for (let r = 2; r <= 5; r++) {
    Cell.setValue(ws, { row: r, col: 1 }, `M${r}`);
    Cell.setValue(ws, { row: r, col: 2 }, r * 10);
  }
  return { wb, ws };
}

describe("tables", () => {
  it("writes a column's data format as dataDxfId, and reads it back", async () => {
    const { wb, ws } = sheetWithData();
    Table.add(ws, {
      name: "T",
      ref: "A1",
      headerRow: true,
      columns: [{ name: "Month" }, { name: "Revenue", style: { numFmt: "0.00" } }],
      rows: [
        ["a", 1],
        ["b", 2]
      ]
    });
    const xml = (await parts(wb)).get("xl/tables/table1.xml")!;
    expect(xml).toMatch(/<tableColumn id="2" name="Revenue" dataDxfId="\d+"\/>/);
    expect(xml).not.toContain(" dxfId=");

    const back = await Workbook.read(await Workbook.toBuffer(wb));
    const table = Table.get(Workbook.getWorksheet(back, "Sales")!, "T")!;
    expect(Table.getModel(table).columns[1]!.style?.numFmt).toBe("0.00");
  });
});

describe("charts", () => {
  it("refuses a first slice angle where the chart type has none", () => {
    for (const type of ["pie3D", "ofPie"] as const) {
      const { ws } = sheetWithData();
      expect(() =>
        Chart.place(ws, "D2:J15", {
          type,
          firstSliceAng: 45,
          series: [{ name: "Revenue", categories: "Sales!$A$2:$A$5", values: "Sales!$B$2:$B$5" }]
        })
      ).toThrow(/firstSliceAng/);
    }
  });

  it("writes a first slice angle into a pie and not into a 3-D pie", async () => {
    const { wb, ws } = sheetWithData();
    const series = [{ name: "Revenue", categories: "Sales!$A$2:$A$5", values: "Sales!$B$2:$B$5" }];
    Chart.place(ws, "D2:J15", { type: "pie", firstSliceAng: 45, series });
    Chart.place(ws, "D20:J35", { type: "pie3D", series });
    const charts = [...(await parts(wb))].filter(([p]) => /xl\/charts\/chart\d+\.xml$/.test(p));
    const pie = charts.find(([, x]) => x.includes("<c:pieChart>"))![1];
    const pie3D = charts.find(([, x]) => x.includes("<c:pie3DChart>"))![1];
    expect(pie).toContain('<c:firstSliceAng val="45"/>');
    expect(pie3D).not.toContain("c:firstSliceAng");
  });

  it("writes where an axis crosses as one of crosses or crossesAt", async () => {
    const { wb, ws } = sheetWithData();
    Chart.place(ws, "D2:J15", {
      type: "line",
      series: [{ name: "Revenue", categories: "Sales!$A$2:$A$5", values: "Sales!$B$2:$B$5" }],
      valueAxis: { crossesAt: 20 }
    });
    const chart = [...(await parts(wb))].find(([p]) => /xl\/charts\/chart\d+\.xml$/.test(p))![1];
    const axes = chart.match(/<c:(?:catAx|valAx)>[\s\S]*?<\/c:(?:catAx|valAx)>/g)!;
    for (const axis of axes) {
      expect(Number(axis.includes("<c:crosses ")) + Number(axis.includes("<c:crossesAt "))).toBe(1);
    }
    expect(chart).toContain('<c:crossesAt val="20"/>');
  });

  it("writes a chartEx series' data labels in CT_DataLabels' order: numFmt before visibility", async () => {
    const { wb, ws } = sheetWithData();
    Chart.place(ws, "D2:J15", {
      type: "waterfall",
      categories: "Sales!$A$2:$A$5",
      series: [
        {
          name: "Revenue",
          values: "Sales!$B$2:$B$5",
          dataLabels: { showValue: true, numFmt: "$#,##0" }
        }
      ]
    });
    const xml = [...(await parts(wb))].find(([p]) => /chartEx\d+\.xml$/.test(p))![1];
    const labels = /<cx:dataLabels[^>]*>([\s\S]*?)<\/cx:dataLabels>/.exec(xml)![1]!;
    expect(labels.indexOf("<cx:numFmt")).toBeGreaterThanOrEqual(0);
    expect(labels.indexOf("<cx:numFmt")).toBeLessThan(labels.indexOf("<cx:visibility"));
  });
});

describe("sparklines", () => {
  it("writes every colour as ARGB, from an option or set on the group", async () => {
    const { wb, ws } = sheetWithData();
    const group = Sparkline.add(ws, {
      type: "line",
      sparklines: [{ dataRef: "Sales!B2:B5", cellRef: "C2" }],
      lineColor: "#C00000"
    });
    // Set on the group itself, which `Sparkline.add` hands back to edit.
    group.colorHigh = { rgb: "70AD47" };
    const sheet = (await parts(wb)).get("xl/worksheets/sheet1.xml")!;
    expect(sheet).toContain('<x14:colorSeries rgb="FFC00000"/>');
    expect(sheet).toContain('<x14:colorHigh rgb="FF70AD47"/>');
    expect(sheet).not.toMatch(/rgb="[0-9A-F]{6}"/);
  });
});

describe("styles", () => {
  it("writes a font's children in Excel's order: underline before size, charset before scheme", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", "x");
    Cell.setStyle(ws, "A1", {
      font: {
        name: "宋体",
        size: 11,
        underline: true,
        charset: 134,
        scheme: "minor",
        vertAlign: "superscript"
      }
    });
    const styles = (await parts(wb)).get("xl/styles.xml")!;
    const font = /<font>(?:(?!<\/font>)[\s\S])*<u\/>(?:(?!<\/font>)[\s\S])*<\/font>/.exec(
      styles
    )![0];
    const order = [...font.matchAll(/<([a-zA-Z]+)[ />]/g)].map(m => m[1]).filter(n => n !== "font");
    expect(order).toEqual(["u", "vertAlign", "sz", "name", "charset", "scheme"]);
  });
});
