/**
 * A region map's topology is the caller's data, so a topology that cannot be read is an
 * argument error naming where it is wrong — `ChartOptionsError` with the exact path from the
 * public parameter (`options.regionMap.topology.arcs[3]`) and the call — not a plain `Error`,
 * a `TypeError` from deep inside the decoder, or a silent fall-back to the centroid preview.
 */
import { ChartClassic, ChartEx, ChartOptionsError, drawChartExPdf } from "@excel/chart/index";
import { describe, expect, it } from "vitest";

const SQUARE = [
  [0, 0],
  [10, 0],
  [0, 10],
  [-10, 0],
  [0, -10]
];

function topology(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    type: "Topology",
    arcs: [SQUARE],
    objects: {
      countries: {
        type: "GeometryCollection",
        geometries: [{ type: "Polygon", id: "USA", arcs: [[0]] }]
      }
    },
    ...overrides
  };
}

const model = ChartEx.createModel({
  type: "regionMap",
  series: [{ values: "Data!$B$2:$B$2", literalValues: [42], literalCategories: ["USA"] }],
  layout: { projection: "miller", regionLabels: "showAll", geoMappingLevel: "country" }
});

function refusal(run: () => unknown): ChartOptionsError {
  try {
    run();
  } catch (err) {
    expect(err).toBeInstanceOf(ChartOptionsError);
    return err as ChartOptionsError;
  }
  throw new Error("expected a ChartOptionsError");
}

const BAD: Array<[string, Record<string, unknown>, string, string]> = [
  ["a topology that is not an object", { topology: 42 }, "topology", "a TopoJSON topology"],
  ["no objects", { topology: topology({ objects: undefined }) }, "topology.objects", "an object"],
  ["arcs missing", { topology: topology({ arcs: undefined }) }, "topology.arcs", "an array"],
  [
    "an arc that is not a list",
    { topology: topology({ arcs: [SQUARE, 7] }) },
    "topology.arcs[1]",
    "an array"
  ],
  [
    "a transform that is not two finite pairs",
    { topology: topology({ transform: { scale: [1, Number.NaN], translate: [0, 0] } }) },
    "topology.transform.scale",
    "[number, number]"
  ],
  [
    "a collection without geometries",
    { topology: topology({ objects: { countries: { type: "GeometryCollection" } } }) },
    'topology.objects["countries"].geometries',
    "an array"
  ],
  [
    "a polygon whose arcs are not rings of arc indices",
    {
      topology: topology({
        objects: {
          countries: { type: "GeometryCollection", geometries: [{ type: "Polygon", arcs: [0] }] }
        }
      })
    },
    'topology.objects["countries"].geometries[0].arcs[0]',
    "an array"
  ],
  [
    "a point without coordinates",
    { topology: topology({ objects: { countries: { type: "Point" } } }) },
    'topology.objects["countries"].coordinates',
    "[number, number]"
  ]
];

describe("a region map topology is checked as the caller's data", () => {
  it.each(BAD)("ChartEx.toSvg refuses %s", (_label, regionMap, path, expected) => {
    const err = refusal(() =>
      ChartEx.toSvg(model, { regionMap: { objectName: "countries", ...regionMap } as never })
    );
    expect(err.field).toBe(`options.regionMap.${path}`);
    expect(err.message.startsWith(`ChartEx.toSvg: options.regionMap.${path} must be `)).toBe(true);
    expect(err.message).toContain(expected);
  });

  it("names an objectName the topology does not have, and lists the ones it has", () => {
    const err = refusal(() =>
      ChartEx.toSvg(model, { regionMap: { topology: topology(), objectName: "states" } as never })
    );
    expect(err.field).toBe("options.regionMap.objectName");
    expect(err.message).toBe(
      'ChartEx.toSvg: options.regionMap.objectName must name a geometry collection of the topology (it has "countries"), got "states"'
    );
  });

  it("refuses at every entry that takes a regionMap, rather than falling back silently", async () => {
    const regionMap = { topology: topology({ arcs: 5 }), objectName: "countries" } as never;
    const surface = { drawRect: () => surface, drawLine: () => surface, drawText: () => surface };
    for (const [where, run] of [
      ["ChartEx.toPng", () => ChartEx.toPng(model, { regionMap })],
      ["ChartEx.toDrawList", () => ChartEx.toDrawList(model, { regionMap })],
      [
        "drawChartExPdf",
        () =>
          drawChartExPdf(
            surface as never,
            model,
            { x: 0, y: 0, width: 100, height: 100 },
            {
              regionMap
            }
          )
      ]
    ] as const) {
      const err = await Promise.resolve()
        .then(() => run() as unknown)
        .then(
          () => undefined,
          (thrown: unknown) => thrown
        );
      expect(err).toBeInstanceOf(ChartOptionsError);
      expect((err as ChartOptionsError).field).toBe("options.regionMap.topology.arcs");
      expect((err as Error).message.startsWith(`${where}: `)).toBe(true);
    }
  });

  it("Backend.resolveTopologyObject names its own parameters and call", () => {
    const err = refusal(() =>
      ChartClassic.Backend.resolveTopologyObject(
        topology({ arcs: [SQUARE, 7] }) as never,
        "countries"
      )
    );
    expect(err.field).toBe("topology.arcs[1]");
    expect(
      err.message.startsWith("ChartClassic.Backend.resolveTopologyObject: topology.arcs[1] ")
    ).toBe(true);
    const missing = refusal(() =>
      ChartClassic.Backend.resolveTopologyObject(topology() as never, "x")
    );
    expect(missing.field).toBe("objectName");
  });

  it("still draws a valid topology, and still falls back when nothing matches", () => {
    expect(
      ChartEx.toSvg(model, {
        regionMap: { topology: topology(), objectName: "countries" } as never
      })
    ).toContain('data-region-map-mode="topojson"');
    const unmatched = ChartEx.toSvg(model, {
      regionMap: {
        topology: topology({
          objects: {
            countries: {
              type: "GeometryCollection",
              geometries: [{ type: "Polygon", id: "NOWHERE", arcs: [[0]] }]
            }
          }
        }),
        objectName: "countries"
      } as never
    });
    expect(unmatched).not.toContain('data-region-map-mode="topojson"');
    // Unknown geometry types are skipped, as the TopoJSON specification allows (`type: null`).
    expect(
      ChartClassic.Backend.resolveTopologyObject(
        topology({
          objects: { countries: { type: "GeometryCollection", geometries: [{ type: null }] } }
        }) as never,
        "countries"
      )
    ).toHaveLength(1);
  });
});
