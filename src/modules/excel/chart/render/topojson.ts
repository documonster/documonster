/**
 * Minimal, zero-dependency TopoJSON feature extractor.
 *
 * TopoJSON (https://github.com/topojson/topojson-specification) encodes
 * geometries as indices into a shared arc table, optionally with an
 * integer quantisation grid that is dequantised via `transform`. Full
 * support (mesh, merge, presimplify…) belongs to a dedicated library;
 * this module implements only what the `regionMap` renderer needs:
 *
 * 1. Parse a {@link TopologyLike} object produced by any TopoJSON 1.x
 *    emitter (d3, topojson-server, world-atlas bundles).
 * 2. Dequantise integer-encoded arcs back to absolute lon/lat pairs.
 * 3. Resolve a geometry (`Polygon` / `MultiPolygon` / `LineString` /
 *    `MultiLineString` / `Point` / `MultiPoint`) into drawable rings.
 *
 * The shapes here mirror what the TopoJSON 1.x spec emits but keep
 * the types narrow — unknown fields (`bbox`, vendor extensions) are
 * ignored, not stripped. Users who need the full spec should use the
 * upstream `topojson-client` package; doing so does not conflict with
 * this module because both return `[lon, lat]` coordinate arrays in
 * the geometry output.
 */

import { ChartOptionsError } from "@excel/errors";

/** Raw TopoJSON topology as produced by d3 / world-atlas. */
export interface TopologyLike {
  type: "Topology";
  /**
   * Optional quantisation transform: each arc point's integer pair
   * `[qx, qy]` is decoded as
   * `[qx * scale[0] + translate[0], qy * scale[1] + translate[1]]`
   * after delta-undo. Absent → arcs already hold absolute coordinates.
   */
  transform?: { scale: [number, number]; translate: [number, number] };
  /** Shared arc pool; each arc is a list of `[x, y]` integer or float pairs. */
  arcs: Array<Array<[number, number]>>;
  /** Named geometry collections (e.g. `countries`, `states`). */
  objects: Record<string, TopoGeometryCollection | TopoGeometry>;
}

export type TopoGeometry =
  | {
      type: "Polygon";
      arcs: number[][];
      id?: string | number;
      properties?: Record<string, unknown>;
    }
  | {
      type: "MultiPolygon";
      arcs: number[][][];
      id?: string | number;
      properties?: Record<string, unknown>;
    }
  | {
      type: "LineString";
      arcs: number[];
      id?: string | number;
      properties?: Record<string, unknown>;
    }
  | {
      type: "MultiLineString";
      arcs: number[][];
      id?: string | number;
      properties?: Record<string, unknown>;
    }
  | {
      type: "Point";
      coordinates: [number, number];
      id?: string | number;
      properties?: Record<string, unknown>;
    }
  | {
      type: "MultiPoint";
      coordinates: Array<[number, number]>;
      id?: string | number;
      properties?: Record<string, unknown>;
    };

export interface TopoGeometryCollection {
  type: "GeometryCollection";
  geometries: TopoGeometry[];
}

/**
 * A resolved polygon — an ordered ring of `[lon, lat]` pairs in
 * absolute coordinates. Holes (inner rings) are represented as
 * subsequent entries in the outer array.
 */
export type ResolvedRing = Array<[number, number]>;

/**
 * Resolve a named object from a {@link TopologyLike} into a flat list
 * of `{ id, rings }` records. Each `rings[0]` is the outer ring and
 * the remainder are holes. For `LineString` / `MultiLineString`,
 * `rings` contains the open polylines; the caller is expected to
 * treat them as un-closed paths.
 *
 * Throws if the object is missing or not a recognised geometry type —
 * preferring a loud failure over silent "empty map" output, which is
 * almost always a wiring bug on the caller's side.
 */
export function resolveTopologyObject(
  topology: TopologyLike,
  objectName: string
): Array<{
  id: string | number | undefined;
  properties: Record<string, unknown> | undefined;
  geometry: TopoGeometry;
  rings: ResolvedRing[];
}> {
  const obj = topology.objects?.[objectName];
  if (!obj) {
    throw new ChartOptionsError(
      `resolveTopologyObject: topology.objects['${objectName}'] not found (available: ${
        Object.keys(topology.objects ?? {}).join(", ") || "<none>"
      }).`
    );
  }
  if (!Array.isArray(topology.arcs)) {
    // Malformed input — TopoJSON spec requires `arcs` to be an array. We
    // bail out explicitly (rather than throwing inside `decode` below
    // from a `TypeError: Cannot read properties of undefined`) so the
    // caller sees a clear message.
    throw new ChartOptionsError(
      `resolveTopologyObject: topology.arcs is missing or not an array; cannot decode geometry.`
    );
  }
  const rawScale = topology.transform?.scale;
  const rawTranslate = topology.transform?.translate;
  // Defensive: verify both are finite 2-tuples before forwarding to
  // `dequantiseArc`. A single `NaN` / `Infinity` in scale or translate
  // would otherwise propagate through every coordinate.
  const isFinitePair = (v: unknown): v is [number, number] =>
    Array.isArray(v) && v.length === 2 && Number.isFinite(v[0]) && Number.isFinite(v[1]);
  const scale = isFinitePair(rawScale) ? rawScale : undefined;
  const translate = isFinitePair(rawTranslate) ? rawTranslate : undefined;
  // Cache decoded arcs by their positive index. A world-atlas topology
  // typically shares each arc between 2–4 neighbouring countries; with
  // no cache, `dequantiseArc` (an O(n) delta-accumulator) runs that
  // many times per render. The cache turns the inner loop into O(1)
  // for repeated references without changing any semantics — decoded
  // points are read-only from the caller's perspective.
  const decodedArcs: Map<number, Array<[number, number]>> = new Map();
  const decode = (arcIndex: number): Array<[number, number]> => {
    const reverse = arcIndex < 0;
    // Topojson spec: a negative index means the 1s-complement of the arc's
    // positive index, used to indicate reverse traversal.
    const positive = reverse ? ~arcIndex : arcIndex;
    let points = decodedArcs.get(positive);
    if (!points) {
      const rawArc = topology.arcs[positive];
      if (!rawArc) {
        return [];
      }
      points = dequantiseArc(rawArc, scale, translate);
      decodedArcs.set(positive, points);
    }
    // Always return a fresh array on reverse — callers mutate the
    // returned array in some code paths (`joinRing` concat), and we
    // don't want the cached copy to be mutated.
    return reverse ? points.slice().reverse() : points.slice();
  };

  const geometries: TopoGeometry[] =
    obj.type === "GeometryCollection" ? obj.geometries : [obj as TopoGeometry];

  const results: Array<{
    id: string | number | undefined;
    properties: Record<string, unknown> | undefined;
    geometry: TopoGeometry;
    rings: ResolvedRing[];
  }> = [];

  for (const geom of geometries) {
    const rings: ResolvedRing[] = [];
    switch (geom.type) {
      case "Polygon": {
        for (const ring of geom.arcs) {
          rings.push(joinRing(ring, decode));
        }
        break;
      }
      case "MultiPolygon": {
        for (const poly of geom.arcs) {
          for (const ring of poly) {
            rings.push(joinRing(ring, decode));
          }
        }
        break;
      }
      case "LineString": {
        rings.push(joinRing(geom.arcs, decode));
        break;
      }
      case "MultiLineString": {
        for (const line of geom.arcs) {
          rings.push(joinRing(line, decode));
        }
        break;
      }
      case "Point": {
        // Point / MultiPoint coordinates live in the same quantised
        // grid as arc coordinates when the topology carries a
        // `transform`. Previously the coordinates were pushed raw,
        // so Point features from a d3 / topojson-server-quantised
        // topology plotted at integer pixel positions (e.g.
        // `[5243, 3012]`) instead of their lon/lat values. Dequantise
        // through the same pipeline arcs use to keep the output space
        // consistent. Point coordinates are NOT delta-encoded — each
        // is absolute — so wrap the single pair in a synthetic "arc"
        // and only apply the scale/translate step.
        if (scale && translate) {
          const [x, y] = geom.coordinates;
          if (Number.isFinite(x) && Number.isFinite(y)) {
            rings.push([[x * scale[0] + translate[0], y * scale[1] + translate[1]]]);
          }
        } else {
          rings.push([geom.coordinates]);
        }
        break;
      }
      case "MultiPoint": {
        if (scale && translate) {
          const ring: Array<[number, number]> = [];
          for (const [x, y] of geom.coordinates) {
            if (Number.isFinite(x) && Number.isFinite(y)) {
              ring.push([x * scale[0] + translate[0], y * scale[1] + translate[1]]);
            }
          }
          if (ring.length > 0) {
            rings.push(ring);
          }
        } else {
          for (const pt of geom.coordinates) {
            rings.push([pt]);
          }
        }
        break;
      }
      default: {
        // Unknown geometry type — skip instead of throwing so a
        // collection containing a few unsupported entries still
        // renders the rest. The caller can detect empty output and
        // decide whether to warn.
        break;
      }
    }
    results.push({
      id: geom.id,
      properties:
        "properties" in geom && geom.properties
          ? (geom.properties as Record<string, unknown>)
          : undefined,
      geometry: geom,
      rings
    });
  }
  return results;
}

/**
 * Concatenate a list of arc indices into one ring, removing the
 * duplicate join points where two arcs share an endpoint. Follows the
 * TopoJSON spec: when joining arc B to arc A, the first point of B is
 * identical to the last point of A and must be dropped.
 */
function joinRing(
  arcIndices: number[],
  decode: (arcIndex: number) => Array<[number, number]>
): ResolvedRing {
  const out: Array<[number, number]> = [];
  for (const arcIndex of arcIndices) {
    const pts = decode(arcIndex);
    if (pts.length === 0) {
      continue;
    }
    // "First non-empty arc" is when `out` is still empty — push all
    // points including the starting one. For every subsequent arc, the
    // join point (pts[0]) is identical to `out[-1]` by TopoJSON spec
    // §2.1.4 and must be dropped.
    //
    // Using `out.length === 0` instead of `i === 0` matters when the
    // first arc decoded empty: a naive "i === 0" check would drop the
    // starting point of the first real (non-empty) arc, shifting the
    // whole ring by one vertex.
    if (out.length === 0) {
      out.push(...pts);
    } else {
      for (let j = 1; j < pts.length; j++) {
        out.push(pts[j]);
      }
    }
  }
  return out;
}

/**
 * Undo delta encoding and apply the quantisation transform. Each
 * consecutive `[dx, dy]` pair in the source represents an offset
 * from the previous point; the first pair is absolute.
 */
function dequantiseArc(
  arc: Array<[number, number]>,
  scale: [number, number] | undefined,
  translate: [number, number] | undefined
): Array<[number, number]> {
  if (!scale || !translate) {
    // Arcs already hold floating-point absolute coordinates.
    //
    // Filter out non-finite points so a single malformed pair doesn't
    // propagate `NaN` into downstream `[x, y]` consumers (SVG coord
    // printing, PDF path construction) where it silently produces
    // broken geometry.
    return arc
      .filter(pair => Array.isArray(pair) && Number.isFinite(pair[0]) && Number.isFinite(pair[1]))
      .map(([x, y]) => [x, y]);
  }
  const out: Array<[number, number]> = [];
  let qx = 0;
  let qy = 0;
  for (const pair of arc) {
    // Defensive check: any non-finite delta would permanently poison
    // the running accumulator, so we abort the whole arc rather than
    // emit progressively worse `NaN` vertices.
    if (!Array.isArray(pair) || !Number.isFinite(pair[0]) || !Number.isFinite(pair[1])) {
      return [];
    }
    const [dx, dy] = pair;
    qx += dx;
    qy += dy;
    out.push([qx * scale[0] + translate[0], qy * scale[1] + translate[1]]);
  }
  return out;
}
