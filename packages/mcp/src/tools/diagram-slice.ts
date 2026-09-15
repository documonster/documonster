/**
 * Cutting a tall diagram across several pages.
 *
 * A page has one width and a fixed height, and a diagram narrowed enough to be read at all
 * can easily be taller than the second of those. Scaling it down again to fit is what this
 * module exists to avoid: it trades the legibility that the narrowing just bought back, and
 * for a deep flowchart it trades a lot of it.
 *
 * So the diagram is cut instead, and the whole problem is *where*. A row of pixels that is
 * entirely background does not exist in a top-to-bottom flowchart — the links between ranks
 * cross every gap — so "find the blank rows" finds nothing. Cutting a link is fine, though:
 * the line continues at the top of the next page, which is how any split figure reads. What
 * must not be cut is a **box or a piece of text**, because half a word is not readable at
 * either end and half a node is not obviously one node.
 *
 * That distinction is available from the display list itself, and from nothing else: a
 * `rect`, `ellipse`, `sector` or `text` is solid, and a `polyline` or `path` is solid exactly
 * when it is *filled* — because `closed` does not separate a connector from a node, a rhombus,
 * hexagon, parallelogram, stadium and cylinder all being drawn as filled polylines and paths
 * rather than rects. So the forbidden bands are collected by walking the list with
 * {@link renderDrawList} through a surface that measures instead of painting — which means the
 * transforms, and therefore the geometry, are the engine's own rather than a second
 * implementation of them, and it works for all twenty-one diagram types without knowing which
 * one it has.
 */

import type {
  DrawClip,
  DrawList,
  DrawPaint,
  DrawPathCommand,
  DrawPoint,
  DrawSurface,
  DrawTextLine,
  DrawTextStyle
} from "documonster/draw";
import { measureText, renderDrawList } from "documonster/draw";

/** A closed interval along the flow axis that a cut must not fall inside. */
interface Band {
  readonly top: number;
  readonly bottom: number;
}

/**
 * A surface that draws nothing and records where the solid marks are.
 *
 * Lines are deliberately ignored. They are the marks a cut may pass through, and treating
 * them as obstacles would forbid every row of a flowchart.
 */
class BandCollector implements DrawSurface {
  readonly bands: Band[] = [];
  private readonly clips: DrawClip[] = [];

  private add(top: number, bottom: number): void {
    const clip = this.clips.at(-1);
    if (clip !== undefined) {
      top = Math.max(top, clip.y);
      bottom = Math.min(bottom, clip.y + clip.height);
    }
    if (bottom > top) {
      this.bands.push({ top, bottom });
    }
  }

  rect(_x: number, y: number, _width: number, height: number, _rx: number, _paint: DrawPaint) {
    this.add(y, y + height);
  }

  ellipse(_cx: number, cy: number, _rx: number, ry: number, _paint: DrawPaint) {
    this.add(cy - ry, cy + ry);
  }

  sector(
    _cx: number,
    cy: number,
    radius: number,
    _innerRadius: number,
    _startAngle: number,
    _endAngle: number,
    _paint: DrawPaint
  ) {
    // The whole disc rather than the swept arc: a pie slice is a solid, and over-reserving a
    // band can only cost a cut position, while under-reserving one cuts through the picture.
    this.add(cy - radius, cy + radius);
  }

  polyline(points: readonly DrawPoint[], _closed: boolean, paint: DrawPaint) {
    this.stroke(
      points.map(point => point.y),
      paint
    );
  }

  path(commands: readonly DrawPathCommand[], paint: DrawPaint) {
    const ys: number[] = [];
    for (const command of commands) {
      if (command.op === "close") {
        continue;
      }
      ys.push(command.y);
      if (command.op === "cubic") {
        // The control points bound the curve, so they are the conservative reading of where it
        // might reach. A flattened curve would be tighter and is not worth the arithmetic here.
        ys.push(command.y1, command.y2);
      }
    }
    this.stroke(ys, paint);
  }

  /**
   * A polyline or a path, which may be either a connector or the outline of a node.
   *
   * The discriminator is the fill, and it has to be: `closed` does not separate them — an
   * arrowhead is a closed triangle and so is nothing else — while a rhombus, a hexagon, a
   * parallelogram, a stadium and a cylinder are all drawn as filled polylines or paths rather
   * than as rects. Ignoring every polyline let a cut fall straight through a diamond.
   *
   * An unfilled stroke is a line, and a line may be cut: it continues at the top of the next
   * page, which is how any split figure reads.
   */
  private stroke(ys: readonly number[], paint: DrawPaint): void {
    if (paint.fill === undefined || ys.length === 0) {
      return;
    }
    this.add(Math.min(...ys), Math.max(...ys));
  }

  text(x: number, y: number, lines: readonly DrawTextLine[], style: DrawTextStyle, rotate: number) {
    const angle = (rotate * Math.PI) / 180;
    const sin = Math.sin(angle);
    const cos = Math.cos(angle);
    const ys: number[] = [];
    for (const line of lines) {
      const width = measureText(line.text, style);
      const lineX = line.x ?? x;
      const left =
        style.anchor === "middle"
          ? lineX - width / 2
          : style.anchor === "end"
            ? lineX - width
            : lineX;
      const baseline = y + line.dy;
      const corners = [
        { x: left, y: baseline - style.size },
        { x: left + width, y: baseline - style.size },
        { x: left, y: baseline + style.size * 0.3 },
        { x: left + width, y: baseline + style.size * 0.3 }
      ];
      for (const corner of corners) {
        const dx = corner.x - x;
        const dy = corner.y - y;
        ys.push(y + dx * sin + dy * cos);
      }
    }
    if (ys.length > 0) {
      this.add(Math.min(...ys), Math.max(...ys));
    }
  }

  pushClip(clip: DrawClip) {
    const parent = this.clips.at(-1);
    if (parent === undefined) {
      this.clips.push(clip);
      return;
    }
    const top = Math.max(parent.y, clip.y);
    const bottom = Math.min(parent.y + parent.height, clip.y + clip.height);
    this.clips.push({ x: 0, y: top, width: 0, height: Math.max(0, bottom - top) });
  }

  popClip() {
    this.clips.pop();
  }
}

export interface SliceCut {
  /** Where the slice starts, in display points from the top. */
  readonly top: number;
  /** Where it ends. */
  readonly bottom: number;
  /** Whether this slice had to be cut through a box or a word because nothing else was free. */
  readonly forced: boolean;
}

/**
 * Choose the cuts that carve a diagram of `height` points into pieces of at most `limit`.
 *
 * Greedy from the top and deliberately so: taking the *lowest* legal cut at or above each
 * limit uses as much of every page as it can, which is what keeps the piece count minimal.
 *
 * `scale` maps display points to the diagram's own units, because the bands are measured in
 * the display list's coordinates while the page budget is in points.
 */
export function planSlices(
  list: DrawList,
  scale: number,
  limit: number,
  minimum = 48
): readonly SliceCut[] {
  const height = list.height * scale;
  if (height <= limit) {
    return [{ top: 0, bottom: height, forced: false }];
  }

  const collector = new BandCollector();
  // Remove only a root canvas plate. Deciding from a band's height after the
  // walk discarded any legitimate shape spanning the canvas and also lost the
  // x/width evidence that distinguished it from a background.
  const first = list.children[0];
  const background =
    first?.kind === "rect" &&
    first.x === 0 &&
    first.y === 0 &&
    first.width === list.width &&
    first.height === list.height
      ? 0
      : -1;
  renderDrawList(
    background === -1
      ? list
      : { ...list, children: list.children.filter((_, index) => index !== background) },
    collector
  );
  const raw = collector.bands.map(band => ({
    top: band.top * scale,
    bottom: band.bottom * scale
  }));

  const content = mergeBands(raw);

  // An obstacle taller than a whole slice cannot be avoided by any choice of cut, so letting it
  // block one only forces a worse position. It still counts for *reporting*: cutting through a
  // node that big is a real compromise and the caller is told.
  const blocking = content.filter(band => band.bottom - band.top < limit);

  const inside = (row: number): boolean =>
    content.some(band => row > band.top && row < band.bottom);

  const cuts: SliceCut[] = [];
  let top = 0;
  // Bounded by the piece count: each turn advances `top` by at least `minimum`.
  while (height - top > limit) {
    const target = top + limit;
    // Do not leave a postage-stamp final image. If the nominal cut would leave
    // less than `minimum`, search earlier so both pieces remain useful.
    const latest = height - target < minimum ? height - minimum : target;
    const free = highestFreeRowAtOrBelow(blocking, latest, top + minimum);
    const at = free ?? target;
    cuts.push({ top, bottom: at, forced: inside(at) });
    top = at;
  }
  cuts.push({ top, bottom: height, forced: false });
  return cuts;
}

/** Sorted, non-overlapping bands. */
function mergeBands(bands: readonly Band[]): readonly Band[] {
  const sorted = [...bands].sort((a, b) => a.top - b.top);
  const out: Band[] = [];
  for (const band of sorted) {
    const last = out.at(-1);
    if (last !== undefined && band.top <= last.bottom) {
      out[out.length - 1] = { top: last.top, bottom: Math.max(last.bottom, band.bottom) };
    } else {
      out.push({ ...band });
    }
  }
  return out;
}

/**
 * The lowest row at or below `target` that is inside no band, and not above `floor`.
 *
 * Searched downwards from `target` because a cut as late as possible fills the page. When
 * `target` itself is inside a band the search continues from that band's top, since every row
 * between them is inside it too.
 */
function highestFreeRowAtOrBelow(
  bands: readonly Band[],
  target: number,
  floor: number
): number | undefined {
  let row = target;
  while (row >= floor) {
    const hit = bands.find(band => row > band.top && row < band.bottom);
    if (hit === undefined) {
      return row;
    }
    // Just above the obstacle. A band is a closed interval, so its own top is free.
    row = hit.top;
  }
  return undefined;
}
