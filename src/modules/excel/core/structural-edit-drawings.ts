/**
 * Drawings and page breaks through a row or column edit — the part of `structural-edit` that is about where things
 * sit rather than what they name.
 *
 * Every drawing follows its own placement, as Excel moves objects (see `moveAnchorCorners`): pictures, charts,
 * shapes, drawings preserved from a file, and form controls — a checkbox by its `editAs`, which is Excel's
 * "move but don't size with cells" unless chosen otherwise, so it stays beside the cell it is linked to.
 *
 * A manual page break belongs to the line after it: a break below row 10 is where row 11 starts a page, so it moves
 * with row 11 and goes when row 11 is deleted.
 */
import type { AnchorData } from "@excel/core/anchor";
import type {
  FormCheckboxModel,
  FormControlAnchor,
  LegacyControlModel
} from "@excel/core/form-control";
import { formCheckboxPlacement, legacyControlPlacement } from "@excel/core/form-control-placement";
import type { ImageModel } from "@excel/core/image";
import { imageCreate, imageModel } from "@excel/core/image";
import type { WorksheetData } from "@excel/core/worksheet-core";
import type { AnchorCorner, AnchorPlacement, LineEdit } from "@excel/utils/reference-rewrite";
import { editLine, moveAnchorCorners } from "@excel/utils/reference-rewrite";
import type { EchoNode } from "@excel/xlsx/xform/drawing/generic-echo-xform";

/** Move every drawing on `ws` and its page breaks through `edit`, which is on `ws`. */
export function moveSheetDrawings(ws: WorksheetData, edit: LineEdit): void {
  for (const image of ws._media) {
    const range = image.type === "image" ? image.range : undefined;
    if (range && !range.pos) {
      // A picture's two-cell anchor without `editAs` is written as `oneCell`.
      moveNative(range.tl, range.br, placementOf(range.editAs, "oneCell"), edit);
    }
  }
  for (const chart of ws._charts) {
    if (!chart.range.pos) {
      moveNative(chart.range.tl, chart.range.br, placementOf(chart.range.editAs, "twoCell"), edit);
    }
  }
  type DrawingAnchor = {
    range?: { tl?: AnchorData; br?: AnchorData; editAs?: string; pos?: unknown };
    graphicFrame?: unknown;
  };
  for (const anchor of (ws._drawing as { anchors?: DrawingAnchor[] } | undefined)?.anchors ?? []) {
    const range = anchor.range;
    if (range?.tl && !range.pos) {
      // As the drawing writer does: a graphic frame (a chart) defaults to `twoCell`, anything else to `oneCell`.
      const fallback = anchor.graphicFrame ? "twoCell" : "oneCell";
      moveNative(range.tl, range.br, placementOf(range.editAs, fallback), edit);
    }
  }
  ws._shapes = ws._shapes.map(shape => {
    let range: Extract<ImageModel, { type: "image" }>["range"];
    try {
      const probe = imageCreate(ws, { type: "image", imageId: "", range: shape.range });
      range = (imageModel(probe) as Extract<ImageModel, { type: "image" }>).range;
    } catch {
      return shape; // an anchor the writer skips too
    }
    if (!range || range.pos) {
      return shape;
    }
    const tl = { ...range.tl };
    const br = range.br ? { ...range.br } : undefined;
    const editAs = placementOf(range.editAs, "oneCell");
    moveNative(tl, br, editAs, edit);
    const anchor = (a: typeof tl) => ({ ...a, col: a.nativeCol, row: a.nativeRow });
    return {
      ...shape,
      range: {
        tl: anchor(tl),
        ...(br ? { br: anchor(br) } : {}),
        ...(range.ext ? { ext: range.ext as { width: number; height: number } } : {}),
        editAs
      }
    };
  });
  for (const checkbox of ws.formControls ?? []) {
    moveControlCorners(checkbox.model, formCheckboxPlacement(checkbox.model), edit);
  }
  for (const control of ws._legacyControls ?? []) {
    moveLegacyControl(control, edit);
  }
  const breaks = edit.kind === "rows" ? "rowBreaks" : "colBreaks";
  ws[breaks] = (ws[breaks] ?? []).flatMap(entry => {
    const line = editLine(entry.id + 1, edit);
    return line === undefined ? [] : [{ ...entry, id: line - 1 }];
  });
}

function placementOf(editAs: string | undefined, fallback: AnchorPlacement): AnchorPlacement {
  return editAs === "twoCell" || editAs === "oneCell" || editAs === "absolute" ? editAs : fallback;
}

/** Move a DrawingML anchor's corners (`nativeRow`/`nativeRowOff` or the column pair) in place. */
function moveNative(
  tl: AnchorData,
  br: AnchorData | undefined,
  placement: AnchorPlacement,
  edit: LineEdit
): void {
  const [index, offset] =
    edit.kind === "rows"
      ? (["nativeRow", "nativeRowOff"] as const)
      : (["nativeCol", "nativeColOff"] as const);
  const corner = (a: AnchorData): AnchorCorner => ({ index: a[index], offset: a[offset] ?? 0 });
  const moved = moveAnchorCorners(corner(tl), br && corner(br), placement, edit);
  tl[index] = moved.tl.index;
  tl[offset] = moved.tl.offset;
  if (br && moved.br) {
    br[index] = moved.br.index;
    br[offset] = moved.br.offset;
  }
}

// ============================================================================
// Preserved form controls
// ============================================================================

function child(node: EchoNode | undefined, tag: string): EchoNode | undefined {
  return node?.children.find(c => c.tag === tag);
}

function textOf(node: EchoNode): string {
  const parts = node.content ?? (node.text === undefined ? [] : [node.text]);
  return parts.filter(part => typeof part === "string").join("");
}

function setText(node: EchoNode, text: string): void {
  node.content = [text];
  node.text = undefined;
}

/** The `<x:ClientData>` of a control's VML shape. */
function clientData(control: LegacyControlModel): EchoNode | undefined {
  return child(control.vmlShape, "x:ClientData");
}

/** Move a control's EMU corners — a checkbox's, or the bridge corners of a preserved control — in place. */
function moveControlCorners(
  control: Pick<FormCheckboxModel, "tl" | "br"> | Pick<LegacyControlModel, "tl" | "br">,
  placement: AnchorPlacement,
  edit: LineEdit
): void {
  if (placement === "absolute" || !control.tl || !control.br) {
    return;
  }
  const [index, offset] =
    edit.kind === "rows" ? (["row", "rowOff"] as const) : (["col", "colOff"] as const);
  const corner = (a: FormControlAnchor): AnchorCorner => ({ index: a[index], offset: a[offset] });
  const moved = moveAnchorCorners(corner(control.tl), corner(control.br), placement, edit);
  control.tl = { ...control.tl, [index]: moved.tl.index, [offset]: moved.tl.offset };
  control.br = { ...control.br, [index]: moved.br!.index, [offset]: moved.br!.offset };
}

/**
 * Move a preserved control in each of the three places its position is written: the corners the drawing bridge is
 * built from, the VML `<x:Anchor>` (column, pixel offset, row, pixel offset per corner) and the `<control>` entry's
 * `<from>`/`<to>`. Each is moved from its own numbers, so the offsets each one states are kept.
 */
function moveLegacyControl(control: LegacyControlModel, edit: LineEdit): void {
  const placement = legacyControlPlacement(control);
  if (placement === "absolute") {
    return;
  }
  const rows = edit.kind === "rows";
  moveControlCorners(control, placement, edit);

  const vmlAnchor = child(clientData(control), "x:Anchor");
  if (vmlAnchor) {
    const numbers = textOf(vmlAnchor)
      .split(",")
      .map(part => parseInt(part.trim(), 10));
    if (numbers.length === 8 && numbers.every(Number.isFinite)) {
      const at = rows ? 2 : 0;
      const moved = moveAnchorCorners(
        { index: numbers[at], offset: numbers[at + 1] },
        { index: numbers[at + 4], offset: numbers[at + 5] },
        placement,
        edit
      );
      numbers[at] = moved.tl.index;
      numbers[at + 1] = moved.tl.offset;
      numbers[at + 4] = moved.br!.index;
      numbers[at + 5] = moved.br!.offset;
      setText(vmlAnchor, numbers.join(", "));
    }
  }

  const anchor = child(child(control.control, "controlPr"), "anchor");
  const from = child(anchor, "from");
  const to = child(anchor, "to");
  const [indexTag, offsetTag] = rows ? ["xdr:row", "xdr:rowOff"] : ["xdr:col", "xdr:colOff"];
  const read = (corner: EchoNode | undefined): AnchorCorner | undefined => {
    const index = child(corner, indexTag);
    const offset = child(corner, offsetTag);
    if (!index || !offset) {
      return undefined;
    }
    return { index: parseInt(textOf(index), 10), offset: parseInt(textOf(offset), 10) };
  };
  const start = read(from);
  const end = read(to);
  if (start && end && Number.isFinite(start.index + start.offset + end.index + end.offset)) {
    const moved = moveAnchorCorners(start, end, placement, edit);
    setText(child(from, indexTag)!, String(moved.tl.index));
    setText(child(from, offsetTag)!, String(moved.tl.offset));
    setText(child(to, indexTag)!, String(moved.br!.index));
    setText(child(to, offsetTag)!, String(moved.br!.offset));
  }
}

/** The VML leaves and `ctrlProp` attributes that hold a cell reference. */
const VML_REFERENCE_LEAVES = ["x:FmlaLink", "x:FmlaRange", "x:FmlaTxbx", "x:FmlaGroup"];
const CTRL_PROP_REFERENCE = /\b(fmlaLink|fmlaRange|fmlaTxbx|fmlaGroup)="([^"]*)"/g;

const decode = (text: string): string =>
  text
    .replace(/&apos;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
const encode = (text: string): string =>
  text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");

/**
 * Rewrite the cells form controls are linked to — a checkbox's `link`, and a preserved control's linked cell, input
 * range and text link, which it states twice: in its VML and in its `ctrlProp` part.
 */
export function rewriteControlLinks(ws: WorksheetData, formula: (text: string) => string): void {
  for (const checkbox of ws.formControls ?? []) {
    if (checkbox.model.link) {
      checkbox.model.link = formula(checkbox.model.link);
    }
  }
  for (const control of ws._legacyControls ?? []) {
    const data = clientData(control);
    for (const tag of VML_REFERENCE_LEAVES) {
      const leaf = child(data, tag);
      if (leaf) {
        const text = textOf(leaf);
        const next = formula(text);
        if (next !== text) {
          setText(leaf, next);
        }
      }
    }
    if (control.ctrlPropXml !== undefined) {
      control.ctrlPropXml = control.ctrlPropXml.replace(
        CTRL_PROP_REFERENCE,
        (whole, name: string, value: string) => {
          const next = formula(decode(value));
          return next === decode(value) ? whole : `${name}="${encode(next)}"`;
        }
      );
    }
  }
}
