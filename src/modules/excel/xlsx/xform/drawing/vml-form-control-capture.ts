import type {
  FormCheckboxModel,
  FormControlAnchor,
  LegacyControlModel
} from "@excel/core/form-control";
import { formCheckboxStateFromValue } from "@excel/core/form-control";
import { controlPlacement } from "@excel/core/form-control-placement";
import type { EchoNode } from "@excel/xlsx/xform/drawing/generic-echo-xform";
import { EMU_PER_PX } from "@utils/units";
import type { ParseOpenTag } from "@xml/types";

/** `<x:ClientData>` leaves whose text a checkbox needs. */
const LEAVES = new Set(["x:Anchor", "x:FmlaLink", "x:Checked", "x:PrintObject"]);

/**
 * Collects what a VML `v:shape` says about a legacy form control while the shape is parsed for a note.
 *
 * Whether a shape is a note or a control is only known at `<x:ClientData ObjectType>`, so the drawing
 * xform runs this beside the note parser and decides when the shape closes. Only a checkbox becomes a
 * model; the rest is reported through {@link objectType} so the caller can tell a note from a control.
 */
export class VmlFormControlCapture {
  objectType?: string;
  private readonly leaves = new Map<string, string>();
  private leaf?: string;
  private noThreeD = false;
  /** `x:MoveWithCells` / `x:SizeWithCells`, present when the object does *not* move / size with cells. */
  private fixedPosition = false;
  private fixedSize = false;
  private divDepth = 0;
  private readonly lines: string[] = [];
  private line = "";

  constructor(private readonly attributes: Record<string, string>) {}

  /** No `ObjectType`, or `Note`: the shape is a comment's box. */
  get isNote(): boolean {
    return this.objectType === undefined || this.objectType === "Note";
  }

  open(node: ParseOpenTag): void {
    switch (node.name) {
      case "x:ClientData":
        this.objectType = node.attributes.ObjectType;
        break;
      case "x:NoThreeD":
        this.noThreeD = true;
        break;
      case "x:MoveWithCells":
        this.fixedPosition = true;
        break;
      case "x:SizeWithCells":
        this.fixedSize = true;
        break;
      case "div":
        this.divDepth++;
        break;
      case "br":
        if (this.divDepth > 0) {
          this.line += "\n";
        }
        break;
      default:
        if (LEAVES.has(node.name)) {
          this.leaf = node.name;
          this.leaves.set(node.name, "");
        }
    }
  }

  text(text: string): void {
    if (this.leaf !== undefined) {
      this.leaves.set(this.leaf, this.leaves.get(this.leaf)! + text);
    } else if (this.divDepth > 0) {
      this.line += text;
    }
  }

  close(name: string): void {
    if (name === this.leaf) {
      this.leaf = undefined;
    } else if (name === "div" && this.divDepth > 0) {
      this.divDepth--;
      if (this.divDepth === 0) {
        this.lines.push(this.line);
        this.line = "";
      }
    }
  }

  /** The shape's `<x:Anchor>`, as EMU corners. */
  anchor(): { tl: FormControlAnchor; br: FormControlAnchor } | undefined {
    return parseVmlAnchor(this.leaves.get("x:Anchor"));
  }

  /** The checkbox this shape draws, or `undefined` when it has no usable anchor or shape id. */
  toCheckbox(): FormCheckboxModel | undefined {
    const shapeId = vmlShapeId(this.attributes["o:spid"]) ?? vmlShapeId(this.attributes.id);
    const anchor = this.anchor();
    if (shapeId === undefined || anchor === undefined) {
      return undefined;
    }
    const link = this.leaves.get("x:FmlaLink")?.trim();
    const print = this.leaves.get("x:PrintObject")?.trim();
    const id = this.attributes.id;
    return {
      shapeId,
      ctrlPropId: 0,
      ...anchor,
      ...(link ? { link } : {}),
      checked: formCheckboxStateFromValue(this.leaves.get("x:Checked")),
      text: this.lines.join("\n"),
      noThreeD: this.noThreeD,
      // The sheet's `<controls>` entry, when there is one, says it too and wins.
      editAs: controlPlacement(!this.fixedPosition, !this.fixedSize),
      // `x:PrintObject` defaults to true when absent, as Excel omits it for a printed control.
      print: print === undefined ? true : !/^(false|f|0)$/i.test(print),
      // Excel names the shape after the control (`Check_x0020_Box_x0020_1`); this library's shapes carry
      // only their id. The sheet's `<controls>` entry, when there is one, has the name verbatim and wins.
      ...(id && vmlShapeId(id) === undefined ? { name: decodeVmlName(id) } : {}),
      ...(this.attributes.style ? { vmlStyle: this.attributes.style } : {})
    };
  }
}

/**
 * A control this library does not model, kept whole: `shape` is the `<v:shape>` element as captured, and
 * the rest is what a write needs to place it — its id, its type and the corners of its bridge anchor.
 * `undefined` when the shape carries no shape id, which nothing could refer to.
 */
export function legacyControlFromVml(
  capture: VmlFormControlCapture,
  shape: EchoNode
): LegacyControlModel | undefined {
  const text = (value: unknown) => (value === undefined ? undefined : String(value));
  const id = text(shape.attrs.id);
  const shapeId = vmlShapeId(text(shape.attrs["o:spid"])) ?? vmlShapeId(id);
  if (shapeId === undefined || capture.objectType === undefined) {
    return undefined;
  }
  const anchor = capture.anchor();
  return {
    shapeId,
    objectType: capture.objectType,
    ...(id && vmlShapeId(id) === undefined ? { name: decodeVmlName(id) } : {}),
    ...anchor,
    vmlShape: shape
  };
}

/** `_x0000_s1025` → 1025. */
function vmlShapeId(value: string | undefined): number | undefined {
  const match = value === undefined ? null : /^_x0000_s(\d+)$/.exec(value);
  return match ? parseInt(match[1], 10) : undefined;
}

/** `Check_x0020_Box_x0020_1` → `Check Box 1`. */
function decodeVmlName(value: string): string {
  return value.replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex: string) =>
    String.fromCharCode(parseInt(hex, 16))
  );
}

/**
 * `<x:Anchor>` — column, pixel offset, row, pixel offset for each corner — as the model's EMU corners.
 */
function parseVmlAnchor(
  value: string | undefined
): { tl: FormControlAnchor; br: FormControlAnchor } | undefined {
  const numbers = (value ?? "").split(",").map(part => parseInt(part.trim(), 10));
  if (numbers.length !== 8 || numbers.some(number => !Number.isFinite(number))) {
    return undefined;
  }
  const [col, colOff, row, rowOff, endCol, endColOff, endRow, endRowOff] = numbers;
  return {
    tl: { col, colOff: colOff * EMU_PER_PX, row, rowOff: rowOff * EMU_PER_PX },
    br: { col: endCol, colOff: endColOff * EMU_PER_PX, row: endRow, rowOff: endRowOff * EMU_PER_PX }
  };
}
