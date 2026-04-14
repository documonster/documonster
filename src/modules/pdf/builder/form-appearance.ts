/**
 * Form field appearance stream generation.
 *
 * Generates visual appearance streams for PDF form fields so that field
 * values are visible in all PDF viewers — even those that do not honor
 * the `/NeedAppearances` flag.
 *
 * @see PDF Reference 1.7, §12.5.5 — Appearance Streams
 * @see PDF Reference 1.7, §12.7.4 — Field Appearance
 */

import { pdfNumber } from "../core/pdf-object";
import { PdfContentStream } from "../core/pdf-stream";

// =============================================================================
// Types
// =============================================================================

/** Options for generating a text field appearance stream. */
export interface TextFieldAppearanceOptions {
  /** The text value to display. */
  value: string;
  /** Widget annotation rectangle [x1, y1, x2, y2]. */
  rect: number[];
  /** Font size in points. 0 or omitted = auto-size to fit the field height. */
  fontSize?: number;
  /** Font resource name to reference (e.g. "Helv"). */
  fontName?: string;
  /** Text alignment within the field. */
  alignment?: "left" | "center" | "right";
}

// =============================================================================
// Text Field Appearance
// =============================================================================

/** Default padding inside the field rectangle (in points). */
const FIELD_PADDING = 2;

/**
 * Generate an appearance stream for a text form field.
 *
 * Builds a minimal content stream that clips to the widget rect and
 * draws the field value using the specified (or default) font.
 *
 * @returns The raw stream bytes and a resources dictionary string.
 */
export function generateTextFieldAppearance(options: TextFieldAppearanceOptions): {
  stream: Uint8Array;
  resources: string;
} {
  const { value, rect, alignment = "left" } = options;
  const fontName = options.fontName ?? "Helv";

  // Derive bounding box dimensions from the widget rect
  const x1 = rect[0] ?? 0;
  const y1 = rect[1] ?? 0;
  const x2 = rect[2] ?? x1;
  const y2 = rect[3] ?? y1;
  const width = Math.abs(x2 - x1);
  const height = Math.abs(y2 - y1);

  // Auto-size font: use ~70% of the field height, clamped to a reasonable range
  let fontSize = options.fontSize ?? 0;
  if (fontSize <= 0) {
    fontSize = Math.max(4, Math.min(height * 0.7, 20));
  }

  // Approximate text width using 0.5 * fontSize per character (conservative estimate
  // for Helvetica). This is intentionally simple — the appearance is a best-effort
  // rendering; the authoritative value is in /V.
  const approxCharWidth = fontSize * 0.5;
  const textWidth = value.length * approxCharWidth;

  // Compute horizontal offset based on alignment
  let tx: number;
  const usableWidth = width - 2 * FIELD_PADDING;
  switch (alignment) {
    case "center":
      tx = FIELD_PADDING + Math.max(0, (usableWidth - textWidth) / 2);
      break;
    case "right":
      tx = FIELD_PADDING + Math.max(0, usableWidth - textWidth);
      break;
    default:
      tx = FIELD_PADDING;
      break;
  }

  // Vertical centering: place baseline so the text is roughly centered
  // Approximate ascent ≈ 0.75 * fontSize, descent ≈ -0.25 * fontSize
  const ascent = fontSize * 0.75;
  const descent = fontSize * 0.25;
  const textHeight = ascent + descent;
  const ty = (height - textHeight) / 2 + descent;

  // Build the content stream
  const cs = new PdfContentStream();
  // Clip to the bounding box (form XObject coordinate space: 0,0 to width,height)
  cs.rect(0, 0, width, height).clip().endPath();
  // Draw the text
  cs.beginText().setFont(fontName, fontSize).moveText(tx, ty).showText(value).endText();

  const stream = cs.toUint8Array();

  // Build a minimal resources dict referencing a standard Helvetica font.
  // The font is declared inline as a Type1 font so the appearance stream
  // is self-contained (no dependency on the page's resources).
  const resources = `<< /Font << /${fontName} << /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >> >> >>`;

  return { stream, resources };
}

// =============================================================================
// Checkbox Appearance
// =============================================================================

/**
 * Generate appearance streams for a checkbox form field.
 *
 * Returns two streams:
 * - `streamOn`:  draws a checkmark (✓-like shape) inside the rect
 * - `streamOff`: empty appearance (blank field)
 *
 * @param checked - Whether to generate the checked or unchecked variant
 *                  (both are always returned; `checked` is ignored — both
 *                  on and off streams are produced).
 * @param rect - Widget annotation rectangle [x1, y1, x2, y2].
 */
export function generateCheckboxAppearance(
  _checked: boolean,
  rect: number[]
): { streamOn: Uint8Array; streamOff: Uint8Array } {
  const x1 = rect[0] ?? 0;
  const y1 = rect[1] ?? 0;
  const x2 = rect[2] ?? x1;
  const y2 = rect[3] ?? y1;
  const width = Math.abs(x2 - x1);
  const height = Math.abs(y2 - y1);

  // --- "On" appearance: draw an X ---
  const onCs = new PdfContentStream();
  const inset = Math.min(width, height) * 0.15;
  onCs
    .save()
    .setStrokeColor({ r: 0, g: 0, b: 0 })
    .setLineWidth(Math.max(1, Math.min(width, height) * 0.1))
    // First stroke of the X: bottom-left to top-right
    .moveTo(inset, inset)
    .lineTo(width - inset, height - inset)
    .stroke()
    // Second stroke of the X: top-left to bottom-right
    .moveTo(inset, height - inset)
    .lineTo(width - inset, inset)
    .stroke()
    .restore();

  // --- "Off" appearance: empty ---
  const offCs = new PdfContentStream();
  // Intentionally empty — nothing to draw for unchecked state

  return {
    streamOn: onCs.toUint8Array(),
    streamOff: offCs.toUint8Array()
  };
}

// =============================================================================
// Helpers
// =============================================================================

/**
 * Build the BBox array string for a form XObject appearance stream.
 * The bounding box is in the widget's coordinate space: [0 0 width height].
 */
export function buildAppearanceBBox(rect: number[]): string {
  const x1 = rect[0] ?? 0;
  const y1 = rect[1] ?? 0;
  const x2 = rect[2] ?? x1;
  const y2 = rect[3] ?? y1;
  const width = Math.abs(x2 - x1);
  const height = Math.abs(y2 - y1);
  return `[0 0 ${pdfNumber(width)} ${pdfNumber(height)}]`;
}
