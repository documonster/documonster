/**
 * `Image` namespace surface.
 *
 * `Image.add(wb, payload)` registers media in the workbook → imageId;
 * `Image.place(ws, imageId, range)` anchors it on a sheet;
 * `Image.setBackground(ws, imageId)` sets a sheet background, `Image.removeBackground(ws)` removes it;
 * `Image.list(ws)` lists a sheet's images, `Image.remove(ws, image)` removes one,
 * `Image.listShapes(ws)` lists its shapes.
 *
 * Two collections share the namespace: the workbook's media (`add` registers, by id) and a
 * sheet's placements of it (`place`, `list`, `remove`). Removing a placement leaves the media
 * registered, so its id can be placed again.
 */
import { checkWorksheetImage } from "@excel/core/feature-handles";
import type { ImageModel, ImageModelInput, WorksheetImage } from "@excel/core/image";
import { applyImageModel, imageClone, imageCreate, imageModel } from "@excel/core/image";
import type { WorkbookData } from "@excel/core/workbook-core";
import { addWorkbookImage, checkWorkbook, getImage } from "@excel/core/workbook-core";
import {
  addBackgroundImage,
  addImage,
  addShape as addWorksheetShape,
  getBackgroundImageId,
  getImages,
  getShapes as listWorksheetShapes,
  removeBackgroundImage,
  removeImage,
  removeShape as removeWorksheetShape,
  setImageModel
} from "@excel/core/worksheet";
import type { WorksheetData } from "@excel/core/worksheet-core";
import {
  checkWorksheet,
  checkWorksheetAttached,
  getSheetWorkbook
} from "@excel/core/worksheet-core";
import { ExcelImageError } from "@excel/errors";
import { checkListTarget } from "@excel/surface/argument-checks";
import type { AddImageRange, AddShapeOptions, ImageData, ShapeModel } from "@excel/types";
import { argumentChecks } from "@excel/utils/arguments";
import { checkDrawingRange, checkShapeOptions } from "@excel/utils/drawing-checks";

/**
 * Register a picture in the workbook and return its id, for {@link place} and {@link setBackground}.
 *
 * @throws {ExcelImageError} when the image has no extension of the supported kinds (`png`, `jpeg`, `jpg`, `gif`), or no
 * bytes — `buffer`, `base64` or `filename` — and no `link`.
 */
export function add(wb: WorkbookData, image: ImageData): number {
  checkWorkbook(wb);
  return addWorkbookImage(wb, image);
}

function checkImageId(where: string, ws: WorksheetData, imageId: string | number): void {
  if (getImage(getSheetWorkbook(ws), imageId) === undefined) {
    throw new ExcelImageError(
      `${where}: the workbook has no image ${JSON.stringify(imageId)}; add it with Image.add first.`
    );
  }
}

/**
 * Anchor the workbook image `imageId` (from {@link add}) on the sheet, and return the placed image — the handle
 * {@link list} lists, which {@link remove} and {@link setModel} take. A lower-case range (`"b2:c3"`) is read as Excel
 * reads it.
 *
 * @throws {ExcelImageError} when the workbook has no such image — which used to be accepted and written as a picture with
 * nothing to show.
 */
export function place(
  ws: WorksheetData,
  imageId: string | number,
  range: AddImageRange
): WorksheetImage {
  checkWorksheet(ws);
  checkImageId("Image.place", ws, imageId);
  // A range that is none was accepted and failed only at write ("Image has no range", or `NaN` in the drawing); a
  // negative extent or a bottom-right corner above the top-left one was written as given.
  checkDrawingRange("Image.place", "range", range);
  return addImage(ws, imageId, range);
}

/** Use the workbook image `imageId` as the sheet's background. @throws {ExcelImageError} when there is no such image. */
export function setBackground(ws: WorksheetData, imageId: string | number): void {
  checkWorksheet(ws);
  checkImageId("Image.setBackground", ws, imageId);
  addBackgroundImage(ws, imageId);
}

/** The sheet's placed images, in the order they were placed: the handles {@link remove} and {@link setModel} take. */
export function list(ws: WorksheetData): WorksheetImage[] {
  checkWorksheet(ws);
  return getImages(ws);
}

/**
 * The workbook image id of the sheet's background picture, or `undefined` when it has none — as the string the sheet
 * stores it as (`"0"`); `Number(id)` is the id {@link add} returned.
 */
export function getBackground(ws: WorksheetData): string | undefined {
  checkWorksheet(ws);
  return getBackgroundImageId(ws);
}

/** Remove the sheet's background picture. Returns `true` if it had one. */
export function removeBackground(ws: WorksheetData): boolean {
  checkWorksheet(ws);
  return removeBackgroundImage(ws);
}

/**
 * Remove a placed image — the handle {@link list} returned, or its index in that list. Returns `true` if one was
 * removed. The picture stays registered in the workbook, so its id can still be placed elsewhere.
 */
export function remove(ws: WorksheetData, image: WorksheetImage | number): boolean {
  checkWorksheet(ws);
  checkListTarget("Image.remove", "image", image, "a placed image");
  return removeImage(ws, image);
}

/**
 * Draw a shape on the sheet over `options.range`, and return the shape as {@link listShapes} lists it.
 *
 * @throws {ExcelArgumentError} when the geometry is not a DrawingML preset (`"rect"`, `"ellipse"`, …), a colour is
 * not hex, the line width is outside 0–1584 points, or the range covers no area.
 */
export function addShape(ws: WorksheetData, options: AddShapeOptions): ShapeModel {
  checkWorksheet(ws);
  checkShapeOptions("Image.addShape", options);
  return addWorksheetShape(ws, options);
}

/** Remove a shape — one {@link listShapes} listed, or its index in that list. Returns `true` if one was removed. */
export function removeShape(ws: WorksheetData, shape: ShapeModel | number): boolean {
  checkWorksheet(ws);
  checkListTarget("Image.removeShape", "shape", shape, "a shape");
  return removeWorksheetShape(ws, shape);
}

/** The sheet's shapes, in the order they were drawn. */
export function listShapes(ws: WorksheetData): ShapeModel[] {
  checkWorksheet(ws);
  return listWorksheetShapes(ws);
}

/** A copy of the image's model: what it shows, where, and how. */
export function getModel(image: WorksheetImage): ImageModel {
  checkWorksheetImage(image);
  return imageModel(image);
}

const IMAGE_TYPES = ["image", "background", "watermark", "headerImage"];

/**
 * Change a placed image to `model`. The model is checked first — its type, that the workbook has its image, and for
 * a picture, its range as {@link place} checks one — so a refused model leaves the image as it was.
 *
 * @throws {ExcelArgumentError} when the image is no longer on its sheet, or the model is malformed.
 * @throws {ExcelImageError} when the workbook has no image `model.imageId`: the picture used to be dropped from the file.
 */
export function setModel(image: WorksheetImage, model: ImageModelInput): void {
  checkWorksheetImage(image);
  checkWorksheetAttached(image.worksheet);
  const check = argumentChecks("Image.setModel");
  check.required("model", model);
  if (check.record("model", model)) {
    check.required("model.type", model.type);
    check.oneOf("model.type", model.type, IMAGE_TYPES);
    checkImageId("Image.setModel", image.worksheet, model.imageId);
    if (model.type === "image") {
      checkDrawingRange("Image.setModel", "model.range", model.range);
    }
  }
  if (!image.worksheet._media.includes(image)) {
    check.fail("image", "an image on its sheet (from Image.list)", image);
  }
  setImageModel(image, model);
}

/** An image handle. */
export type { WorksheetImage as Handle } from "@excel/core/image";

// --- deprecated names ---

/**
 * @deprecated Use {@link place}. This builds an image record bound to `worksheet` without placing it, so it is not
 * listed, not written, and nothing public attaches it. Removed in 1.0.
 */
export function create(worksheet: WorksheetData, model?: ImageModelInput): WorksheetImage {
  checkWorksheet(worksheet);
  return imageCreate(worksheet, model);
}
/**
 * @deprecated Use {@link place} with the image's id and range ({@link getModel} gives both). This copies the record
 * without placing the copy, so it is not listed and not written. Removed in 1.0.
 */
export function clone(image: WorksheetImage, worksheet?: WorksheetData): WorksheetImage {
  checkWorksheetImage(image);
  return imageClone(image, worksheet);
}

/** @deprecated Use {@link getModel}. Removed in 1.0. */
export const model = getModel;
/**
 * @deprecated Use {@link setModel}. Removed in 1.0. Applies the model only, as it
 * always did; `setModel` also records that a picture the image stops showing is no longer used.
 */
export function applyModel(image: WorksheetImage, value: ImageModelInput): void {
  checkWorksheetImage(image);
  checkWorksheetAttached(image.worksheet);
  applyImageModel(image, value);
}
/** @deprecated Use {@link listShapes}. Removed in 1.0. */
export const getShapes = listShapes;
