/**
 * `Image` namespace surface.
 *
 * `Image.add(wb, payload)` registers media in the workbook → imageId;
 * `Image.place(ws, imageId, range)` anchors it on a sheet;
 * `Image.setBackground(ws, imageId)` sets a sheet background;
 * `Image.list(ws)` lists a sheet's images.
 */
export { addWorkbookImage as add } from "@excel/workbook-core";
export {
  addImage as place,
  getImages as list,
  addBackgroundImage as setBackground,
  getBackgroundImageId as getBackground,
  addShape,
  getShapes
} from "@excel/worksheet";
