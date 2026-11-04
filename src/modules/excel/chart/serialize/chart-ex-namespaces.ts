/**
 * The declarations the ChartEx writer's `<cx:chartSpace>` carries, by prefix — shared by the writer,
 * which emits them, and the reader, whose preserved subtrees need not repeat them.
 */
export const CHART_EX_ROOT_NAMESPACES: Readonly<Record<string, string>> = {
  cx: "http://schemas.microsoft.com/office/drawing/2014/chartex",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main"
};
