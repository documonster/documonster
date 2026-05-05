/**
 * GraphicFrameXform — handles `xdr:graphicFrame` elements in drawing XML.
 *
 * A graphicFrame is the container for embedded objects such as charts.
 * The OOXML structure:
 *
 *   <xdr:graphicFrame macro="">
 *     <xdr:nvGraphicFramePr>
 *       <xdr:cNvPr id="2" name="Chart 1"/>
 *       <xdr:cNvGraphicFramePr/>
 *     </xdr:nvGraphicFramePr>
 *     <xdr:xfrm>
 *       <a:off x="0" y="0"/>
 *       <a:ext cx="0" cy="0"/>
 *     </xdr:xfrm>
 *     <a:graphic>
 *       <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
 *         <c:chart xmlns:c="..." xmlns:r="..." r:id="rId1"/>
 *       </a:graphicData>
 *     </a:graphic>
 *   </xdr:graphicFrame>
 */

import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { uuidV4 } from "@utils/uuid";

const CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const CHART_EX_URI = "http://schemas.microsoft.com/office/drawing/2014/chartex";
// URI identifying the Office `creationId` extension registered on
// `xdr:cNvPr`. This GUID is the Microsoft-assigned registry value
// for the extension — every Office-authored drawing carries the
// same uri on this element.
const CREATION_ID_EXT_URI = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}";

export interface GraphicFrameModel {
  /** Drawing object index (1-based unique id within the drawing part) */
  index?: number;
  /** Display name (e.g. "Chart 1") */
  name?: string;
  /** Relationship ID referencing the chart part */
  rId?: string;
  /** The graphic data URI — defaults to chart URI */
  graphicDataUri?: string;
  /** Raw XML for a:extLst inside xdr:cNvPr (round-trip preservation) */
  cNvPrExtLst?: string;
  /** True when the graphic frame references a cx:chart (Office 2016+ extended chart) */
  isChartEx?: boolean;
}

class GraphicFrameXform extends BaseXform {
  declare public model: GraphicFrameModel;
  declare public parser: any;

  // SAX parse state
  private _inNvPr = false;
  private _inGraphicData = false;
  /** Raw XML capture for a:extLst inside xdr:cNvPr */
  private _extLstCapture: { depth: number; parts: string[]; skipNextClose: boolean } | null = null;

  get tag(): string {
    return "xdr:graphicFrame";
  }

  prepare(model: GraphicFrameModel, options: { index: number }): void {
    if (model.index === undefined) {
      model.index = options.index + 1;
    }
  }

  render(xmlStream: any, model: GraphicFrameModel): void {
    xmlStream.openNode(this.tag, { macro: "" });

    // Non-visual properties
    xmlStream.openNode("xdr:nvGraphicFramePr");
    // Microsoft Excel starts `xdr:cNvPr/@id` at 2 (id 1 is
    // reserved for the anchor's own non-visual drawing id slot
    // internal to Office's drawing engine). Mirror the convention
    // here so freshly-authored drawings round-trip byte-for-byte
    // against Excel's output. Loaded files retain whatever id they
    // carried at parse time via `model.index`.
    const cNvPrId = model.index ?? 2;
    const cNvPrName = model.name ?? `Chart ${cNvPrId}`;
    if (model.cNvPrExtLst) {
      xmlStream.openNode("xdr:cNvPr", {
        id: cNvPrId,
        name: cNvPrName
      });
      xmlStream.writeRaw(model.cNvPrExtLst);
      xmlStream.closeNode(); // xdr:cNvPr
    } else if (model.isChartEx) {
      // ChartEx drawings — auto-generate Microsoft's standard
      // `<a:extLst>` with an `<a16:creationId>` extension so
      // Excel 2019+ can track the drawing across sessions.
      // Without this extension, strict Excel builds have been
      // observed to reject the drawing part on load with
      // "Removed Part: /xl/drawings/drawingN.xml (Drawing shape)"
      // even when the chartEx reference inside is otherwise
      // valid. A random UUID suffices — the id only has to be
      // stable within a given saved file; it isn't a content
      // address.
      xmlStream.openNode("xdr:cNvPr", {
        id: cNvPrId,
        name: cNvPrName
      });
      xmlStream.openNode("a:extLst");
      xmlStream.openNode("a:ext", { uri: CREATION_ID_EXT_URI });
      xmlStream.leafNode("a16:creationId", {
        "xmlns:a16": "http://schemas.microsoft.com/office/drawing/2014/main",
        id: `{${uuidV4().toUpperCase()}}`
      });
      xmlStream.closeNode(); // a:ext
      xmlStream.closeNode(); // a:extLst
      xmlStream.closeNode(); // xdr:cNvPr
    } else {
      xmlStream.leafNode("xdr:cNvPr", {
        id: cNvPrId,
        name: cNvPrName
      });
    }
    xmlStream.leafNode("xdr:cNvGraphicFramePr");
    xmlStream.closeNode(); // xdr:nvGraphicFramePr

    // Transform (position/size handled by anchor, so use zeros)
    xmlStream.openNode("xdr:xfrm");
    xmlStream.leafNode("a:off", { x: 0, y: 0 });
    xmlStream.leafNode("a:ext", { cx: 0, cy: 0 });
    xmlStream.closeNode(); // xdr:xfrm

    // Graphic
    xmlStream.openNode("a:graphic");
    if (model.isChartEx) {
      xmlStream.openNode("a:graphicData", { uri: CHART_EX_URI });
      xmlStream.leafNode("cx:chart", {
        "xmlns:cx": CHART_EX_URI,
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "r:id": model.rId
      });
    } else {
      xmlStream.openNode("a:graphicData", {
        uri: model.graphicDataUri ?? CHART_URI
      });
      xmlStream.leafNode("c:chart", {
        "xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "r:id": model.rId
      });
    }
    xmlStream.closeNode(); // a:graphicData
    xmlStream.closeNode(); // a:graphic

    xmlStream.closeNode(); // xdr:graphicFrame
  }

  parseOpen(node: any): boolean {
    // Raw XML capture mode for extLst
    if (this._extLstCapture) {
      const attrs = Object.entries(node.attributes || {})
        .map(([k, v]) => ` ${k}="${v}"`)
        .join("");
      if (node.isSelfClosing) {
        this._extLstCapture.parts.push(`<${node.name}${attrs}/>`);
        this._extLstCapture.skipNextClose = true;
      } else {
        this._extLstCapture.parts.push(`<${node.name}${attrs}>`);
        this._extLstCapture.depth++;
      }
      return true;
    }

    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {} as GraphicFrameModel;
        this._extLstCapture = null;
        break;
      case "xdr:cNvPr":
        this._inNvPr = true;
        this.model.index = parseInt(node.attributes.id, 10) || undefined;
        this.model.name = node.attributes.name;
        break;
      case "a:extLst":
        if (this._inNvPr) {
          // Start capturing raw XML
          const attrs = Object.entries(node.attributes || {})
            .map(([k, v]) => ` ${k}="${v}"`)
            .join("");
          this._extLstCapture = {
            depth: 1,
            parts: [`<a:extLst${attrs}>`],
            skipNextClose: false
          };
          return true;
        }
        break;
      case "a:graphicData":
        this._inGraphicData = true;
        this.model.graphicDataUri = node.attributes.uri;
        break;
      case "c:chart":
        if (this._inGraphicData) {
          this.model.rId = node.attributes["r:id"];
        }
        break;
      case "cx:chart":
        if (this._inGraphicData) {
          this.model.rId = node.attributes["r:id"];
          this.model.isChartEx = true;
        }
        break;
      default:
        break;
    }
    return true;
  }

  parseText(text: string): void {
    if (this._extLstCapture && text) {
      this._extLstCapture.parts.push(text);
    }
  }

  parseClose(name: string): boolean {
    if (this._extLstCapture) {
      if (this._extLstCapture.skipNextClose) {
        this._extLstCapture.skipNextClose = false;
        return true;
      }
      this._extLstCapture.parts.push(`</${name}>`);
      this._extLstCapture.depth--;
      if (this._extLstCapture.depth === 0) {
        this.model.cNvPrExtLst = this._extLstCapture.parts.join("");
        this._extLstCapture = null;
      }
      return true;
    }

    switch (name) {
      case this.tag:
        return false;
      case "xdr:cNvPr":
        this._inNvPr = false;
        return true;
      case "a:graphicData":
        this._inGraphicData = false;
        return true;
      default:
        return true;
    }
  }
}

export { GraphicFrameXform };
