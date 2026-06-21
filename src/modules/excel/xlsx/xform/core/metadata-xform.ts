import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { XmlSink } from "@xml/types";
import { StdDocAttributes } from "@xml/writer";

// Namespace URIs
const NS_SPREADSHEETML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_DYNAMIC_ARRAY = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray";
const XLDAPR_EXT_URI = "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}";

/**
 * Internal model for xl/metadata.xml.
 *
 * The metadata file supports various metadata types via futureMetadata blocks.
 * For dynamic array formulas, the only type we care about is "XLDAPR" (Excel
 * Dynamic Array PRoperties). The structure is:
 *
 *   metadataTypes: declares which metadata types exist (always XLDAPR for DA)
 *   futureMetadata: per-type blocks with ext data (dynamicArrayProperties)
 *   cellMetadata: maps cm indices (1-indexed on <c>) to futureMetadata entries
 *
 * For simplicity, all dynamic array formulas share a single XLDAPR futureMetadata
 * block with fDynamic="1" fCollapsed="0", and a single cellMetadata record
 * pointing to it. Every dynamic array cell gets cm="1".
 */
interface MetadataModel {
  /** Number of cells that reference dynamic array metadata */
  dynamicArrayCount: number;
}

/**
 * Parsed result from xl/metadata.xml.
 *
 * `dynamicArrayCmIndices` is a set of cm values (1-indexed, matching the `cm`
 * attribute on `<c>` elements) whose corresponding cellMetadata record points
 * to an XLDAPR metadataType. This allows precise per-cell identification of
 * dynamic array formulas without assuming all cm values are XLDAPR.
 *
 * `hasDynamicArrays` is a convenience shorthand: `dynamicArrayCmIndices.size > 0`.
 */
interface MetadataParseResult {
  hasDynamicArrays: boolean;
  dynamicArrayCmIndices: Set<number>;
}

class MetadataXform extends BaseXform {
  // Parsing state
  /** metadataType names in declaration order (1-indexed in OOXML) */
  private _metadataTypeNames: string[] = [];
  /** Per cellMetadata bk: the rc.t value (metadataType index, 1-indexed) */
  private _cellMetadataTypeRefs: number[] = [];
  /** Whether we are currently inside <cellMetadata> */
  private _inCellMetadata = false;
  /** Current rc.t value being collected inside a <bk> */
  private _currentRcType: number | undefined = undefined;

  get tag(): string {
    return "metadata";
  }

  /**
   * Render xl/metadata.xml for the given model.
   * Only emits content when dynamicArrayCount > 0.
   */
  render(xmlStream: XmlSink, model: MetadataModel): void {
    if (!model || model.dynamicArrayCount <= 0) {
      return;
    }

    xmlStream.openXml(StdDocAttributes);

    xmlStream.openNode("metadata", {
      xmlns: NS_SPREADSHEETML,
      "xmlns:xda": NS_DYNAMIC_ARRAY
    });

    // metadataTypes: declare XLDAPR
    xmlStream.openNode("metadataTypes", { count: "1" });
    xmlStream.leafNode("metadataType", {
      name: "XLDAPR",
      minSupportedVersion: "120000",
      copy: "1",
      pasteAll: "1",
      pasteValues: "1",
      merge: "1",
      splitFirst: "1",
      rowColShift: "1",
      clearFormats: "1",
      clearComments: "1",
      assign: "1",
      coerce: "1",
      adjust: "1",
      cellMeta: "1"
    });
    xmlStream.closeNode(); // </metadataTypes>

    // futureMetadata: one shared block for all DA formulas
    xmlStream.openNode("futureMetadata", { name: "XLDAPR", count: "1" });
    xmlStream.openNode("bk");
    xmlStream.openNode("extLst");
    xmlStream.openNode("ext", { uri: XLDAPR_EXT_URI });
    xmlStream.leafNode("xda:dynamicArrayProperties", {
      fDynamic: "1",
      fCollapsed: "0"
    });
    xmlStream.closeNode(); // </ext>
    xmlStream.closeNode(); // </extLst>
    xmlStream.closeNode(); // </bk>
    xmlStream.closeNode(); // </futureMetadata>

    // cellMetadata: one record (all DA cells share cm="1" which points here)
    xmlStream.openNode("cellMetadata", { count: "1" });
    xmlStream.openNode("bk");
    // t="1" → metadataType index (1-indexed), v="0" → futureMetadata block index (0-indexed)
    xmlStream.leafNode("rc", { t: "1", v: "0" });
    xmlStream.closeNode(); // </bk>
    xmlStream.closeNode(); // </cellMetadata>

    xmlStream.closeNode(); // </metadata>
  }

  // =========================================================================
  // Parse: build precise cm → metadataType mapping
  // =========================================================================

  reset(): void {
    this._metadataTypeNames = [];
    this._cellMetadataTypeRefs = [];
    this._inCellMetadata = false;
    this._currentRcType = undefined;
  }

  parseOpen(node: any): boolean {
    switch (node.name) {
      case "metadata":
        this.reset();
        return true;

      case "metadataType":
        // Collect type names in declaration order (1-indexed in OOXML,
        // but stored 0-indexed here — we adjust when looking up).
        this._metadataTypeNames.push(node.attributes.name || "");
        return true;

      case "cellMetadata":
        this._inCellMetadata = true;
        return true;

      case "bk":
        if (this._inCellMetadata) {
          this._currentRcType = undefined;
        }
        return true;

      case "rc":
        if (this._inCellMetadata && node.attributes.t !== undefined) {
          this._currentRcType = parseInt(node.attributes.t, 10);
        }
        return true;

      // Parse nested elements without action
      case "metadataTypes":
      case "futureMetadata":
      case "extLst":
      case "ext":
      case "xda:dynamicArrayProperties":
        return true;

      default:
        return false;
    }
  }

  parseText(): void {}

  parseClose(name: string): boolean {
    switch (name) {
      case "bk":
        if (this._inCellMetadata) {
          // Record the metadataType reference for this cellMetadata entry.
          // rc.t is 1-indexed into metadataTypes; store as-is.
          this._cellMetadataTypeRefs.push(this._currentRcType ?? 0);
          this._currentRcType = undefined;
        }
        return true;

      case "cellMetadata":
        this._inCellMetadata = false;
        return true;

      case "metadata": {
        // Build the set of cm values (1-indexed) that map to XLDAPR.
        // cm on <c> is 1-indexed into cellMetadata bk entries.
        const dynamicArrayCmIndices = new Set<number>();
        for (let i = 0; i < this._cellMetadataTypeRefs.length; i++) {
          const typeIndex = this._cellMetadataTypeRefs[i]; // 1-indexed
          // metadataTypeNames is 0-indexed, so subtract 1
          const typeName = this._metadataTypeNames[typeIndex - 1];
          if (typeName === "XLDAPR") {
            // cm is 1-indexed: bk[0] → cm=1, bk[1] → cm=2, etc.
            dynamicArrayCmIndices.add(i + 1);
          }
        }

        this.model = {
          hasDynamicArrays: dynamicArrayCmIndices.size > 0,
          dynamicArrayCmIndices
        } satisfies MetadataParseResult;
        return false; // done parsing
      }

      default:
        return true;
    }
  }
}

export { MetadataXform };
export type { MetadataModel, MetadataParseResult };
