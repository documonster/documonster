/**
 * The vocabularies the worksheet settings are checked against are the schema's, held to the ECMA-376 schemas
 * themselves rather than to a second hand-written list. A value the schema allows but a list here lacks is kept as
 * read, so nothing is lost — but its field reads as absent, and that is a bug this test exists to catch.
 */
import {
  CELL_COMMENTS,
  ORIENTATIONS,
  PAGE_ORDERS,
  PANES,
  PRINT_ERRORS,
  SHEET_STATES,
  VIEW_STATES,
  VIEW_STYLES
} from "@excel/utils/sheet-settings-checks";
import { OOXML_ENUMERATIONS } from "@test/fixtures/ooxml-enumerations";
import { describe, expect, it } from "vitest";

const { sml } = OOXML_ENUMERATIONS;

const sorted = (values: readonly string[]) => [...values].sort();

describe("the worksheet settings take the schema's vocabularies", () => {
  it.each([
    ["state", SHEET_STATES, sml.ST_SheetState],
    ["pageSetup.orientation", ORIENTATIONS, sml.ST_Orientation],
    ["pageSetup.pageOrder", PAGE_ORDERS, sml.ST_PageOrder],
    ["pageSetup.errors", PRINT_ERRORS, sml.ST_PrintError],
    ["views[].style", VIEW_STYLES, sml.ST_SheetViewType],
    ["views[].activePane", PANES, sml.ST_Pane]
  ] as const)("%s", (_name, vocabulary, schema) => {
    expect(sorted(vocabulary)).toEqual(sorted(schema));
  });

  it("pageSetup.cellComments, beside the model's default and the spelling this library once wrote", () => {
    expect(sorted(CELL_COMMENTS)).toEqual(sorted([...sml.ST_CellComments, "None", "asDisplyed"]));
  });

  it("views[].state, which reads ST_PaneState's frozenSplit as frozen and names a sheet with no pane", () => {
    expect(sorted(VIEW_STATES)).toEqual(
      sorted([...sml.ST_PaneState.filter(state => state !== "frozenSplit"), "normal"])
    );
  });
});
