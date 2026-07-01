/**
 * Integration test: Office 365 threaded comments round-trip.
 *
 * Exercises the parts documonster must wire up end-to-end when a workbook
 * carries modern threaded comments:
 *
 *   - `xl/persons/person.xml` (workbook-level person directory)
 *   - `xl/threadedComments/threadedComment{N}.xml` (per-sheet part)
 *   - Sheet-level rel pointing at the threaded comments part
 *   - Workbook-level rel pointing at the person list
 *   - `[Content_Types].xml` overrides for both new content types
 *
 * These parts exist entirely separately from classic VML comments;
 * the two flavours can coexist on the same cell. This test covers
 * creation, save, reload, and re-save to verify no information is
 * lost across rounds.
 */

import { extractAll } from "@archive/unzip/extract";
import { expectValidXlsx } from "@excel/__tests__/helpers/expect-valid-xlsx";
import { getPersons } from "@excel/core/workbook";
import { Cell, Workbook } from "@excel/index";
import type { ThreadedComment } from "@excel/types";
import { describe, it, expect } from "vitest";

const decoder = new TextDecoder();

describe("threaded comments round-trip", () => {
  it("emits persons.xml + threadedComment part + rels + content types on save", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", "Discuss this");

    // Register two commenters.
    const aliceId = Workbook.registerPerson(wb, "Alice", "alice@example.com", "AD");
    const bobId = Workbook.registerPerson(wb, "Bob", "bob@example.com", "AD");

    const top: ThreadedComment = {
      id: "{11111111-1111-1111-1111-111111111111}",
      personId: aliceId,
      date: "2024-01-01T10:00:00Z",
      text: "Why is this number so low?"
    };
    const reply: ThreadedComment = {
      id: "{22222222-2222-2222-2222-222222222222}",
      parentId: top.id,
      personId: bobId,
      date: "2024-01-01T10:05:00Z",
      text: "Methodology changed — see the deck.",
      done: true
    };
    ws.threadedComments.push({ ref: "A1", comment: top });
    ws.threadedComments.push({ ref: "A1", comment: reply });

    const buf = await Workbook.toBuffer(wb);
    await expectValidXlsx(buf, { label: "threaded-comments emit" });
    const entries = await extractAll(new Uint8Array(buf));

    // Parts landed at the expected paths.
    expect(entries.get("xl/persons/person.xml")).toBeDefined();
    expect(entries.get("xl/threadedComments/threadedComment1.xml")).toBeDefined();

    // Content Types override list mentions both.
    const contentTypes = decoder.decode(entries.get("[Content_Types].xml")!.data);
    expect(contentTypes).toContain("application/vnd.ms-excel.person+xml");
    expect(contentTypes).toContain("application/vnd.ms-excel.threadedcomments+xml");

    // Workbook rels references the persons part.
    const wbRels = decoder.decode(entries.get("xl/_rels/workbook.xml.rels")!.data);
    expect(wbRels).toContain("persons/person.xml");
    expect(wbRels).toContain("/relationships/person");

    // Sheet rels references the threaded comments part.
    const sheetRels = decoder.decode(entries.get("xl/worksheets/_rels/sheet1.xml.rels")!.data);
    expect(sheetRels).toContain("../threadedComments/threadedComment1.xml");
    expect(sheetRels).toContain("/relationships/threadedComment");

    // Threaded comments XML has the expected structure.
    const tcXml = decoder.decode(entries.get("xl/threadedComments/threadedComment1.xml")!.data);
    expect(tcXml).toContain(`id="${top.id}"`);
    expect(tcXml).toContain(`parentId="${top.id}"`);
    expect(tcXml).toContain(`ref="A1"`);
    expect(tcXml).toContain(`personId="${aliceId}"`);
    expect(tcXml).toContain(`personId="${bobId}"`);
    expect(tcXml).toContain("Why is this number so low?");
    expect(tcXml).toContain('done="1"');

    // Person list XML has both people with their provider info.
    const personsXml = decoder.decode(entries.get("xl/persons/person.xml")!.data);
    expect(personsXml).toContain(`displayName="Alice"`);
    expect(personsXml).toContain(`userId="alice@example.com"`);
    expect(personsXml).toContain(`providerId="AD"`);
    expect(personsXml).toContain(`displayName="Bob"`);
  });

  it("round-trips a workbook without corrupting threaded comments", async () => {
    // Load an xlsx we authored, save it again, and verify the parts
    // are identical across the second round.
    const wb1 = Workbook.create();
    const ws = Workbook.addWorksheet(wb1, "Sheet1");
    Cell.setValue(ws, "A1", 42);
    const personId = Workbook.registerPerson(wb1, "Loader", "loader@example.com");
    ws.threadedComments.push({
      ref: "A1",
      comment: { personId, text: "Inspected", date: "2024-06-01T00:00:00Z" }
    });

    const firstBuf = await Workbook.toBuffer(wb1);
    await expectValidXlsx(firstBuf, { label: "threaded-comments roundtrip 1" });
    const wb2 = Workbook.create();
    await Workbook.read(wb2, firstBuf);

    // After reload the workbook carries the same person directory and
    // the worksheet carries the same threaded comments.
    expect(getPersons(wb2)).toHaveLength(1);
    expect(getPersons(wb2)[0].displayName).toBe("Loader");
    expect(getPersons(wb2)[0].userId).toBe("loader@example.com");

    const reloadedSheet = Workbook.getWorksheet(wb2, "Sheet1")!;
    expect(reloadedSheet.threadedComments).toHaveLength(1);
    expect(reloadedSheet.threadedComments[0].ref).toBe("A1");
    expect(reloadedSheet.threadedComments[0].comment.text).toBe("Inspected");
    expect(reloadedSheet.threadedComments[0].comment.personId).toBe(personId);

    // A second save preserves the parts byte-compatibly enough that
    // Excel would accept them — we assert on structured content, not
    // byte equality, because element attribute order isn't strictly
    // controlled by the writer.
    const secondBuf = await Workbook.toBuffer(wb2);
    await expectValidXlsx(secondBuf, { label: "threaded-comments roundtrip 2" });
    const secondEntries = await extractAll(new Uint8Array(secondBuf));
    expect(secondEntries.get("xl/persons/person.xml")).toBeDefined();
    expect(secondEntries.get("xl/threadedComments/threadedComment1.xml")).toBeDefined();
    const secondTc = decoder.decode(
      secondEntries.get("xl/threadedComments/threadedComment1.xml")!.data
    );
    expect(secondTc).toContain("Inspected");
    expect(secondTc).toContain(`personId="${personId}"`);
  });

  it("supports @mentions with startIndex/length", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    const alice = Workbook.registerPerson(wb, "Alice");
    const bob = Workbook.registerPerson(wb, "Bob");
    ws.threadedComments.push({
      ref: "B2",
      comment: {
        personId: alice,
        text: "@Bob please review",
        mentions: [
          {
            mentionPersonId: bob,
            startIndex: 0,
            length: 4
          }
        ]
      }
    });
    const buf = await Workbook.toBuffer(wb);
    await expectValidXlsx(buf, { label: "threaded-comments mentions" });
    const entries = await extractAll(new Uint8Array(buf));
    const xml = decoder.decode(entries.get("xl/threadedComments/threadedComment1.xml")!.data);
    expect(xml).toContain("<mentions>");
    expect(xml).toContain(`mentionpersonId="${bob}"`);
    expect(xml).toContain(`startIndex="0"`);
    expect(xml).toContain(`length="4"`);

    const wb2 = Workbook.create();
    await Workbook.read(wb2, buf);
    const reloaded = Workbook.getWorksheet(wb2, "Sheet1")!.threadedComments[0].comment;
    expect(reloaded.mentions).toHaveLength(1);
    expect(reloaded.mentions![0].startIndex).toBe(0);
    expect(reloaded.mentions![0].length).toBe(4);
  });

  it("registerPerson deduplicates identical (displayName, userId) pairs", () => {
    const wb = Workbook.create();
    const a = Workbook.registerPerson(wb, "Alice", "alice@example.com");
    const b = Workbook.registerPerson(wb, "Alice", "alice@example.com");
    expect(a).toBe(b);
    expect(getPersons(wb)).toHaveLength(1);

    const c = Workbook.registerPerson(wb, "Alice", "alice@other.com");
    expect(c).not.toBe(a);
    expect(getPersons(wb)).toHaveLength(2);
  });

  it("does not emit persons.xml when no threaded comments are used", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    const buf = await Workbook.toBuffer(wb);
    await expectValidXlsx(buf, { label: "no-threaded-comments" });
    const entries = await extractAll(new Uint8Array(buf));
    expect(entries.get("xl/persons/person.xml")).toBeUndefined();
    expect(entries.get("xl/threadedComments/threadedComment1.xml")).toBeUndefined();

    // Content Types must not mention the extra URIs either.
    const contentTypes = decoder.decode(entries.get("[Content_Types].xml")!.data);
    expect(contentTypes).not.toContain("application/vnd.ms-excel.person+xml");
    expect(contentTypes).not.toContain("application/vnd.ms-excel.threadedcomments+xml");
  });
});
