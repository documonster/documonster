/**
 * The text of an XML part as a document reader parses it.
 */

/** A byte order mark and the white space after it, at the start of a part. */
const LEAD = /^\uFEFF?[ \t\r\n]*/;
const SPACE = /^[ \t\r\n]*/;

/**
 * `text` without the byte order mark and white space it may begin with.
 *
 * Some producers write a package part with white space — a line break, say — before its XML
 * declaration, where XML allows none, and the XML parser refuses it. Office opens such a file, and so
 * do the document readers: they parse a part from its first `<`. White space before the root element
 * where there is no declaration is allowed and means nothing, so dropping it changes no part that was
 * well-formed. The XML parser itself stays strict; this is the readers' policy for the parts of a
 * package, not for XML a caller hands them.
 */
export function trimXmlPartLead(text: string): string {
  const lead = LEAD.exec(text)![0].length;
  return lead === 0 ? text : text.slice(lead);
}

/**
 * {@link trimXmlPartLead} for a part that arrives in pieces: each call takes the next piece of its
 * text and returns it without what is still the part's lead, however the lead is split.
 */
export function createXmlPartLeadTrimmer(): (text: string) => string {
  let atStart = true;
  let first = true;
  return text => {
    if (!atStart || text.length === 0) {
      return text;
    }
    // A byte order mark is the part's first character only: one after white space is not lead.
    const lead = (first ? LEAD : SPACE).exec(text)![0].length;
    first = false;
    const trimmed = lead === 0 ? text : text.slice(lead);
    // A piece that was all lead leaves the next one at the start still.
    atStart = trimmed.length === 0;
    return trimmed;
  };
}
