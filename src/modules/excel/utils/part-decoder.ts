import { textDecoder } from "@utils/binary";
import { createXmlPartLeadTrimmer, trimXmlPartLead } from "@utils/xml-part";
import { XmlParseError } from "@xml/errors";

/** The one method of `TextDecoder` an XML part reader uses. */
export interface PartDecoder {
  decode(input?: Uint8Array, options?: { stream?: boolean }): string;
}

/** The text of an XML part's bytes, decoded as this reader decodes a part it holds whole, from its first `<`. */
export function partText(bytes: Uint8Array): string {
  return trimXmlPartLead(textDecoder.decode(bytes));
}

/**
 * A strict UTF-8 decoder for the bytes of an XML part, whose failure is reported as the malformed part it is.
 *
 * Strict because a part that is not UTF-8 is not the document its bytes claim to be, and replacing what does not
 * decode would hand the parser a different one. But `TextDecoder` reports that as a bare `TypeError`, which a caller
 * cannot tell apart from a bug in this library — the readers that parse a part as it streams in were the one route by
 * which a damaged package surfaced as one.
 *
 * The white space a part may begin with is dropped as it arrives, however it is split: see `trimXmlPartLead`.
 */
export function createPartDecoder(): PartDecoder {
  const decoder = new TextDecoder("utf-8", { fatal: true });
  const trimLead = createXmlPartLeadTrimmer();
  return {
    decode(input, options) {
      let text: string;
      try {
        text = decoder.decode(input, options);
      } catch (error) {
        throw new XmlParseError("the part is not valid UTF-8", undefined, { cause: error });
      }
      return trimLead(text);
    }
  };
}
