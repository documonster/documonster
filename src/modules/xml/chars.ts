/**
 * XML 1.0 character classes (§2.2, §2.3), shared by the SAX parser and the DOCTYPE reader so
 * both read a name the same way.
 */

export const NAME_START_CHAR_ASCII = /* @__PURE__ */ (() => {
  const t = new Uint8Array(128);
  for (let i = 0x61; i <= 0x7a; i++) {
    t[i] = 1;
  } // a-z
  for (let i = 0x41; i <= 0x5a; i++) {
    t[i] = 1;
  } // A-Z
  t[0x5f] = 1; // _
  t[0x3a] = 1; // :
  return t;
})();

export const NAME_CHAR_ASCII = /* @__PURE__ */ (() => {
  const t = new Uint8Array(128);
  for (let i = 0x61; i <= 0x7a; i++) {
    t[i] = 1;
  } // a-z
  for (let i = 0x41; i <= 0x5a; i++) {
    t[i] = 1;
  } // A-Z
  for (let i = 0x30; i <= 0x39; i++) {
    t[i] = 1;
  } // 0-9
  t[0x5f] = 1; // _
  t[0x3a] = 1; // :
  t[0x2d] = 1; // -
  t[0x2e] = 1; // .
  return t;
})();

/** XML white space: space, tab, line feed, carriage return. */
export function isS(c: number): boolean {
  return c === 0x20 || c === 0xa || c === 0xd || c === 0x9;
}

export function isNameStartChar(c: number): boolean {
  if (c < 128) {
    return NAME_START_CHAR_ASCII[c] === 1;
  }
  return (
    (c >= 0xc0 && c <= 0xd6) ||
    (c >= 0xd8 && c <= 0xf6) ||
    (c >= 0xf8 && c <= 0x2ff) ||
    (c >= 0x370 && c <= 0x37d) ||
    (c >= 0x37f && c <= 0x1fff) ||
    c === 0x200c ||
    c === 0x200d ||
    (c >= 0x2070 && c <= 0x218f) ||
    (c >= 0x2c00 && c <= 0x2fef) ||
    (c >= 0x3001 && c <= 0xd7ff) ||
    (c >= 0xf900 && c <= 0xfdcf) ||
    (c >= 0xfdf0 && c <= 0xfffd) ||
    (c >= 0x10000 && c <= 0xeffff)
  );
}

export function isNameChar(c: number): boolean {
  if (c < 128) {
    return NAME_CHAR_ASCII[c] === 1;
  }
  return (
    c === 0xb7 ||
    (c >= 0xc0 && c <= 0xd6) ||
    (c >= 0xd8 && c <= 0xf6) ||
    (c >= 0xf8 && c <= 0x2ff) ||
    (c >= 0x300 && c <= 0x36f) ||
    (c >= 0x370 && c <= 0x37d) ||
    (c >= 0x37f && c <= 0x1fff) ||
    c === 0x200c ||
    c === 0x200d ||
    (c >= 0x203f && c <= 0x2040) ||
    (c >= 0x2070 && c <= 0x218f) ||
    (c >= 0x2c00 && c <= 0x2fef) ||
    (c >= 0x3001 && c <= 0xd7ff) ||
    (c >= 0xf900 && c <= 0xfdcf) ||
    (c >= 0xfdf0 && c <= 0xfffd) ||
    (c >= 0x10000 && c <= 0xeffff)
  );
}

/** XML 1.0 §2.2 `Char`. */
export function isChar10(c: number): boolean {
  return (
    (c >= 0x20 && c <= 0xd7ff) ||
    c === 0xa ||
    c === 0xd ||
    c === 0x9 ||
    (c >= 0xe000 && c <= 0xfffd) ||
    (c >= 0x10000 && c <= 0x10ffff)
  );
}

/** Whether every character of `text` is an XML 1.0 `Char` — no control character, lone surrogate or U+FFFE. */
export function isXmlText(text: string): boolean {
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    if (c >= 0x20 && c < 0xd800) {
      continue;
    }
    if (c >= 0xd800 && c <= 0xdbff) {
      const next = text.charCodeAt(i + 1);
      if (!(next >= 0xdc00 && next <= 0xdfff)) {
        return false;
      }
      i++;
    } else if (!isChar10(c)) {
      return false;
    }
  }
  return true;
}

/** An XML `Name` (§2.3): a NameStartChar, then NameChars. */
export function isXmlName(text: string): boolean {
  let first = true;
  for (const ch of text) {
    const c = ch.codePointAt(0)!;
    if (first ? !isNameStartChar(c) : !isNameChar(c)) {
      return false;
    }
    first = false;
  }
  return !first;
}

/**
 * Whether `ref` holds only digits from `start` on, and at least one: hex digits when `hex`.
 * `parseInt` stops at the first character it cannot read, so `&#65abc;` read as `A` without this.
 */
function isCharRefDigits(ref: string, start: number, hex: boolean): boolean {
  if (ref.length <= start) {
    return false;
  }
  for (let i = start; i < ref.length; i++) {
    const c = ref.charCodeAt(i);
    if (
      !(c >= 0x30 && c <= 0x39) &&
      !(hex && ((c >= 0x61 && c <= 0x66) || (c >= 0x41 && c <= 0x46)))
    ) {
      return false;
    }
  }
  return true;
}

/**
 * The character a character reference names (`#65`, `#x41` — the text between `&` and `;`), or
 * `undefined` when it is malformed or names a character XML cannot hold. `&#X41;` is not XML's
 * spelling, only `x` is, but has always been read.
 */
export function charRefValue(ref: string): string | undefined {
  const hex = ref.charCodeAt(1) === 0x78 || ref.charCodeAt(1) === 0x58;
  if (!isCharRefDigits(ref, hex ? 2 : 1, hex)) {
    return undefined;
  }
  const code = hex ? parseInt(ref.slice(2), 16) : parseInt(ref.slice(1), 10);
  return isChar10(code) ? String.fromCodePoint(code) : undefined;
}
