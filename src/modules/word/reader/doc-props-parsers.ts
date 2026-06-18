/**
 * DOCX Reader - Document Properties Parsers
 *
 * Parses standard OPC document properties XML files:
 * - docProps/core.xml (CoreProperties — title, author, dc:* fields)
 * - docProps/app.xml (AppProperties — application info)
 * - docProps/custom.xml (CustomProperty[] — user-defined name/value pairs)
 * - word/fontTable.xml (FontDef[] — embedded font references)
 */

import { type Mutable } from "@word/core/internal-utils";
import { attrVal, findChildNs, findChildrenNs } from "@word/reader/parse-utils";
import type {
  AppProperties,
  CoreProperties,
  CustomProperty,
  CustomPropertyValue,
  FontDef
} from "@word/types";
import { findChild, parseXml, textContent } from "@xml/dom";

// =============================================================================
// Core Properties Parser
// =============================================================================

function parseCoreProps(xmlStr: string): CoreProperties {
  const doc = parseXml(xmlStr);
  const root = doc.root;
  const props: Mutable<CoreProperties> & Record<string, unknown> = {};

  const fields: [string, string][] = [
    ["dc:title", "title"],
    ["dc:subject", "subject"],
    ["dc:creator", "creator"],
    ["dc:description", "description"],
    ["cp:keywords", "keywords"],
    ["cp:lastModifiedBy", "lastModifiedBy"],
    ["cp:revision", "revision"],
    ["cp:category", "category"]
  ];

  for (const [tag, prop] of fields) {
    const el = findChild(root, tag);
    if (el) {
      const val = textContent(el);
      if (val) {
        props[prop] = val;
      }
    }
  }

  const createdEl = findChild(root, "dcterms:created");
  if (createdEl) {
    const val = textContent(createdEl);
    if (val) {
      props.created = new Date(val);
    }
  }

  const modifiedEl = findChild(root, "dcterms:modified");
  if (modifiedEl) {
    const val = textContent(modifiedEl);
    if (val) {
      props.modified = new Date(val);
    }
  }

  return props;
}

// =============================================================================
// App Properties Parser
// =============================================================================

function parseAppProps(xmlStr: string): AppProperties {
  const doc = parseXml(xmlStr);
  const root = doc.root;
  const props: Mutable<AppProperties> & Record<string, unknown> = {};

  const strFields = ["Application", "AppVersion", "Company", "Manager"];
  const intFields = ["Pages", "Words", "Characters", "Lines", "Paragraphs"];

  for (const field of strFields) {
    const el = findChild(root, field);
    if (el) {
      const val = textContent(el);
      if (val) {
        props[field.charAt(0).toLowerCase() + field.slice(1)] = val;
      }
    }
  }

  for (const field of intFields) {
    const el = findChild(root, field);
    if (el) {
      const val = textContent(el);
      if (val) {
        props[field.charAt(0).toLowerCase() + field.slice(1)] = parseInt(val, 10);
      }
    }
  }

  return props;
}

// =============================================================================
// Custom Properties Parser
// =============================================================================

function parseCustomPropsXml(xmlStr: string): CustomProperty[] {
  const doc = parseXml(xmlStr);
  const root = doc.root;
  const props: CustomProperty[] = [];

  for (const propEl of root.children) {
    if (propEl.type !== "element" || propEl.name !== "property") {
      continue;
    }
    const name = propEl.attributes["name"];
    if (!name) {
      continue;
    }

    let value: CustomPropertyValue | undefined;
    for (const child of propEl.children) {
      if (child.type !== "element") {
        continue;
      }
      const tn = child.name;
      const tv = textContent(child);
      if (tn === "vt:lpwstr") {
        value = { type: "string", value: tv };
      } else if (tn === "vt:i4") {
        value = { type: "number", value: parseInt(tv, 10) };
      } else if (tn === "vt:r8") {
        value = { type: "number", value: parseFloat(tv) };
      } else if (tn === "vt:bool") {
        value = { type: "boolean", value: tv === "true" };
      } else if (tn === "vt:filetime") {
        value = { type: "date", value: new Date(tv) };
      }
    }

    if (value) {
      props.push({ name, value });
    }
  }

  return props;
}

// =============================================================================
// Font Table Parser
// =============================================================================

function parseFontTableXml(xmlStr: string): FontDef[] {
  const doc = parseXml(xmlStr);
  const root = doc.root;
  const fonts: FontDef[] = [];

  for (const fontEl of findChildrenNs(root, "font")) {
    const f: Mutable<FontDef> = { name: attrVal(fontEl, "name") ?? "" };
    const p1 = findChildNs(fontEl, "panose1");
    if (p1) {
      f.panose1 = attrVal(p1, "val");
    }
    const cs = findChildNs(fontEl, "charset");
    if (cs) {
      f.charset = attrVal(cs, "val");
    }
    const fam = findChildNs(fontEl, "family");
    if (fam) {
      f.family = attrVal(fam, "val") as FontDef["family"];
    }
    const pitch = findChildNs(fontEl, "pitch");
    if (pitch) {
      f.pitch = attrVal(pitch, "val") as FontDef["pitch"];
    }
    // Signature
    const sigEl = findChildNs(fontEl, "sig");
    if (sigEl) {
      const sig: Record<string, string> = {};
      for (const key of ["usb0", "usb1", "usb2", "usb3", "csb0", "csb1"]) {
        const v = attrVal(sigEl, key);
        if (v !== undefined) {
          sig[key] = v;
        }
      }
      if (Object.keys(sig).length > 0) {
        f.sig = sig;
      }
    }
    // Embedded fonts
    for (const [tag, rIdKey, keyKey] of [
      ["embedRegular", "embedRegular", "embedRegularKey"],
      ["embedBold", "embedBold", "embedBoldKey"],
      ["embedItalic", "embedItalic", "embedItalicKey"],
      ["embedBoldItalic", "embedBoldItalic", "embedBoldItalicKey"]
    ] as const) {
      const el = findChildNs(fontEl, tag);
      if (el) {
        const rId = el.attributes["r:id"] ?? el.attributes["id"];
        if (rId) {
          f[rIdKey] = rId;
          const fontKey = attrVal(el, "fontKey");
          if (fontKey) {
            f[keyKey] = fontKey;
          }
        }
      }
    }
    fonts.push(f);
  }

  return fonts;
}

export { parseCoreProps, parseAppProps, parseCustomPropsXml, parseFontTableXml };
