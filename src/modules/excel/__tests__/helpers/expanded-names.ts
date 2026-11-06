/**
 * A namespace-aware reading of XML text for tests: names resolved to URIs, and the attribute values that
 * name prefixes (ECMA-376 Part 3, `xsi:type`) resolved with them. Written independently of the
 * normaliser, so the two cannot share a mistake.
 */
import { SaxParser } from "@xml/sax";

const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const XSI = "http://www.w3.org/2001/XMLSchema-instance";

/**
 * How an attribute value names namespaces, decided by the URIs the names resolve to (ECMA-376 Part 3):
 * written independently of the normaliser, so the two cannot share a mistake about which attribute is which.
 */
export function kindOf(
  element: string,
  key: string,
  scope: ReadonlyMap<string, string>
): "prefixes" | "qnames" | "qname" | undefined {
  const uriOf = (qname: string) =>
    scope.get(qname.includes(":") ? qname.slice(0, qname.indexOf(":")) : "");
  const local = (qname: string) => qname.slice(qname.indexOf(":") + 1);
  if (!key.includes(":")) {
    return key === "Requires" && uriOf(element) === MC && local(element) === "Choice"
      ? "prefixes"
      : undefined;
  }
  const uri = uriOf(key);
  if (uri === MC) {
    return ["Ignorable", "MustUnderstand"].includes(local(key))
      ? "prefixes"
      : ["ProcessContent", "PreserveElements", "PreserveAttributes"].includes(local(key))
        ? "qnames"
        : undefined;
  }
  return uri === XSI && local(key) === "type" ? "qname" : undefined;
}

/** Walk `xml` with the declarations in force; `visit` sees each start tag and its scope. */
export function walk(
  xml: string,
  visit: (
    tag: { name: string; attributes: Record<string, string>; isSelfClosing: boolean },
    scope: ReadonlyMap<string, string>
  ) => void,
  close?: (name: string) => void
): void {
  const parser = new SaxParser({ position: false });
  const stack: Map<string, string>[] = [new Map([["xml", "http://www.w3.org/XML/1998/namespace"]])];
  parser.on("opentag", tag => {
    const scope = new Map(stack[stack.length - 1]);
    for (const key in tag.attributes) {
      if (key === "xmlns") {
        scope.set("", tag.attributes[key]);
      } else if (key.startsWith("xmlns:")) {
        scope.set(key.slice(6), tag.attributes[key]);
      }
    }
    stack.push(scope);
    visit(tag, scope);
  });
  parser.on("closetag", tag => {
    stack.pop();
    close?.(tag.name);
  });
  parser.on("error", error => {
    throw error;
  });
  parser.write(xml);
  parser.close();
}

/**
 * Every name **and every prefix named in a value** resolved to its namespace URI — so two documents that
 * mean the same thing give the same list. Throws on a prefix nothing declares, in a name or in a value.
 *
 * With `within` (an expanded name, `{uri}local`), only the first element of that name and its subtree
 * are listed: how a fragment reads **in the part it was written into**, its ancestors' declarations and
 * all.
 */
export function expandedNames(xml: string, within?: string): string[] {
  const names: string[] = [];
  // Inside `within`: the depth reached in it, 0 before it is found, -1 once it has closed.
  let depth = within === undefined ? 1 : 0;
  walk(
    xml,
    (tag, scope) => {
      const uri = (prefix: string, where: string) => {
        const bound = scope.get(prefix);
        if (bound === undefined || (prefix !== "" && bound === "")) {
          throw new Error(`undeclared prefix "${prefix}" in ${where}`);
        }
        return bound;
      };
      const qname = (name: string, isAttribute: boolean, where: string) => {
        const colon = name.indexOf(":");
        if (colon === -1) {
          return isAttribute ? name : `{${scope.get("") ?? ""}}${name}`;
        }
        return `{${uri(name.slice(0, colon), where)}}${name.slice(colon + 1)}`;
      };
      const name = qname(tag.name, false, tag.name);
      if (depth === 0 && name === within) {
        depth = 1;
      } else if (depth > 0 && within !== undefined) {
        depth++;
      }
      if (depth <= 0) {
        return;
      }
      names.push(name);
      const attributes: string[] = [];
      for (const key in tag.attributes) {
        if (key === "xmlns" || key.startsWith("xmlns:")) {
          continue;
        }
        const value = tag.attributes[key];
        const where = `${tag.name}/@${key}="${value}"`;
        const entries = value.split(/\s+/).filter(Boolean);
        const kind = kindOf(tag.name, key, scope);
        const expanded =
          kind === "prefixes"
            ? entries.map(prefix => `{${uri(prefix, where)}}`).join(" ")
            : kind === "qnames"
              ? entries.map(entry => qname(entry, true, where)).join(" ")
              : kind === "qname"
                ? qname(value.trim(), false, where)
                : value;
        attributes.push(`@${qname(key, true, where)}=${expanded}`);
      }
      names.push(...attributes.sort());
    },
    () => {
      if (within !== undefined && depth > 0) {
        depth = depth === 1 ? -1 : depth - 1;
      }
    }
  );
  if (within !== undefined && depth === 0) {
    throw new Error(`no ${within} in the part`);
  }
  return names;
}

function escapeText(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeAttribute(text: string): string {
  return escapeText(text)
    .replace(/"/g, "&quot;")
    .replace(/\t/g, "&#x9;")
    .replace(/\n/g, "&#xA;")
    .replace(/\r/g, "&#xD;");
}

/**
 * `xml` with each declared prefix renamed by `rename(prefix, uri)` — the default namespace too (`""`),
 * which `rename` may make a prefix — in names, declarations and the attribute values that name prefixes. The result means exactly
 * what `xml` means; only its spelling differs. An element in no namespace stays unprefixed, and
 * `xmlns=""` stays as it is: a prefix cannot be undeclared.
 */
export function reprefix(xml: string, rename: (prefix: string, uri: string) => string): string {
  const out: string[] = [];
  let open = false;
  const flush = () => {
    if (open) {
      out.push(">");
      open = false;
    }
  };
  const parser = new SaxParser({ position: false });
  const stack: Map<string, string>[] = [new Map()];
  parser.on("pi", (target, body) => {
    flush();
    out.push(`<?${target}${body ? ` ${body}` : ""}?>`);
  });
  parser.on("comment", text => {
    flush();
    out.push(`<!--${text}-->`);
  });
  parser.on("cdata", text => {
    flush();
    out.push(`<![CDATA[${text}]]>`);
  });
  parser.on("text", text => {
    flush();
    out.push(escapeText(text));
  });
  parser.on("opentag", tag => {
    flush();
    const scope = new Map(stack[stack.length - 1]);
    for (const key in tag.attributes) {
      if (key === "xmlns") {
        scope.set("", tag.attributes[key]);
      } else if (key.startsWith("xmlns:")) {
        scope.set(key.slice(6), tag.attributes[key]);
      }
    }
    stack.push(scope);
    const declared = (prefix: string) => (scope.get(prefix) ?? "") !== "";
    const qname = (name: string, isAttribute: boolean) => {
      const colon = name.indexOf(":");
      if (colon === -1) {
        return isAttribute || !declared("") ? name : `${to("")}:${name}`;
      }
      const prefix = name.slice(0, colon);
      return declared(prefix) ? `${to(prefix)}${name.slice(colon)}` : name;
    };
    const to = (entry: string) => rename(entry, scope.get(entry)!);
    const prefix = (entry: string) => (declared(entry) ? to(entry) : entry);
    out.push(`<${qname(tag.name, false)}`);
    for (const key in tag.attributes) {
      let name: string;
      let value = tag.attributes[key];
      if (key === "xmlns") {
        name = value === "" ? "xmlns" : `xmlns:${to("")}`;
      } else if (key.startsWith("xmlns:")) {
        name = `xmlns:${to(key.slice(6))}`;
      } else {
        name = qname(key, true);
        const kind = kindOf(tag.name, key, scope);
        const entries = value.split(/\s+/).filter(Boolean);
        if (kind === "prefixes") {
          value = entries.map(prefix).join(" ");
        } else if (kind === "qnames") {
          value = entries.map(entry => qname(entry, true)).join(" ");
        } else if (kind === "qname") {
          value = qname(value.trim(), false);
        }
      }
      out.push(` ${name}="${escapeAttribute(value)}"`);
    }
    open = true;
  });
  parser.on("closetag", tag => {
    const scope = stack[stack.length - 1];
    stack.pop();
    if (open) {
      out.push("/>");
      open = false;
      return;
    }
    const colon = tag.name.indexOf(":");
    const declared = (p: string) => (scope.get(p) ?? "") !== "";
    const to = (entry: string) => rename(entry, scope.get(entry)!);
    const name =
      colon === -1
        ? declared("")
          ? `${to("")}:${tag.name}`
          : tag.name
        : declared(tag.name.slice(0, colon))
          ? `${to(tag.name.slice(0, colon))}${tag.name.slice(colon)}`
          : tag.name;
    out.push(`</${name}>`);
  });
  parser.on("error", error => {
    throw error;
  });
  parser.write(xml);
  parser.close();
  return out.join("");
}
