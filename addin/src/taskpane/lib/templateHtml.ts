import DOMPurify from "dompurify";

export interface PreparedTemplate {
  /** Sanitized body HTML with absolute image and link URLs, ready to insert. */
  html: string;
  /** Text-only version for plain-text messages. */
  text: string;
}

const BLOCK_TAGS = new Set([
  "ADDRESS", "BLOCKQUOTE", "DIV", "H1", "H2", "H3", "H4", "H5", "H6", "HR",
  "OL", "P", "PRE", "SECTION", "TABLE", "TR", "UL",
]);

function parse(html: string): Document {
  return new DOMParser().parseFromString(html, "text/html");
}

/** Returns everything inside <body>; the <head> (metadata) is never inserted. */
export function extractBody(html: string): string {
  return parse(html).body.innerHTML.trim();
}

export function sanitizeHtml(html: string): string {
  return String(
    DOMPurify.sanitize(html, {
      USE_PROFILES: { html: true },
      // <style> in the body would restyle the whole message; form controls do not work in email.
      FORBID_TAGS: ["style", "form", "input", "button", "select", "textarea"],
    }),
  );
}

/** Turns relative img src and link href values into absolute URLs based on the template's own URL. */
export function absolutizeUrls(html: string, baseUrl: string): string {
  const doc = parse(html);
  const rewrite = (element: Element, attribute: string): void => {
    const value = element.getAttribute(attribute);
    if (!value || value.startsWith("#")) {
      return;
    }
    try {
      element.setAttribute(attribute, new URL(value, baseUrl).href);
    } catch {
      // Leave values that are not valid URLs unchanged.
    }
  };
  doc.body.querySelectorAll("img[src]").forEach((el) => rewrite(el, "src"));
  doc.body.querySelectorAll("a[href]").forEach((el) => rewrite(el, "href"));
  return doc.body.innerHTML;
}

export function htmlToPlainText(html: string): string {
  let out = "";
  const walk = (node: Node): void => {
    if (node.nodeType === Node.TEXT_NODE) {
      out += (node.textContent ?? "").replace(/\s+/g, " ");
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) {
      return;
    }
    const element = node as Element;
    const tag = element.tagName;
    if (tag === "BR") {
      out += "\n";
      return;
    }
    if (tag === "LI") {
      out += "\n- ";
    } else if (BLOCK_TAGS.has(tag)) {
      out += "\n";
    }
    element.childNodes.forEach(walk);
    if (tag === "A") {
      const href = element.getAttribute("href") ?? "";
      if (/^https?:/i.test(href) && href !== element.textContent?.trim()) {
        out += ` (${href})`;
      }
    } else if (tag === "TD" || tag === "TH") {
      out += " ";
    } else if (BLOCK_TAGS.has(tag)) {
      out += "\n";
    }
  };
  walk(parse(html).body);
  return out
    .split("\n")
    .map((line) => line.replace(/ {2,}/g, " ").trim())
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

/** Full pipeline: body only, sanitized, URLs made absolute, plus a text version. */
export function prepareTemplate(rawHtml: string, templateUrl: string): PreparedTemplate {
  const html = absolutizeUrls(sanitizeHtml(extractBody(rawHtml)), templateUrl);
  return { html, text: htmlToPlainText(html) };
}
