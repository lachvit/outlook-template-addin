import {
  absolutizeUrls,
  extractBody,
  htmlToPlainText,
  prepareTemplate,
  sanitizeHtml,
} from "../src/taskpane/lib/templateHtml";

const TEMPLATE_URL = "https://contoso.github.io/repo/templates/welcome.html";

describe("extractBody", () => {
  it("returns only the body content", () => {
    const html =
      '<!DOCTYPE html><html><head><meta name="template-name" content="X"><title>T</title></head>' +
      "<body><p>Hello</p></body></html>";
    expect(extractBody(html)).toBe("<p>Hello</p>");
  });
});

describe("sanitizeHtml", () => {
  it("removes scripts, event handlers, javascript URLs and style blocks", () => {
    const result = sanitizeHtml(
      '<p onclick="alert(1)">Hi</p><script>alert(2)</script><a href="javascript:alert(3)">x</a><style>p{}</style>',
    );
    expect(result).not.toMatch(/script|onclick|javascript:|<style/i);
    expect(result).toContain("<p>Hi</p>");
  });

  it("keeps images, links and inline styles", () => {
    const result = sanitizeHtml('<p style="color: red">A <a href="https://x.test">b</a></p><img src="images/logo.png" alt="Logo">');
    expect(result).toContain('style="color: red"');
    expect(result).toContain('href="https://x.test"');
    expect(result).toContain('src="images/logo.png"');
  });
});

describe("absolutizeUrls", () => {
  it("makes relative image and link URLs absolute", () => {
    const result = absolutizeUrls('<img src="images/logo.png"><a href="../docs/a.pdf">A</a>', TEMPLATE_URL);
    expect(result).toContain('src="https://contoso.github.io/repo/templates/images/logo.png"');
    expect(result).toContain('href="https://contoso.github.io/repo/docs/a.pdf"');
  });

  it("leaves absolute, mailto and anchor URLs alone", () => {
    const html = '<img src="https://cdn.test/a.png"><a href="mailto:a@b.test">m</a><a href="#top">t</a>';
    expect(absolutizeUrls(html, TEMPLATE_URL)).toBe(html);
  });
});

describe("htmlToPlainText", () => {
  it("keeps paragraphs, line breaks and list items", () => {
    expect(htmlToPlainText("<p>Hello</p><p>World<br>Line</p><ul><li>One</li><li>Two</li></ul>")).toBe(
      "Hello\n\nWorld\nLine\n\n- One\n- Two",
    );
  });

  it("appends link targets that differ from the link text", () => {
    expect(htmlToPlainText('<p>See <a href="https://x.test/a">the docs</a></p>')).toBe("See the docs (https://x.test/a)");
  });
});

describe("prepareTemplate", () => {
  it("extracts, sanitizes and absolutizes in one step", () => {
    const raw =
      '<html><head><meta name="template-name" content="X"></head>' +
      '<body><p onmouseover="x()">Hi</p><img src="images/logo.png" alt="Logo"></body></html>';
    const result = prepareTemplate(raw, TEMPLATE_URL);
    expect(result.html).not.toContain("onmouseover");
    expect(result.html).not.toContain("template-name");
    expect(result.html).toContain('src="https://contoso.github.io/repo/templates/images/logo.png"');
    expect(result.text).toBe("Hi");
  });
});
