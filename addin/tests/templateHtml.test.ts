import {
  extractBody,
  htmlToPlainText,
  prepareTemplate,
  removeRelativeUrls,
  sanitizeHtml,
} from "../src/taskpane/lib/templateHtml";

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

describe("removeRelativeUrls", () => {
  it("removes images whose source recipients could not load", () => {
    const result = removeRelativeUrls(
      '<p>A<img src="images/logo.png"><img src="/x.png"><img src="http://cdn.test/a.png"><img></p>',
    );
    expect(result).toBe("<p>A</p>");
  });

  it("keeps https, data and cid images", () => {
    const html = '<img src="https://cdn.test/a.png"><img src="data:image/png;base64,AAAA"><img src="cid:logo">';
    expect(removeRelativeUrls(html)).toBe(html);
  });

  it("unlinks relative links but keeps their text", () => {
    expect(removeRelativeUrls('<a href="../docs/a.pdf">A</a>')).toBe("<a>A</a>");
  });

  it("leaves absolute, mailto, tel and anchor links alone", () => {
    const html =
      '<a href="https://x.test">h</a><a href="http://x.test">p</a><a href="mailto:a@b.test">m</a>' +
      '<a href="tel:+15550100">t</a><a href="#top">a</a>';
    expect(removeRelativeUrls(html)).toBe(html);
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
  it("extracts, sanitizes and removes unreachable images in one step", () => {
    const raw =
      '<html><head><meta name="template-name" content="X"></head>' +
      '<body><p onmouseover="x()">Hi</p><img src="images/logo.png" alt="Logo">' +
      '<img src="https://www.contoso.test/logo.png" alt="Contoso"></body></html>';
    const result = prepareTemplate(raw);
    expect(result.html).not.toContain("onmouseover");
    expect(result.html).not.toContain("template-name");
    expect(result.html).not.toContain("images/logo.png");
    expect(result.html).toContain('src="https://www.contoso.test/logo.png"');
    expect(result.text).toBe("Hi");
  });
});
