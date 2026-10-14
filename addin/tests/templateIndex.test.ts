import { withVersion } from "../src/taskpane/lib/fetchTemplates";
import { parseTemplateIndex } from "../src/taskpane/lib/templateIndex";

const BASE = "https://contoso.github.io/templates-repo/templates/";

describe("parseTemplateIndex", () => {
  it("returns valid entries sorted by category, then name", () => {
    const result = parseTemplateIndex(
      {
        templates: [
          { id: "b", name: "Zeta", category: "Sales", url: `${BASE}b.html`, keywords: ["one", " two "] },
          { id: "a", name: "Alpha", category: "Sales", url: `${BASE}a.html` },
          { id: "c", name: "Welcome", category: "HR", url: `${BASE}c.html`, description: "Hi" },
        ],
      },
      BASE,
    );
    expect(result.map((t) => t.id)).toEqual(["c", "a", "b"]);
    expect(result[2].keywords).toEqual(["one", "two"]);
    expect(result[1].keywords).toEqual([]);
    expect(result[0].description).toBe("Hi");
  });

  it("accepts a single keyword serialized as a string", () => {
    const [t] = parseTemplateIndex(
      { templates: [{ id: "a", name: "A", category: "C", url: `${BASE}a.html`, keywords: "solo" }] },
      BASE,
    );
    expect(t.keywords).toEqual(["solo"]);
  });

  it("drops incomplete entries and URLs outside the templates folder", () => {
    const result = parseTemplateIndex(
      {
        templates: [
          { id: "a", name: "", category: "C", url: `${BASE}a.html` },
          { id: "b", name: "B", category: "C", url: "https://evil.example/b.html" },
          null,
          "nope",
          { id: "ok", name: "OK", category: "C", url: `${BASE}ok.html` },
        ],
      },
      BASE,
    );
    expect(result.map((t) => t.id)).toEqual(["ok"]);
  });

  it("throws when the document has no templates array", () => {
    expect(() => parseTemplateIndex({}, BASE)).toThrow();
    expect(() => parseTemplateIndex(null, BASE)).toThrow();
  });
});

describe("withVersion", () => {
  it("adds the build id as the v query parameter", () => {
    expect(withVersion(`${BASE}index.json`, "123-1")).toBe(`${BASE}index.json?v=123-1`);
  });

  it("replaces an existing v parameter and keeps others", () => {
    expect(withVersion(`${BASE}a.html?x=1&v=old`, "new")).toBe(`${BASE}a.html?x=1&v=new`);
  });
});
