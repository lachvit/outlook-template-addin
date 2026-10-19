import { parseTemplateItems, UNCATEGORIZED } from "../src/taskpane/lib/templateItems";

function item(id: string, fields: Record<string, unknown>, lastModifiedDateTime = "2026-01-02T03:04:05Z") {
  return { id, lastModifiedDateTime, fields };
}

describe("parseTemplateItems", () => {
  it("maps library columns and sorts by category, then name", () => {
    const result = parseTemplateItems([
      item("2", { FileLeafRef: "b.html", Title: "Zeta", TemplateCategory: "Sales", TemplateKeywords: "one, two ;three" }),
      item("1", { FileLeafRef: "a.html", Title: "Alpha", TemplateCategory: "Sales" }),
      item("3", { FileLeafRef: "c.htm", Title: " Welcome ", TemplateCategory: "HR", TemplateDescription: " Hi " }),
    ]);
    expect(result.map((t) => t.id)).toEqual(["3", "1", "2"]);
    expect(result[0]).toEqual({
      id: "3",
      name: "Welcome",
      category: "HR",
      description: "Hi",
      keywords: [],
      fileName: "c.htm",
      lastModified: "2026-01-02T03:04:05Z",
    });
    expect(result[2].keywords).toEqual(["one", "two", "three"]);
  });

  it("falls back to the file name and the Uncategorized category", () => {
    const [t] = parseTemplateItems([item("1", { FileLeafRef: "holiday-closure.HTML", Title: "", TemplateCategory: null })]);
    expect(t.name).toBe("holiday-closure");
    expect(t.category).toBe(UNCATEGORIZED);
  });

  it("accepts multi-choice category and keyword columns", () => {
    const [t] = parseTemplateItems([
      item("1", { FileLeafRef: "a.html", Title: "A", TemplateCategory: ["", "Sales", "HR"], TemplateKeywords: ["x", " y "] }),
    ]);
    expect(t.category).toBe("Sales");
    expect(t.keywords).toEqual(["x", "y"]);
  });

  it("skips folders, other file types and malformed entries", () => {
    const result = parseTemplateItems([
      item("1", { FileLeafRef: "Archive" }),
      item("2", { FileLeafRef: "logo.png", Title: "Logo" }),
      item("", { FileLeafRef: "no-id.html" }),
      { id: "4" },
      null,
      "nope",
      item("5", { FileLeafRef: "ok.html", Title: "OK", TemplateCategory: "C" }),
    ]);
    expect(result.map((t) => t.id)).toEqual(["5"]);
  });

  it("throws when the response is not a list", () => {
    expect(() => parseTemplateItems({})).toThrow();
    expect(() => parseTemplateItems(null)).toThrow();
  });
});
