import { filterTemplates, getCategories } from "../src/taskpane/lib/filter";
import { TemplateInfo } from "../src/taskpane/lib/templateItems";

function template(id: string, category: string, name: string, description = "", keywords: string[] = []): TemplateInfo {
  return { id, name, category, description, keywords, fileName: `${id}.html`, lastModified: "" };
}

const templates = [
  template("welcome", "HR", "Welcome aboard", "First day information", ["onboarding", "new hire"]),
  template("thanks", "Customer Service", "Thank you for contacting us", "Standard reply", ["support"]),
  template("pricing", "Sales", "Product pricing", "Price list follow-up", ["quote", "offer"]),
];

describe("getCategories", () => {
  it("returns unique sorted categories", () => {
    expect(getCategories([...templates, template("x", "HR", "Other")])).toEqual(["Customer Service", "HR", "Sales"]);
  });
});

describe("filterTemplates", () => {
  it("returns everything for an empty query and category", () => {
    expect(filterTemplates(templates, "", "")).toHaveLength(3);
  });

  it("matches name, description and keywords case-insensitively", () => {
    expect(filterTemplates(templates, "WELCOME", "").map((t) => t.id)).toEqual(["welcome"]);
    expect(filterTemplates(templates, "standard", "").map((t) => t.id)).toEqual(["thanks"]);
    expect(filterTemplates(templates, "quote", "").map((t) => t.id)).toEqual(["pricing"]);
  });

  it("requires every word to match", () => {
    expect(filterTemplates(templates, "new hire", "").map((t) => t.id)).toEqual(["welcome"]);
    expect(filterTemplates(templates, "welcome quote", "")).toHaveLength(0);
  });

  it("filters by category", () => {
    expect(filterTemplates(templates, "", "Sales").map((t) => t.id)).toEqual(["pricing"]);
    expect(filterTemplates(templates, "welcome", "Sales")).toHaveLength(0);
  });
});
