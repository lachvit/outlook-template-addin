import { TemplateInfo } from "./templateItems";

export function getCategories(templates: TemplateInfo[]): string[] {
  return Array.from(new Set(templates.map((t) => t.category))).sort((a, b) => a.localeCompare(b));
}

/**
 * Every word in the query must appear in the name, description or a keyword.
 * An empty category means "all categories".
 */
export function filterTemplates(templates: TemplateInfo[], query: string, category: string): TemplateInfo[] {
  const terms = query.toLowerCase().split(/\s+/).filter((t) => t.length > 0);
  return templates.filter((template) => {
    if (category && template.category !== category) {
      return false;
    }
    const fields = [template.name, template.description, ...template.keywords].map((f) => f.toLowerCase());
    return terms.every((term) => fields.some((field) => field.includes(term)));
  });
}
