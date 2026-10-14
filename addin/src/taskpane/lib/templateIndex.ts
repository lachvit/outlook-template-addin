export interface TemplateInfo {
  id: string;
  name: string;
  category: string;
  description: string;
  keywords: string[];
  url: string;
  lastModified: string;
}

function asString(value: unknown): string {
  return typeof value === "string" ? value.trim() : "";
}

function asKeywords(value: unknown): string[] {
  const list = Array.isArray(value) ? value : typeof value === "string" ? [value] : [];
  return list.map(asString).filter((k) => k.length > 0);
}

/**
 * Validates the JSON produced by scripts/Publish-TemplateIndex.ps1.
 * Entries that are incomplete or point outside the templates folder are dropped.
 */
export function parseTemplateIndex(data: unknown, templatesBaseUrl: string): TemplateInfo[] {
  const templates = (data as { templates?: unknown } | null)?.templates;
  if (!Array.isArray(templates)) {
    throw new Error("The template list is not in the expected format.");
  }

  const result: TemplateInfo[] = [];
  for (const raw of templates) {
    if (!raw || typeof raw !== "object") {
      continue;
    }
    const entry = raw as Record<string, unknown>;
    const template: TemplateInfo = {
      id: asString(entry.id),
      name: asString(entry.name),
      category: asString(entry.category),
      description: asString(entry.description),
      keywords: asKeywords(entry.keywords),
      url: asString(entry.url),
      lastModified: asString(entry.lastModified),
    };
    if (template.id && template.name && template.category && template.url.startsWith(templatesBaseUrl)) {
      result.push(template);
    }
  }

  return result.sort((a, b) => a.category.localeCompare(b.category) || a.name.localeCompare(b.name));
}
