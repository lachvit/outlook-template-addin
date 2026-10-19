export interface TemplateInfo {
  /** SharePoint driveItem id of the file. */
  id: string;
  name: string;
  category: string;
  description: string;
  keywords: string[];
  fileName: string;
  lastModified: string;
}

/** Internal names of the library columns the add-in reads (created by scripts/Initialize-TemplateLibrary.ps1). */
export const TEMPLATE_FIELDS = ["FileLeafRef", "Title", "TemplateCategory", "TemplateDescription", "TemplateKeywords"];

export const UNCATEGORIZED = "Uncategorized";

const TEMPLATE_FILE = /\.html?$/i;

function asString(value: unknown): string {
  return typeof value === "string" ? value.trim() : "";
}

/** Choice columns return a string; multi-choice columns return an array. */
function asFirstString(value: unknown): string {
  return Array.isArray(value) ? (value.map(asString).find((v) => v.length > 0) ?? "") : asString(value);
}

/** Accepts "a, b; c" text or a multi-choice array. */
function asKeywords(value: unknown): string[] {
  const list: unknown[] = Array.isArray(value) ? value : [value];
  return list
    .flatMap((v) => asString(v).split(/[,;]/))
    .map((k) => k.trim())
    .filter((k) => k.length > 0);
}

/**
 * Turns SharePoint files (driveItems with their list item columns) into templates.
 * Each entry is { id, lastModifiedDateTime, fields } where fields holds FileLeafRef and the library columns.
 * Only .html/.htm files are kept. The Title column is the display name (falls back to the file name)
 * and an empty category becomes "Uncategorized" so the template stays visible to authors.
 */
export function parseTemplateItems(items: unknown): TemplateInfo[] {
  if (!Array.isArray(items)) {
    throw new Error("The template list is not in the expected format.");
  }

  const result: TemplateInfo[] = [];
  for (const raw of items) {
    if (!raw || typeof raw !== "object") {
      continue;
    }
    const entry = raw as Record<string, unknown>;
    const fields = (entry.fields && typeof entry.fields === "object" ? entry.fields : {}) as Record<string, unknown>;
    const id = asString(entry.id);
    const fileName = asString(fields.FileLeafRef);
    if (!id || !TEMPLATE_FILE.test(fileName)) {
      continue;
    }
    result.push({
      id,
      name: asString(fields.Title) || fileName.replace(TEMPLATE_FILE, ""),
      category: asFirstString(fields.TemplateCategory) || UNCATEGORIZED,
      description: asString(fields.TemplateDescription),
      keywords: asKeywords(fields.TemplateKeywords),
      fileName,
      lastModified: asString(entry.lastModifiedDateTime),
    });
  }

  return result.sort((a, b) => a.category.localeCompare(b.category) || a.name.localeCompare(b.name));
}
