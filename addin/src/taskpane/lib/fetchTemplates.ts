import { parseTemplateIndex, TemplateInfo } from "./templateIndex";

/** Adds ?v=<build id> so each deployment bypasses GitHub Pages' ~10 minute cache. */
export function withVersion(url: string, buildId: string): string {
  const result = new URL(url);
  result.searchParams.set("v", buildId);
  return result.href;
}

async function getResponse(url: string, what: string): Promise<Response> {
  let response: Response;
  try {
    response = await fetch(url, { cache: "no-cache" });
  } catch {
    throw new Error(`Could not load ${what}. Check your network connection and try again.`);
  }
  if (!response.ok) {
    throw new Error(`Could not load ${what} (HTTP ${response.status}).`);
  }
  return response;
}

export async function fetchTemplateIndex(templatesBaseUrl: string, buildId: string): Promise<TemplateInfo[]> {
  const url = withVersion(new URL("index.json", templatesBaseUrl).href, buildId);
  const response = await getResponse(url, "the template list");
  return parseTemplateIndex(await response.json(), templatesBaseUrl);
}

export async function fetchTemplateHtml(templateUrl: string, buildId: string): Promise<string> {
  const response = await getResponse(withVersion(templateUrl, buildId), "the template");
  return response.text();
}
