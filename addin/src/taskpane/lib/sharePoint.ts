import { parseTemplateItems, TEMPLATE_FIELDS, TemplateInfo } from "./templateItems";

/** Returns a Microsoft Graph access token for the signed-in user. */
export type TokenProvider = () => Promise<string>;

export interface SharePointLocation {
  /** Full URL of the site, for example https://contoso.sharepoint.com/sites/EmailTemplates */
  siteUrl: string;
  /** Display name (or URL name) of the document library that holds the templates. */
  libraryName: string;
  /** Optional folder inside the library, for example "Email" or "Email/HR". Empty means the whole library. */
  folderPath?: string;
}

const GRAPH = "https://graph.microsoft.com/v1.0";
const MAX_RETRIES = 2;
const MAX_RETRY_DELAY_MS = 10000;
const MAX_FOLDER_DEPTH = 5;

/**
 * Turns a site URL into the URL-ready Graph path-based site key:
 * https://contoso.sharepoint.com/sites/Email Templates -> contoso.sharepoint.com:/sites/Email%20Templates
 */
export function siteKeyFromUrl(siteUrl: string): string {
  let url: URL;
  try {
    url = new URL(siteUrl);
  } catch {
    throw new Error(`"${siteUrl}" is not a valid SharePoint site URL.`);
  }
  if (url.protocol !== "https:") {
    throw new Error("The SharePoint site URL must start with https://.");
  }
  // pathname is already percent-encoded; decode each segment so it is encoded exactly once.
  const segments = url.pathname
    .split("/")
    .filter((s) => s.length > 0)
    .map((s) => encodeURIComponent(decodeURIComponent(s)));
  return segments.length > 0 ? `${url.hostname}:/${segments.join("/")}` : url.hostname;
}

function describeFailure(status: number, what: string): string {
  switch (status) {
    case 401:
      return `Could not load ${what} because your sign-in was not accepted. Close and reopen the add-in, then try again.`;
    case 403:
      return (
        `You don't have access to ${what}. Ask IT to check that you can open the templates site in SharePoint ` +
        "and that the add-in has been granted access to it (Sites.Selected)."
      );
    case 404:
      return `Could not find ${what}. Ask IT to check the SharePoint site URL and library name the add-in was built with.`;
    case 429:
    case 503:
      return `SharePoint is busy and could not return ${what}. Wait a minute and try again.`;
    default:
      return `Could not load ${what} (HTTP ${status}).`;
  }
}

function retryDelayMs(response: Response, attempt: number): number {
  const seconds = Number(response.headers.get("Retry-After"));
  const delay = Number.isFinite(seconds) && seconds > 0 ? seconds * 1000 : 1000 * 2 ** attempt;
  return Math.min(delay, MAX_RETRY_DELAY_MS);
}

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

/** Encodes a folder path ("Email/HR") for use in a Graph path-based address (root:/Email/HR:). */
export function encodeFolderPath(folderPath: string): string {
  return folderPath
    .split("/")
    .map((s) => s.trim())
    .filter((s) => s.length > 0)
    .map((s) => encodeURIComponent(s))
    .join("/");
}

interface DriveItem {
  id?: string;
  name?: string;
  lastModifiedDateTime?: string;
  folder?: unknown;
  file?: unknown;
  listItem?: { id?: string; fields?: Record<string, unknown> };
}

interface Ids {
  driveId: string;
  folderId: string;
}

/** Maps a driveItem with its expanded listItem to the shape parseTemplateItems expects. */
function toTemplateItem(item: DriveItem): unknown {
  return {
    id: item.id,
    lastModifiedDateTime: item.lastModifiedDateTime,
    fields: { ...(item.listItem?.fields ?? {}), FileLeafRef: item.name },
  };
}

export class SharePointTemplateStore {
  private ids?: Promise<Ids>;

  constructor(
    private readonly location: SharePointLocation,
    private readonly getToken: TokenProvider,
  ) {}

  /** Lists the .html files in the configured folder and its subfolders. */
  async listTemplates(): Promise<TemplateInfo[]> {
    const { driveId, folderId } = await this.resolveIds();
    const files: DriveItem[] = [];
    const pending: { id: string; depth: number }[] = [{ id: folderId, depth: 0 }];
    while (pending.length > 0) {
      const folder = pending.shift() as { id: string; depth: number };
      const url =
        `${GRAPH}/drives/${driveId}/items/${encodeURIComponent(folder.id)}/children` +
        "?$select=id,name,lastModifiedDateTime,file,folder" +
        `&$expand=listItem($select=id;$expand=fields($select=${TEMPLATE_FIELDS.join(",")}))&$top=200`;
      for (const item of (await this.getAllPages(url, "the template list")) as DriveItem[]) {
        if (item.folder && item.id && folder.depth < MAX_FOLDER_DEPTH) {
          pending.push({ id: item.id, depth: folder.depth + 1 });
        } else if (item.file) {
          files.push(item);
        }
      }
    }
    return parseTemplateItems(files.map(toTemplateItem));
  }

  async getTemplateHtml(templateId: string): Promise<string> {
    const { driveId } = await this.resolveIds();
    // The /content endpoint answers with a redirect that browsers can't follow with an Authorization
    // header (CORS), so ask for the short-lived pre-authenticated download URL instead.
    const item = await this.getJson<Record<string, unknown>>(
      `${GRAPH}/drives/${driveId}/items/${encodeURIComponent(templateId)}?$select=id,@microsoft.graph.downloadUrl`,
      "the template",
    );
    const downloadUrl = item["@microsoft.graph.downloadUrl"];
    if (typeof downloadUrl !== "string" || !downloadUrl.startsWith("https://")) {
      throw new Error("Could not load the template. SharePoint did not return a download link.");
    }
    const response = await this.send(downloadUrl, "the template", false);
    return response.text();
  }

  /** Looks up the library's drive and the folder once; a failure is not cached so "Try again" works. */
  private resolveIds(): Promise<Ids> {
    if (!this.ids) {
      this.ids = this.lookupIds().catch((error: unknown) => {
        this.ids = undefined;
        throw error;
      });
    }
    return this.ids;
  }

  private async lookupIds(): Promise<Ids> {
    const siteKey = siteKeyFromUrl(this.location.siteUrl);
    const site = await this.getJson<{ id?: string }>(`${GRAPH}/sites/${siteKey}?$select=id`, "the templates site");
    if (!site.id) {
      throw new Error("Could not find the templates site.");
    }
    const lists = await this.getAllPages(`${GRAPH}/sites/${site.id}/lists?$select=id,displayName,name`, "the templates library");
    const wanted = this.location.libraryName.trim().toLowerCase();
    const list = (lists as { id?: string; displayName?: string; name?: string }[]).find(
      (l) => l.displayName?.toLowerCase() === wanted || l.name?.toLowerCase() === wanted,
    );
    if (!list?.id) {
      throw new Error(`The "${this.location.libraryName}" library was not found on the templates site.`);
    }
    const drive = await this.getJson<{ id?: string }>(
      `${GRAPH}/sites/${site.id}/lists/${list.id}/drive?$select=id`,
      "the templates library",
    );
    if (!drive.id) {
      throw new Error(`The "${this.location.libraryName}" library is not a document library.`);
    }
    const folderPath = encodeFolderPath(this.location.folderPath ?? "");
    const folderUrl = folderPath
      ? `${GRAPH}/drives/${drive.id}/root:/${folderPath}?$select=id,folder`
      : `${GRAPH}/drives/${drive.id}/root?$select=id,folder`;
    const folderLabel = folderPath ? `the "${this.location.folderPath}" folder` : "the templates library";
    const folder = await this.getJson<DriveItem>(folderUrl, folderLabel);
    if (!folder.id || !folder.folder) {
      throw new Error(`"${this.location.folderPath}" in the "${this.location.libraryName}" library is not a folder.`);
    }
    return { driveId: drive.id, folderId: folder.id };
  }

  private async getAllPages(firstUrl: string, what: string): Promise<unknown[]> {
    const all: unknown[] = [];
    let url: string | undefined = firstUrl;
    while (url) {
      const page: { value?: unknown; "@odata.nextLink"?: unknown } = await this.getJson(url, what);
      if (!Array.isArray(page.value)) {
        throw new Error(`${what} is not in the expected format.`);
      }
      all.push(...page.value);
      const next = page["@odata.nextLink"];
      url = typeof next === "string" && next.startsWith(`${GRAPH}/`) ? next : undefined;
    }
    return all;
  }

  private async getJson<T>(url: string, what: string): Promise<T> {
    const response = await this.send(url, what, true);
    return (await response.json()) as T;
  }

  private async send(url: string, what: string, authenticated: boolean): Promise<Response> {
    const headers: Record<string, string> = {};
    if (authenticated) {
      headers.Authorization = `Bearer ${await this.getToken()}`;
    }
    for (let attempt = 0; ; attempt++) {
      let response: Response;
      try {
        response = await fetch(url, { headers, cache: "no-store" });
      } catch {
        throw new Error(`Could not load ${what}. Check your network connection and try again.`);
      }
      if (response.ok) {
        return response;
      }
      if ((response.status === 429 || response.status === 503) && attempt < MAX_RETRIES) {
        await sleep(retryDelayMs(response, attempt));
        continue;
      }
      throw new Error(describeFailure(response.status, what));
    }
  }
}
