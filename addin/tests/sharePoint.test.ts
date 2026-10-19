import { encodeFolderPath, SharePointTemplateStore, siteKeyFromUrl } from "../src/taskpane/lib/sharePoint";

const GRAPH = "https://graph.microsoft.com/v1.0";
const SITE_ID = "contoso.sharepoint.com,11111111-1111-1111-1111-111111111111,22222222-2222-2222-2222-222222222222";
const LIST_ID = "33333333-3333-3333-3333-333333333333";
const DRIVE_ID = "b!drive";

type Route = (url: string, init?: RequestInit) => { status?: number; body?: unknown; text?: string } | undefined;

function jsonResponse(status: number, body: unknown, text?: string): Response {
  return {
    ok: status >= 200 && status < 300,
    status,
    headers: { get: () => null },
    json: async () => body,
    text: async () => text ?? JSON.stringify(body),
  } as unknown as Response;
}

function mockFetch(route: Route) {
  const calls: { url: string; init?: RequestInit }[] = [];
  globalThis.fetch = jest.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = String(input);
    calls.push({ url, init });
    const result = route(url, init);
    if (!result) {
      return jsonResponse(404, { error: { code: "itemNotFound" } });
    }
    return jsonResponse(result.status ?? 200, result.body, result.text);
  }) as unknown as typeof fetch;
  return calls;
}

function file(id: string, name: string, fields: Record<string, unknown> = {}) {
  return { id, name, file: {}, lastModifiedDateTime: "2026-01-01T00:00:00Z", listItem: { id: `li-${id}`, fields } };
}

const standardRoutes: Route = (url) => {
  if (url === `${GRAPH}/sites/contoso.sharepoint.com:/sites/Office%20Templates?$select=id`) {
    return { body: { id: SITE_ID } };
  }
  if (url.startsWith(`${GRAPH}/sites/${SITE_ID}/lists?`)) {
    return {
      body: {
        value: [
          { id: "other", displayName: "Site Pages", name: "SitePages" },
          { id: LIST_ID, displayName: "Documents", name: "Shared Documents" },
        ],
      },
    };
  }
  if (url === `${GRAPH}/sites/${SITE_ID}/lists/${LIST_ID}/drive?$select=id`) {
    return { body: { id: DRIVE_ID } };
  }
  if (url === `${GRAPH}/drives/${DRIVE_ID}/root:/Email/Team%20A?$select=id,folder`) {
    return { body: { id: "folder-email", folder: { childCount: 3 } } };
  }
  if (url === `${GRAPH}/drives/${DRIVE_ID}/root?$select=id,folder`) {
    return { body: { id: "root", folder: {} } };
  }
  if (url.startsWith(`${GRAPH}/drives/${DRIVE_ID}/items/folder-email/children?`)) {
    return {
      body: {
        value: [
          file("f1", "Email-Template.html"),
          { id: "sub", name: "HR", folder: { childCount: 1 } },
          file("f2", "notes.docx"),
        ],
        "@odata.nextLink": `${GRAPH}/drives/${DRIVE_ID}/items/folder-email/children/page2`,
      },
    };
  }
  if (url === `${GRAPH}/drives/${DRIVE_ID}/items/folder-email/children/page2`) {
    return { body: { value: [file("f3", "thanks.html", { Title: "Thanks", TemplateCategory: "CS" })] } };
  }
  if (url.startsWith(`${GRAPH}/drives/${DRIVE_ID}/items/sub/children?`)) {
    return { body: { value: [file("f4", "welcome.htm", { Title: "Welcome", TemplateCategory: "HR" })] } };
  }
  if (url === `${GRAPH}/drives/${DRIVE_ID}/items/f1`) {
    return { body: { id: "f1", "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/download.aspx?tempauth=abc" } };
  }
  if (url === "https://contoso.sharepoint.com/download.aspx?tempauth=abc") {
    return { text: "<html><body><p>Hi</p></body></html>" };
  }
  return undefined;
};

const location = {
  siteUrl: "https://contoso.sharepoint.com/sites/Office Templates/",
  libraryName: "documents",
  folderPath: "/Email/Team A/",
};

describe("siteKeyFromUrl", () => {
  it("builds the URL-ready Graph path-based site key", () => {
    expect(siteKeyFromUrl("https://contoso.sharepoint.com/sites/Templates")).toBe("contoso.sharepoint.com:/sites/Templates");
    expect(siteKeyFromUrl("https://contoso.sharepoint.com/sites/Templates/")).toBe("contoso.sharepoint.com:/sites/Templates");
    expect(siteKeyFromUrl("https://contoso.sharepoint.com/sites/Email Templates")).toBe("contoso.sharepoint.com:/sites/Email%20Templates");
    expect(siteKeyFromUrl("https://contoso.sharepoint.com/sites/Email%20Templates")).toBe("contoso.sharepoint.com:/sites/Email%20Templates");
    expect(siteKeyFromUrl("https://contoso.sharepoint.com/")).toBe("contoso.sharepoint.com");
  });

  it("rejects invalid and non-https URLs", () => {
    expect(() => siteKeyFromUrl("not a url")).toThrow(/not a valid/);
    expect(() => siteKeyFromUrl("http://contoso.sharepoint.com/sites/x")).toThrow(/https/);
  });
});

describe("encodeFolderPath", () => {
  it("trims slashes and encodes each segment", () => {
    expect(encodeFolderPath("/Email/Team A/")).toBe("Email/Team%20A");
    expect(encodeFolderPath("")).toBe("");
    expect(encodeFolderPath("a#b")).toBe("a%23b");
  });
});

describe("SharePointTemplateStore", () => {
  it("lists html files in the folder and its subfolders, across pages", async () => {
    const calls = mockFetch(standardRoutes);
    const store = new SharePointTemplateStore(location, async () => "token-1");

    const templates = await store.listTemplates();

    expect(templates.map((t) => [t.id, t.name, t.category])).toEqual([
      ["f3", "Thanks", "CS"],
      ["f4", "Welcome", "HR"],
      ["f1", "Email-Template", "Uncategorized"],
    ]);
    expect(calls.every((c) => (c.init?.headers as Record<string, string>).Authorization === "Bearer token-1")).toBe(true);
    const childrenCall = calls.find((c) => c.url.includes("/children?"));
    expect(childrenCall?.url).toContain(
      "$expand=listItem($select=id;$expand=fields($select=FileLeafRef,Title,TemplateCategory,TemplateDescription,TemplateKeywords))",
    );
  });

  it("uses the library root when no folder is set", async () => {
    const calls = mockFetch((url, init) =>
      url.includes("/items/root/children?") ? { body: { value: [file("r1", "a.html")] } } : standardRoutes(url, init),
    );
    const store = new SharePointTemplateStore({ ...location, folderPath: "" }, async () => "t");
    await expect(store.listTemplates()).resolves.toHaveLength(1);
    expect(calls.some((c) => c.url === `${GRAPH}/drives/${DRIVE_ID}/root?$select=id,folder`)).toBe(true);
  });

  it("downloads template HTML through the pre-authenticated URL without the token", async () => {
    const calls = mockFetch(standardRoutes);
    const store = new SharePointTemplateStore(location, async () => "token-1");

    await expect(store.getTemplateHtml("f1")).resolves.toBe("<html><body><p>Hi</p></body></html>");

    const download = calls.find((c) => c.url.startsWith("https://contoso.sharepoint.com/download.aspx"));
    expect((download?.init?.headers as Record<string, string>).Authorization).toBeUndefined();
    // Site, library and folder are looked up only once.
    await store.getTemplateHtml("f1");
    expect(calls.filter((c) => c.url.includes("/lists?")).length).toBe(1);
  });

  it("matches the library by URL name too", async () => {
    mockFetch(standardRoutes);
    const store = new SharePointTemplateStore({ ...location, libraryName: "Shared Documents" }, async () => "t");
    await expect(store.listTemplates()).resolves.toHaveLength(3);
  });

  it("reports a missing library by name", async () => {
    mockFetch(standardRoutes);
    const store = new SharePointTemplateStore({ ...location, libraryName: "Nope" }, async () => "t");
    await expect(store.listTemplates()).rejects.toThrow('The "Nope" library was not found');
  });

  it("reports a missing folder", async () => {
    mockFetch(standardRoutes);
    const store = new SharePointTemplateStore({ ...location, folderPath: "Missing" }, async () => "t");
    await expect(store.listTemplates()).rejects.toThrow(/Could not find the "Missing" folder/);
  });

  it("rejects a folder path that points to a file", async () => {
    mockFetch((url, init) =>
      url.includes("/root:/Email/Team%20A?") ? { body: { id: "x", file: {} } } : standardRoutes(url, init),
    );
    const store = new SharePointTemplateStore(location, async () => "t");
    await expect(store.listTemplates()).rejects.toThrow(/is not a folder/);
  });

  it("explains access denied and retries the lookup after a failure", async () => {
    let denied = true;
    mockFetch((url, init) => (denied ? { status: 403, body: {} } : standardRoutes(url, init)));
    const store = new SharePointTemplateStore(location, async () => "t");

    await expect(store.listTemplates()).rejects.toThrow(/don't have access.*Sites\.Selected/);
    denied = false;
    await expect(store.listTemplates()).resolves.toHaveLength(3);
  });

  it("falls back to /content with the token when there is no usable download link", async () => {
    for (const body of [{ id: "f1" }, { id: "f1", "@microsoft.graph.downloadUrl": "http://evil.test/x" }]) {
      const calls = mockFetch((url, init) => {
        if (url === `${GRAPH}/drives/${DRIVE_ID}/items/f1`) {
          return { body };
        }
        if (url === `${GRAPH}/drives/${DRIVE_ID}/items/f1/content`) {
          return { text: "<p>From content</p>" };
        }
        return standardRoutes(url, init);
      });
      const store = new SharePointTemplateStore(location, async () => "t");

      await expect(store.getTemplateHtml("f1")).resolves.toBe("<p>From content</p>");
      const content = calls.find((c) => c.url.endsWith("/items/f1/content"));
      expect((content?.init?.headers as Record<string, string>).Authorization).toBe("Bearer t");
      expect(calls.some((c) => c.url.startsWith("http://evil.test"))).toBe(false);
    }
  });

  it("does not follow nextLink URLs outside Microsoft Graph", async () => {
    const calls = mockFetch((url, init) =>
      url.includes("/folder-email/children?")
        ? { body: { value: [], "@odata.nextLink": "https://evil.test/steal" } }
        : standardRoutes(url, init),
    );
    const store = new SharePointTemplateStore(location, async () => "t");
    await expect(store.listTemplates()).resolves.toEqual([]);
    expect(calls.some((c) => c.url.startsWith("https://evil.test"))).toBe(false);
  });
});
