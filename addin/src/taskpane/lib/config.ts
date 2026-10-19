/** Application (client) ID of the Entra ID app registration. */
export const AUTH_CLIENT_ID: string = __AUTH_CLIENT_ID__;

/** Directory (tenant) ID of the Microsoft 365 tenant. */
export const AUTH_TENANT_ID: string = __AUTH_TENANT_ID__;

/** SharePoint site that holds the templates library, for example https://contoso.sharepoint.com/sites/EmailTemplates */
export const SHAREPOINT_SITE_URL: string = __SHAREPOINT_SITE_URL__;

/** Display name of the document library with the template files. */
export const SHAREPOINT_LIBRARY: string = __SHAREPOINT_LIBRARY__;

/** Optional folder inside the library (for example "Email"); empty means the whole library. */
export const SHAREPOINT_FOLDER: string = __SHAREPOINT_FOLDER__;

/** Lists the build settings that are missing, so a misconfigured build explains itself in the task pane. */
export function getMissingSettings(): string[] {
  const settings: Record<string, string> = {
    AUTH_CLIENT_ID,
    AUTH_TENANT_ID,
    SHAREPOINT_SITE_URL,
    SHAREPOINT_LIBRARY,
  };
  return Object.keys(settings).filter((name) => !settings[name]);
}
