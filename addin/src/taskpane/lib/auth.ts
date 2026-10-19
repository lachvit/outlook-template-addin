import {
  AccountInfo,
  AuthError,
  BrowserAuthErrorCodes,
  createNestablePublicClientApplication,
  IPublicClientApplication,
} from "@azure/msal-browser";
import { AUTH_CLIENT_ID, AUTH_TENANT_ID } from "./config";

/**
 * Delegated Microsoft Graph permission. With Sites.Selected the add-in can only read SharePoint sites
 * that an admin has explicitly granted to the app registration, and only what the user can access.
 */
export const GRAPH_SCOPES = ["https://graph.microsoft.com/Sites.Selected"];

/** Thrown when no token can be obtained without showing a sign-in window. */
export class SignInRequiredError extends Error {
  constructor(readonly cause?: unknown) {
    super("Sign in with your work account to see the templates.");
    this.name = "SignInRequiredError";
  }
}

let clientPromise: Promise<IPublicClientApplication> | undefined;

/**
 * Uses Nested App Authentication (NAA) when Outlook supports it: the token comes from the account
 * already signed in to Outlook. Otherwise MSAL falls back to a normal browser client with a popup.
 */
function getClient(): Promise<IPublicClientApplication> {
  if (!clientPromise) {
    clientPromise = createNestablePublicClientApplication({
      auth: {
        clientId: AUTH_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${AUTH_TENANT_ID}`,
        // Only used by the popup fallback; must be registered as a SPA redirect URI.
        redirectUri: new URL("auth.html", window.location.href).href,
      },
      cache: { cacheLocation: "localStorage" },
    }).catch((error: unknown) => {
      clientPromise = undefined;
      throw error;
    });
  }
  return clientPromise;
}

/** The mailbox owner's address, used as a login hint so the popup preselects the right account. */
function mailboxUser(): string | undefined {
  try {
    return Office.context.mailbox?.userProfile?.emailAddress || undefined;
  } catch {
    return undefined;
  }
}

function currentAccount(client: IPublicClientApplication): AccountInfo | undefined {
  return client.getActiveAccount() ?? client.getAllAccounts()[0] ?? undefined;
}

/** Returns a Graph token without user interaction, or throws SignInRequiredError. */
export async function getGraphToken(): Promise<string> {
  const client = await getClient();
  try {
    const result = await client.acquireTokenSilent({ scopes: GRAPH_SCOPES, account: currentAccount(client) });
    return result.accessToken;
  } catch (error) {
    throw new SignInRequiredError(error);
  }
}

function describeSignInError(error: unknown): string {
  const code = error instanceof AuthError ? error.errorCode : "";
  switch (code) {
    case BrowserAuthErrorCodes.userCancelled:
      return "Sign-in was cancelled.";
    case BrowserAuthErrorCodes.popupWindowError:
    case BrowserAuthErrorCodes.emptyWindowError:
      return "The sign-in window could not be opened. Allow pop-ups for this add-in and try again.";
    case BrowserAuthErrorCodes.interactionInProgress:
      return "A sign-in window is already open. Finish signing in there, then try again.";
    default:
      return `Sign-in failed: ${error instanceof Error ? error.message : String(error)}`;
  }
}

/** Shows the sign-in (and, the first time, consent) window. Call it from a button click. */
export async function signIn(): Promise<void> {
  const client = await getClient();
  try {
    const result = await client.acquireTokenPopup({ scopes: GRAPH_SCOPES, loginHint: mailboxUser() });
    if (result.account) {
      client.setActiveAccount(result.account);
    }
  } catch (error) {
    throw new Error(describeSignInError(error));
  }
}
