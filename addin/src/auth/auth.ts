import { broadcastResponseToMainFrame } from "@azure/msal-browser/redirect-bridge";

// Redirect page for the MSAL popup fallback: hands the sign-in response back to the task pane.
broadcastResponseToMainFrame().catch((error: unknown) => {
  document.body.textContent = `Sign-in could not be completed: ${error instanceof Error ? error.message : String(error)}`;
});
