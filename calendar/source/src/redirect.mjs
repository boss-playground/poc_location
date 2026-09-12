import { broadcastResponseToMainFrame } from "@azure/msal-browser/redirect-bridge";
broadcastResponseToMainFrame().catch(() => {
  document.getElementById("message").textContent =
    "No active sign-in response. Close this window and connect Outlook from your calendar.";
});
