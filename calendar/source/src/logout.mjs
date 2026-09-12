import { broadcastResponseToMainFrame } from "@azure/msal-browser/redirect-bridge";

// MSAL v5 logout popups also complete via the state-based redirect bridge.
// Do not mount the calendar or start a new authentication flow on this page.
broadcastResponseToMainFrame().catch(() => {
  document.getElementById("message").textContent =
    "You can close this window and return to your calendar.";
});
