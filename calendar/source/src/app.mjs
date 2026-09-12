import { createLocalEventStore } from "./local-event-store.mjs";
import { createCalendarController } from "./calendar-controller.mjs";
import { mountCalendarApp } from "./calendar-ui.mjs";
import { microsoftConfig } from "./config.mjs";
import { initializeOutlookAuth } from "./outlook-auth.mjs";
import { createOutlookCalendarClient } from "./outlook-calendar-client.mjs";
import "./immersive.mjs";

const redirectUri = new URL(
  "redirect.html",
  new URL(import.meta.env.BASE_URL, window.location.origin),
).href;
const settingsKey = "outlook-calendar-poc.client-id";
let clientId = microsoftConfig.clientId;
try {
  clientId = sessionStorage.getItem(settingsKey) || clientId;
} catch {
  /* A configured file still works without browser settings storage. */
}
let authInstance = null,
  initializationError = null;
const ready = clientId
  ? initializeOutlookAuth({ ...microsoftConfig, clientId, redirectUri })
      .then((value) => {
        authInstance = value;
      })
      .catch((error) => {
        initializationError = error;
      })
  : Promise.resolve();
async function requireAuth() {
  await ready;
  if (initializationError) throw initializationError;
  if (!authInstance)
    throw new Error(
      "Add a Microsoft application Client ID to connect Outlook.",
    );
  return authInstance;
}
const auth = {
  async restore() {
    await ready;
    if (initializationError) throw initializationError;
    return authInstance ? authInstance.restore() : null;
  },
  async signIn() {
    return (await requireAuth()).signIn();
  },
  async signOut() {
    return (await requireAuth()).signOut();
  },
};
const outlookClient = createOutlookCalendarClient({
  getAccessToken: async () => (await requireAuth()).getAccessToken(),
});
const controller = createCalendarController({
  localStore: createLocalEventStore(),
  auth: clientId ? auth : null,
  outlookClient,
});
mountCalendarApp({
  controller,
  clientId,
  redirectUri,
  onConnect: () => controller.connectOutlook(),
  onConfigure: async (id) => {
    if (
      !/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i.test(
        id,
      )
    )
      throw new Error(
        "Enter the Application (client) ID from your Microsoft app registration.",
      );
    sessionStorage.setItem(settingsKey, id);
    window.location.reload();
  },
});
controller.start().catch((error) => controller.setStatus(error.message));
