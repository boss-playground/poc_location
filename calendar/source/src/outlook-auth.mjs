const interactionCodes = new Set([
  "interaction_required",
  "consent_required",
  "login_required",
  "no_tokens_found",
  "refresh_token_expired",
]);
function authError(error) {
  const code = error?.errorCode ?? "";
  if (code === "user_cancelled")
    return new Error(
      "Microsoft sign-in was cancelled. Your browser events are still available.",
    );
  if (code === "popup_window_error" || code === "empty_window_error")
    return new Error("Allow popups for this page, then connect Outlook again.");
  if (code === "interaction_in_progress")
    return new Error(
      "Finish the open Microsoft sign-in window first. If it was already closed, open this calendar's address in a new tab and connect again.",
    );
  if (interactionCodes.has(code))
    return new Error(
      "Your Microsoft session needs permission or renewal. Disconnect and connect Outlook again.",
    );
  return new Error(
    "Microsoft sign-in could not finish. Check your connection, Client ID and registered redirect URI, then try again.",
  );
}
export function createOutlookAuth({
  pca,
  scopes,
  postLogoutRedirectUri,
  isInteractionRequired = (error) => interactionCodes.has(error?.errorCode),
}) {
  let signingIn = false;
  return {
    async restore() {
      const active = pca.getActiveAccount();
      if (active) return active;
      const accounts = pca.getAllAccounts();
      if (accounts.length === 1) {
        pca.setActiveAccount(accounts[0]);
        return accounts[0];
      }
      return null;
    },
    async signIn() {
      if (signingIn)
        throw new Error("Microsoft sign-in is already in progress.");
      signingIn = true;
      try {
        const result = await pca.loginPopup({
          scopes,
          prompt: "select_account",
        });
        if (!result.account) throw new Error("No account");
        pca.setActiveAccount(result.account);
        return result.account;
      } catch (error) {
        throw authError(error);
      } finally {
        signingIn = false;
      }
    },
    async getAccessToken({ interactive = false } = {}) {
      const account = pca.getActiveAccount();
      if (!account) throw new Error("Connect Outlook to load your calendar.");
      const request = { scopes, account };
      let result;
      try {
        result = await pca.acquireTokenSilent(request);
      } catch (error) {
        if (!interactive || !isInteractionRequired(error))
          throw authError(error);
        try {
          result = await pca.acquireTokenPopup(request);
        } catch (popupError) {
          throw authError(popupError);
        }
      }
      if (pca.getActiveAccount()?.homeAccountId !== account.homeAccountId)
        throw new Error(
          "Microsoft account changed or disconnected during the request.",
        );
      if (!result.accessToken)
        throw new Error(
          "Microsoft did not return access to the calendar. Connect again.",
        );
      return result.accessToken;
    },
    async signOut() {
      const account = pca.getActiveAccount();
      pca.setActiveAccount(null);
      try {
        await pca.logoutPopup({
          account,
          postLogoutRedirectUri,
        });
      } catch {
        await pca.clearCache({ account });
      } finally {
        pca.setActiveAccount(null);
      }
    },
  };
}
export async function initializeOutlookAuth({
  clientId,
  redirectUri,
  authority,
  scopes,
}) {
  if (!clientId)
    throw new Error(
      "Add your Microsoft Application (client) ID before connecting Outlook.",
    );
  const msal = await import("@azure/msal-browser");
  const { PopupAwarePublicClientApplication, installDiscoveryDeadline } =
    await import("./outlook-popup-client.mjs");
  const pca = new PopupAwarePublicClientApplication({
    auth: {
      clientId,
      authority,
      redirectUri,
      navigateToLoginRequestUrl: false,
    },
    cache: { cacheLocation: "sessionStorage" },
    system: {
      loggerOptions: { piiLoggingEnabled: false, loggerCallback: () => {} },
    },
  });
  installDiscoveryDeadline(pca.getConfiguration().system.networkClient);
  await pca.initialize();
  return createOutlookAuth({
    pca,
    scopes,
    postLogoutRedirectUri: new URL("logout.html", redirectUri).href,
    isInteractionRequired: (error) =>
      error instanceof msal.InteractionRequiredAuthError ||
      interactionCodes.has(error?.errorCode),
  });
}
