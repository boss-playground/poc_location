import { BrowserAuthError, PublicClientApplication } from "@azure/msal-browser";

// Install on this application's own default network client before initialize().
// Keep MSAL's POST/retry implementation and all non-discovery GETs intact.
export function installDiscoveryDeadline(networkClient) {
  const originalGet = networkClient.sendGetRequestAsync.bind(networkClient);
  networkClient.sendGetRequestAsync = async (url, options, timeout) => {
    const { pathname } = new URL(url);
    if (
      !pathname.endsWith("/.well-known/openid-configuration") &&
      !pathname.endsWith("/discovery/instance")
    )
      return originalGet(url, options, timeout);
    const controller = new AbortController();
    const deadline = setTimeout(() => controller.abort(), 10000);
    try {
      // Abort the actual fetch AND response body, not just an outer promise.
      // MSAL must unwind and release its own lock before signIn settles.
      const response = await fetch(url, {
        method: "GET",
        headers: new Headers(options?.headers),
        signal: controller.signal,
      });
      return {
        headers: Object.fromEntries(response.headers),
        status: response.status,
        body: await response.json(),
      };
    } finally {
      clearTimeout(deadline);
    }
  };
}

// MSAL 5.21.0's default popup bridge does not observe Window.closed. Use its
// response hook and request-specific channel; never override its shared lock
// or cancelPendingBridgeResponse (which can cancel another request's monitor).
export class PopupAwarePublicClientApplication extends PublicClientApplication {
  async waitForPopupResponse(request, popup, parent) {
    // MSAL library state is base64url JSON, followed by optional user state.
    const encoded = request.state
      .split("|")[0]
      .replace(/-/g, "+")
      .replace(/_/g, "/");
    const { id } = JSON.parse(atob(encoded));
    const cancellationChannel = new BroadcastChannel(id);
    let closedAt;
    let cancelled = false;
    const interval = setInterval(() => {
      if (!popup.closed) return;
      // The redirect bridge posts a response immediately before window.close().
      // Give that queued response time to arrive before reporting cancellation.
      closedAt ??= Date.now();
      if (Date.now() - closedAt < 250 || cancelled) return;
      cancelled = true;
      // MSAL has no request-scoped abort API. An empty response on this request's
      // random channel rejects its bridge and disposes its timer/channel. Await
      // that rejection so StandardController's finally releases its own lock.
      cancellationChannel.postMessage({ v: 1, payload: "" });
    }, 100);
    try {
      return await super.waitForPopupResponse(request, popup, parent);
    } catch (error) {
      if (cancelled && error.errorCode === "redirect_bridge_empty_response")
        throw new BrowserAuthError("user_cancelled", request.correlationId);
      throw error;
    } finally {
      clearInterval(interval);
      cancellationChannel.close();
    }
  }
}
