import test from "node:test";
import assert from "node:assert/strict";
import { createOutlookAuth } from "../src/outlook-auth.mjs";
function harness() {
  let active = null;
  const account = { homeAccountId: "a", username: "me@outlook.com" };
  const calls = [];
  const pca = {
    getActiveAccount: () => active,
    getAllAccounts: () => [],
    setActiveAccount: (a) => (active = a),
    loginPopup: async (request) => {
      calls.push(["login", request]);
      return { account };
    },
    acquireTokenSilent: async (request) => {
      calls.push(["silent", request]);
      return { accessToken: "token", account };
    },
    acquireTokenPopup: async (request) => {
      calls.push(["popup", request]);
      return { accessToken: "renewed", account };
    },
    logoutPopup: async () => {
      active = null;
    },
    clearCache: async () => {
      active = null;
    },
  };
  const auth = createOutlookAuth({
    pca,
    scopes: ["User.Read", "Calendars.ReadWrite"],
    postLogoutRedirectUri: "http://localhost:5173/calendar/logout.html",
  });
  return { auth, pca, calls, account };
}
test("logout uses its dedicated landing without navigating the main calendar", async () => {
  const { auth, pca, account } = harness();
  await auth.signIn();
  let logoutRequest;
  pca.logoutPopup = async (request) => {
    logoutRequest = request;
  };
  await auth.signOut();
  assert.equal(logoutRequest.account, account);
  assert.equal(
    logoutRequest.postLogoutRedirectUri,
    "http://localhost:5173/calendar/logout.html",
  );
  assert.equal(Object.hasOwn(logoutRequest, "mainWindowRedirectUri"), false);
  assert.equal(pca.getActiveAccount(), null);
});
test("sign-in sets a specific account used by silent Graph token requests", async () => {
  const { auth, pca, calls, account } = harness();
  assert.equal(await auth.restore(), null);
  await auth.signIn();
  assert.equal(pca.getActiveAccount(), account);
  assert.equal(await auth.getAccessToken(), "token");
  assert.equal(calls[1][1].account, account);
});
test("background token expiration requests reconnect without opening a popup", async () => {
  const { auth, pca, calls } = harness();
  await auth.signIn();
  pca.acquireTokenSilent = async () => {
    throw { errorCode: "interaction_required" };
  };
  await assert.rejects(auth.getAccessToken(), /reconnect|connect/i);
  assert.equal(
    calls.some((c) => c[0] === "popup"),
    false,
  );
  assert.equal(await auth.getAccessToken({ interactive: true }), "renewed");
});
test("network auth failures do not trigger an interactive prompt", async () => {
  const { auth, pca, calls } = harness();
  await auth.signIn();
  pca.acquireTokenSilent = async () => {
    throw { errorCode: "network_error" };
  };
  await assert.rejects(auth.getAccessToken({ interactive: true }));
  assert.equal(
    calls.some((c) => c[0] === "popup"),
    false,
  );
});
test("a token acquired for an account that was disconnected is rejected", async () => {
  const { auth, pca } = harness();
  await auth.signIn();
  let resolve;
  pca.acquireTokenSilent = () => new Promise((r) => (resolve = r));
  const pending = auth.getAccessToken();
  await auth.signOut();
  resolve({ accessToken: "stale" });
  await assert.rejects(pending, /account|disconnect/i);
});
test("cancelled login has a safe message and logout keeps account cleared even if popup fails", async () => {
  const { auth, pca } = harness();
  pca.loginPopup = async () => {
    throw { errorCode: "user_cancelled" };
  };
  await assert.rejects(auth.signIn(), /cancel/i);
  pca.logoutPopup = async () => {
    throw { errorCode: "popup_window_error" };
  };
  await auth.signOut();
  assert.equal(await auth.restore(), null);
});
