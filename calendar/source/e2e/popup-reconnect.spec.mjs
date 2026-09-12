import { test, expect } from "@playwright/test";

const clientId = "00000000-0000-4000-8000-000000000001";
const authority = "https://login.microsoftonline.com/consumers";
const localKey = "outlook-calendar-poc.local-events.v1";

// Only the identity service is mocked: MSAL, popup windows, storage and the
// application's auth module all run unchanged in an isolated browser context.
async function setup({ page, context, authority: testAuthority = authority }) {
  const authorityHost = new URL(testAuthority).hostname;
  // Keep this auth-only harness independent of concurrent UI edits/Vite HMR.
  await page.route("**/calendar/", (route) =>
    route.fulfill({
      contentType: "text/html",
      body: "<!doctype html><title>Popup auth regression</title><body></body>",
    }),
  );
  let authorization;
  await context.route("https://**/*", (route) => {
    const url = new URL(route.request().url());
    if (!["login.microsoftonline.com", authorityHost].includes(url.hostname))
      return route.abort();
    if (url.pathname.includes("/discovery/instance"))
      return route.fulfill({
        json: {
          tenant_discovery_endpoint: `${testAuthority}/v2.0/.well-known/openid-configuration`,
          metadata: [
            {
              preferred_network: authorityHost,
              preferred_cache: "login.windows.net",
              aliases: [
                authorityHost,
                "login.microsoftonline.com",
                "login.windows.net",
                "login.microsoft.com",
                "sts.windows.net",
              ],
            },
          ],
        },
      });
    if (url.pathname.endsWith("/.well-known/openid-configuration"))
      return route.fulfill({
        json: {
          authorization_endpoint: `${testAuthority}/oauth2/v2.0/authorize`,
          token_endpoint: `${testAuthority}/oauth2/v2.0/token`,
          end_session_endpoint: `${testAuthority}/oauth2/v2.0/logout`,
          issuer: `${testAuthority}/v2.0`,
          jwks_uri: `${testAuthority}/discovery/v2.0/keys`,
        },
      });
    if (url.pathname.endsWith("/authorize")) {
      authorization = url;
      return route.fulfill({
        contentType: "text/html",
        body: "<title>Mock Microsoft sign-in</title>Close this window to cancel.",
      });
    }
    if (url.pathname.endsWith("/token")) {
      const encode = (value) =>
        Buffer.from(JSON.stringify(value)).toString("base64url");
      const now = Math.floor(Date.now() / 1000);
      return route.fulfill({
        json: {
          token_type: "Bearer",
          scope: "openid profile User.Read Calendars.ReadWrite",
          expires_in: 3600,
          ext_expires_in: 3600,
          access_token: "synthetic-calendar-token",
          refresh_token: "synthetic-refresh-token",
          client_info: encode({ uid: "popup-user", utid: "popup-tenant" }),
          id_token: `${encode({ alg: "none", typ: "JWT" })}.${encode({
            aud: clientId,
            iss: `${authority}/v2.0`,
            iat: now,
            nbf: now,
            exp: now + 3600,
            nonce: authorization.searchParams.get("nonce"),
            sub: "popup-user",
            oid: "popup-user",
            tid: "popup-tenant",
            name: "Popup Test",
            preferred_username: "popup@example.test",
          })}.synthetic`,
        },
      });
    }
    return route.abort();
  });
  await page.goto("/calendar/");
  await page.evaluate(
    async ({ clientId, authority, localKey }) => {
      const { initializeOutlookAuth } =
        await import("/calendar/src/outlook-auth.mjs");
      const options = {
        clientId,
        authority,
        redirectUri: `${location.origin}/calendar/redirect.html`,
        scopes: ["User.Read", "Calendars.ReadWrite"],
      };
      localStorage.setItem(localKey, '[{"id":"keep-local-event"}]');
      window.popupResults = [];
      window.makePopupAuth = () => initializeOutlookAuth(options);
      window.popupAuth = await window.makePopupAuth();
      const button = document.createElement("button");
      button.textContent = "Test Connect";
      button.onclick = () =>
        window.popupAuth.signIn().then(
          (account) => window.popupResults.push({ account }),
          (error) => window.popupResults.push({ error: error.message }),
        );
      document.body.append(button);
    },
    { clientId, authority: testAuthority, localKey },
  );
}

async function openPopup(page) {
  const opened = page.waitForEvent("popup");
  await page.getByRole("button", { name: "Test Connect", exact: true }).click();
  const popup = await opened;
  await expect(popup).toHaveTitle("Mock Microsoft sign-in");
  return popup;
}

// Removing close detection must leave this pending and prevent reconnect.
test("closing the real MSAL popup cancels promptly and Connect opens a fresh popup", async ({
  page,
  context,
}) => {
  await setup({ page, context });
  const popup = await openPopup(page);
  await popup.close();
  await expect
    .poll(() => page.evaluate(() => window.popupResults), { timeout: 4000 })
    .toEqual([{ error: expect.stringMatching(/cancel/i) }]);
  const second = await openPopup(page);
  await second.close();
  await expect
    .poll(() => page.evaluate(() => window.popupResults.length))
    .toBe(2);
  expect(
    await page.evaluate((key) => localStorage.getItem(key), localKey),
  ).toBe('[{"id":"keep-local-event"}]');
});

// Overriding the shared interaction lock would cancel/steal the first popup.
test("another auth instance cannot replace an active MSAL popup", async ({
  page,
  context,
}) => {
  await setup({ page, context });
  const popup = await openPopup(page);
  await page.evaluate(async () => {
    window.popupAuth = await window.makePopupAuth();
  });
  await page.getByRole("button", { name: "Test Connect", exact: true }).click();
  await expect
    .poll(() => page.evaluate(() => window.popupResults))
    .toEqual([{ error: expect.stringMatching(/progress|open|sign-in/i) }]);
  expect(popup.isClosed()).toBe(false);
  await popup.close();
  await expect
    .poll(() => page.evaluate(() => window.popupResults.length), {
      timeout: 4000,
    })
    .toBe(2);
  const next = await openPopup(page);
  await next.close();
});

// A valid bridge message followed by close must win over cancellation. Losing
// that race breaks real login; clearing caches on cancel breaks silent tokens.
test("a real redirect response signs in and later cancellation preserves accounts and tokens", async ({
  page,
  context,
}) => {
  await setup({ page, context });
  const popup = await openPopup(page);
  const url = new URL(popup.url());
  const landing = new URL(url.searchParams.get("redirect_uri"));
  landing.hash = new URLSearchParams({
    code: "synthetic-auth-code",
    state: url.searchParams.get("state"),
  }).toString();
  await popup.goto(landing.href).catch((error) => {
    if (!popup.isClosed()) throw error;
  });
  await expect
    .poll(() => page.evaluate(() => window.popupResults[0]?.account?.username))
    .toBe("popup@example.test");
  await expect.poll(() => popup.isClosed()).toBe(true);
  const before = await page.evaluate(async () => ({
    account: await window.popupAuth.restore(),
    local: { ...localStorage },
    token: await window.popupAuth.getAccessToken(),
  }));
  expect(before.token).toBe("synthetic-calendar-token");
  const cancelled = await openPopup(page);
  await cancelled.close();
  await expect
    .poll(() => page.evaluate(() => window.popupResults[1]?.error))
    .toMatch(/cancel/i);
  expect(
    await page.evaluate(async () => ({
      account: await window.popupAuth.restore(),
      local: { ...localStorage },
      token: await window.popupAuth.getAccessToken(),
    })),
  ).toEqual(before);
});

// A lock alone has no popup identity or lifetime. Never assume it is stale,
// including same-client locks surviving reload and other-client interactions.
for (const owner of [clientId, "00000000-0000-4000-8000-000000000002"]) {
  test(`an unowned lock (${owner}) stays intact and offers safe recovery`, async ({
    page,
    context,
  }) => {
    await setup({ page, context });
    const key = "msal.interaction.status";
    const value = JSON.stringify({ clientId: owner, type: "signin" });
    await page.evaluate(
      ({ key, value }) => sessionStorage.setItem(key, value),
      { key, value },
    );
    await page
      .getByRole("button", { name: "Test Connect", exact: true })
      .click();
    await expect
      .poll(() => page.evaluate(() => window.popupResults[0]?.error))
      .toMatch(/new tab/i);
    expect(await page.evaluate((key) => sessionStorage.getItem(key), key)).toBe(
      value,
    );
    expect(context.pages()).toHaveLength(1);
  });
}

test("MSAL timeout releases the lock and permits a fresh Connect", async ({
  page,
  context,
}) => {
  await page.clock.install();
  await setup({ page, context });
  const popup = await openPopup(page);
  await page.clock.runFor(61000);
  await expect
    .poll(() => page.evaluate(() => window.popupResults[0]?.error))
    .toMatch(/could not finish/i);
  await expect.poll(() => popup.isClosed()).toBe(true);
  const next = await openPopup(page);
  await next.close();
  await page.clock.runFor(1000);
  await expect
    .poll(() => page.evaluate(() => window.popupResults.length))
    .toBe(2);
});

test("a lock orphaned by reload recovers in a fresh tab without deleting the old session", async ({
  page,
  context,
}) => {
  await setup({ page, context });
  const abandoned = await openPopup(page);
  // Replacing the document destroys the pending promise but MSAL's session
  // lock survives. Its record cannot tell us whether a popup is still active.
  await setup({ page, context });
  await abandoned.close();
  await page.getByRole("button", { name: "Test Connect", exact: true }).click();
  await expect
    .poll(() => page.evaluate(() => window.popupResults[0]?.error))
    .toMatch(/new tab/i);
  const before = await page.evaluate(() => ({ ...sessionStorage }));
  const recovery = await context.newPage();
  await setup({ page: recovery, context });
  const popup = await openPopup(recovery);
  await popup.close();
  await expect
    .poll(() => recovery.evaluate(() => window.popupResults[0]?.error))
    .toMatch(/cancel/i);
  expect(await page.evaluate(() => ({ ...sessionStorage }))).toEqual(before);
  expect(
    await recovery.evaluate((key) => localStorage.getItem(key), localKey),
  ).toBe('[{"id":"keep-local-event"}]');
});

test("closing a blank popup during stalled discovery settles before retry and late discovery cannot unlock the next login", async ({
  page,
  context,
}) => {
  await page.clock.install();
  // The configured Microsoft host has baked-in metadata in 5.21.0. Exercise
  // actual network discovery using a synthetic authority without baked-in data.
  await setup({
    page,
    context,
    authority: "https://login.example.test/consumers",
  });
  const held = [];
  let released = 0;
  let stall = true;
  await context.route(
    /https:\/\/.*\/(?:.*discovery\/instance|.*\.well-known\/openid-configuration)/,
    async (route) => {
      const delayed = stall;
      if (delayed) await new Promise((resolve) => held.push(resolve));
      await route.fallback();
      if (delayed) released++;
    },
  );
  const opened = page.waitForEvent("popup");
  await page.getByRole("button", { name: "Test Connect", exact: true }).click();
  const blank = await opened;
  await expect.poll(() => held.length).toBeGreaterThan(0);
  expect(blank.url()).toBe("about:blank");
  await blank.close();
  // Two sequential metadata GETs may be needed. Neither may wait forever.
  await page.clock.runFor(21000);
  await expect
    .poll(() => page.evaluate(() => window.popupResults.length), {
      timeout: 2000,
    })
    .toBe(1);
  expect(
    await page.evaluate(() =>
      sessionStorage.getItem("msal.interaction.status"),
    ),
  ).toBeNull();

  stall = false;
  const next = await openPopup(page);
  const lock = await page.evaluate(() =>
    sessionStorage.getItem("msal.interaction.status"),
  );
  expect(lock).not.toBeNull();
  for (const release of held) release();
  await expect.poll(() => released).toBeGreaterThanOrEqual(held.length);
  await page.clock.runFor(1000);
  expect(
    await page.evaluate(() =>
      sessionStorage.getItem("msal.interaction.status"),
    ),
  ).toBe(lock);
  expect(await page.evaluate(() => window.popupResults.length)).toBe(1);
  expect(next.isClosed()).toBe(false);
  await next.close();
  await page.clock.runFor(1000);
  await expect
    .poll(() => page.evaluate(() => window.popupResults.length))
    .toBe(2);
});
