import { test, expect } from "@playwright/test";

const clientId = "00000000-0000-4000-8000-000000000001";
const localKey = "outlook-calendar-poc.local-events.v1";
const clientKey = "outlook-calendar-poc.client-id";
const appointment = {
  id: "seconds-event",
  source: "local",
  title: "Precise appointment",
  start: "2026-09-12T09:00:10",
  end: "2026-09-12T09:00:50",
  allDay: false,
};

for (const source of ["local", "outlook"]) {
  test(`title-only ${source} edit preserves seconds including sub-minute events`, async ({
    page,
  }) => {
    await page.clock.install({ time: new Date("2026-09-12T03:00:00Z") });
    let written;
    if (source === "local") {
      await page.addInitScript(
        ({ key, event }) => localStorage.setItem(key, JSON.stringify([event])),
        { key: localKey, event: appointment },
      );
    } else {
      await page.addInitScript(
        ({ key, id }) => sessionStorage.setItem(key, id),
        { key: clientKey, id: clientId },
      );
      await page.route("**/src/outlook-auth.mjs", (route) =>
        route.fulfill({
          contentType: "application/javascript",
          body: `export async function initializeOutlookAuth(){return {restore:async()=>({homeAccountId:'test',username:'test@outlook.com'}),getAccessToken:async()=> 'synthetic-token'};}`,
        }),
      );
      let raw = {
        id: appointment.id,
        subject: appointment.title,
        isAllDay: false,
        type: "singleInstance",
        attendees: [],
        start: {
          dateTime: appointment.start,
          timeZone: "SE Asia Standard Time",
        },
        end: { dateTime: appointment.end, timeZone: "SE Asia Standard Time" },
      };
      await page.route("https://graph.microsoft.com/v1.0/**", (route) => {
        if (route.request().method() === "PATCH") {
          written = route.request().postDataJSON();
          raw = { ...raw, ...written };
          return route.fulfill({ json: raw });
        }
        return route.fulfill({ json: { value: [raw] } });
      });
    }
    await page.goto("/calendar/");
    await page.getByRole("button", { name: /Precise appointment/ }).click();
    await page.getByLabel("Title", { exact: true }).fill("Title changed only");
    await page.getByRole("button", { name: "Save event", exact: true }).click();
    await expect(
      page.getByRole("dialog", { name: "Edit event" }),
    ).not.toBeVisible();
    if (source === "local") {
      const [saved] = await page.evaluate(
        (key) => JSON.parse(localStorage.getItem(key)),
        localKey,
      );
      expect(saved.start).toBe("2026-09-12T09:00:10");
      expect(saved.end).toBe("2026-09-12T09:00:50");
    } else {
      expect(written.start.dateTime).toBe("2026-09-12T09:00:10");
      expect(written.end.dateTime).toBe("2026-09-12T09:00:50");
    }
    await page.getByRole("button", { name: /Title changed only/ }).click();
    await page.getByLabel("Start time", { exact: true }).fill("09:00:15");
    await page.getByRole("button", { name: "Save event", exact: true }).click();
    await expect(
      page.getByRole("dialog", { name: "Edit event" }),
    ).not.toBeVisible();
    if (source === "local") {
      const [saved] = await page.evaluate(
        (key) => JSON.parse(localStorage.getItem(key)),
        localKey,
      );
      expect(saved.start).toBe("2026-09-12T09:00:15");
    } else expect(written.start.dateTime).toBe("2026-09-12T09:00:15");
  });
}

test("storage getter denial mounts calendar and presents an actionable error", async ({
  page,
}) => {
  const errors = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.clock.install({ time: new Date("2026-09-12T03:00:00Z") });
  await page.addInitScript(() =>
    Object.defineProperty(window, "localStorage", {
      get() {
        throw new DOMException("Blocked by browser policy", "SecurityError");
      },
    }),
  );
  await page.goto("/calendar/");
  await expect(page.locator("#month-title")).toHaveText("September 2026");
  await expect(page.locator("#error-banner")).toContainText(
    "Could not read stored events",
  );
  await page.getByRole("button", { name: "New event", exact: true }).click();
  await page
    .getByLabel("Title", { exact: true })
    .fill("Do not claim this was saved");
  await page.getByRole("button", { name: "Save event", exact: true }).click();
  await expect(page.locator("#form-error")).toContainText(
    "Could not read stored events",
  );
  expect(errors).toEqual([]);
});

test("real MSAL logout completes through dedicated landing and preserves browser data", async ({
  page,
  context,
}) => {
  const requests = [];
  context.on("request", (request) => requests.push(request.url()));
  const authority = "https://login.microsoftonline.com/consumers";
  let logoutReturn;
  await context.route("https://login.microsoftonline.com/**", async (route) => {
    const url = new URL(route.request().url());
    if (url.pathname.endsWith("/logout")) {
      logoutReturn = url.searchParams.get("post_logout_redirect_uri");
      const destination = new URL(logoutReturn);
      destination.searchParams.set("state", url.searchParams.get("state"));
      return route.fulfill({
        contentType: "text/html",
        body: `<script>location.replace(${JSON.stringify(destination.href)})</script>`,
      });
    }
    if (url.pathname.includes("/discovery/instance"))
      return route.fulfill({
        json: {
          tenant_discovery_endpoint: `${authority}/v2.0/.well-known/openid-configuration`,
          metadata: [
            {
              preferred_network: "login.microsoftonline.com",
              preferred_cache: "login.windows.net",
              aliases: [
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
          authorization_endpoint: `${authority}/oauth2/v2.0/authorize`,
          token_endpoint: `${authority}/oauth2/v2.0/token`,
          end_session_endpoint: `${authority}/oauth2/v2.0/logout`,
          issuer: "https://login.microsoftonline.com/{tenantid}/v2.0",
          jwks_uri: `${authority}/discovery/v2.0/keys`,
        },
      });
    return route.abort();
  });
  await page.goto("/calendar/");
  await page.evaluate(
    async ({ id, event, localKey, clientKey, authority }) => {
      localStorage.setItem(localKey, JSON.stringify([event]));
      sessionStorage.setItem(clientKey, id);
      const { initializeOutlookAuth } =
        await import("/calendar/src/outlook-auth.mjs");
      const auth = await initializeOutlookAuth({
        clientId: id,
        redirectUri: `${location.origin}/calendar/redirect.html`,
        authority,
        scopes: ["User.Read", "Calendars.ReadWrite"],
      });
      window.reviewPageMarker = "main page retained";
      const button = document.createElement("button");
      button.textContent = "Run real MSAL logout";
      button.onclick = () => {
        auth
          .signOut()
          .then(() => {
            window.reviewLogoutDone = true;
          })
          .catch((error) => {
            window.reviewLogoutError = error.message;
          });
      };
      document.body.append(button);
    },
    { id: clientId, event: appointment, localKey, clientKey, authority },
  );
  const popupPromise = page.waitForEvent("popup");
  await page
    .getByRole("button", { name: "Run real MSAL logout", exact: true })
    .click();
  const popup = await popupPromise;
  await expect
    .poll(() => logoutReturn)
    .toBe("http://localhost:5173/calendar/logout.html");
  await expect
    .poll(() => page.evaluate(() => window.reviewLogoutDone))
    .toBe(true);
  await expect.poll(() => popup.isClosed()).toBe(true);
  expect(await page.evaluate(() => window.reviewPageMarker)).toBe(
    "main page retained",
  );
  expect(
    await page.evaluate((key) => sessionStorage.getItem(key), clientKey),
  ).toBe(clientId);
  expect(
    await page.evaluate(
      (key) => JSON.parse(localStorage.getItem(key)),
      localKey,
    ),
  ).toEqual([appointment]);
  expect(
    requests.some(
      (url) =>
        url.includes("/src/redirect.mjs") ||
        /\/oauth2\/v2\.0\/(authorize|token)/.test(url),
    ),
  ).toBe(false);
});
