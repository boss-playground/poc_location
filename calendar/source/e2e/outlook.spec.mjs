import { test, expect } from "@playwright/test";
test.use({ hasTouch: true });
const raw = (id, subject, day) => ({
  id,
  subject,
  start: {
    dateTime: `2026-09-${day}T09:00:00`,
    timeZone: "SE Asia Standard Time",
  },
  end: {
    dateTime: `2026-09-${day}T10:00:00`,
    timeZone: "SE Asia Standard Time",
  },
  isAllDay: false,
  type: "singleInstance",
  attendees: [],
  isOrganizer: true,
});
async function mockMicrosoft(page) {
  const records = [raw("cloud-existing", "Outlook appointment", "14")];
  const writes = [];
  let deny = false;
  await page.clock.install({ time: new Date("2026-09-12T03:00:00Z") });
  await page.addInitScript(() =>
    sessionStorage.setItem(
      "outlook-calendar-poc.client-id",
      "00000000-0000-4000-8000-000000000001",
    ),
  );
  await page.route(/\/src\/outlook-auth\.mjs(?:\?.*)?$/, (route) =>
    route.fulfill({
      contentType: "application/javascript",
      body: `export async function initializeOutlookAuth(){let active=sessionStorage.getItem('test-connected')==='yes';const account={homeAccountId:'test-user',username:'test@outlook.com'};return {restore:async()=>active?account:null,signIn:async()=>{active=true;sessionStorage.setItem('test-connected','yes');return account;},signOut:async()=>{active=false;sessionStorage.removeItem('test-connected');},getAccessToken:async()=>{if(!active)throw Error('Not connected');return 'test-only-token';}};}`,
    }),
  );
  await page.route("https://graph.microsoft.com/v1.0/**", async (route) => {
    const req = route.request();
    const method = req.method();
    if (method === "OPTIONS")
      return route.fulfill({
        status: 204,
        headers: {
          "access-control-allow-origin": "*",
          "access-control-allow-headers": "*",
          "access-control-allow-methods": "GET,POST,PATCH,DELETE,OPTIONS",
        },
      });
    expect(req.headers().authorization).toBe("Bearer test-only-token");
    if (method === "GET") return route.fulfill({ json: { value: records } });
    const body = req.postDataJSON();
    writes.push({ method, body });
    if (deny)
      return route.fulfill({
        status: 403,
        json: { error: { code: "ErrorAccessDenied" } },
      });
    if (method === "POST") {
      const saved = { ...raw("created", "", "12"), ...body };
      records.push(saved);
      return route.fulfill({ status: 201, json: saved });
    }
    const id = decodeURIComponent(
        new URL(req.url()).pathname.split("/").at(-1),
      ),
      index = records.findIndex((e) => e.id === id);
    if (method === "PATCH") {
      records[index] = { ...records[index], ...body };
      return route.fulfill({ json: records[index] });
    }
    if (method === "DELETE") {
      records.splice(index, 1);
      return route.fulfill({ status: 204 });
    }
  });
  return { records, writes, deny: () => (deny = true) };
}
test("Microsoft connection merges sources and cloud CRUD uses real Graph client payloads", async ({
  page,
}) => {
  const { records, writes, deny } = await mockMicrosoft(page);
  await page.goto("/calendar/");
  await page.getByRole("button", { name: "New event", exact: true }).click();
  await page.getByLabel("Title", { exact: true }).fill("Local-only plan");
  await page.getByRole("button", { name: "Save event", exact: true }).click();
  expect(writes).toHaveLength(0);
  await page
    .getByRole("button", { name: "Connect Outlook", exact: true })
    .click();
  await expect(
    page.getByRole("button", { name: /Outlook: Outlook appointment/ }),
  ).toBeVisible();
  await expect(
    page.getByRole("button", { name: /This calendar: Local-only plan/ }),
  ).toBeVisible();
  await page.getByRole("button", { name: "New event", exact: true }).click();
  await page.getByLabel("Title", { exact: true }).fill("Cloud-only plan");
  await page.getByRole("radio", { name: "Outlook Sync with Outlook" }).check();
  await page.getByRole("button", { name: "Save event", exact: true }).click();
  await expect(
    page.getByRole("button", { name: /Outlook: Cloud-only plan/ }),
  ).toBeVisible();
  expect(records.some((e) => e.subject === "Cloud-only plan")).toBe(true);
  const local = await page.evaluate(() =>
    JSON.parse(localStorage.getItem("outlook-calendar-poc.local-events.v1")),
  );
  expect(local).toHaveLength(1);
  expect(local[0].title).toBe("Local-only plan");
  await page.getByRole("button", { name: /Outlook: Cloud-only plan/ }).click();
  await page.getByLabel("Title", { exact: true }).fill("Cloud updated");
  await page.getByRole("button", { name: "Save event", exact: true }).click();
  await page.getByRole("button", { name: /Outlook: Cloud updated/ }).click();
  await page.getByRole("button", { name: "Delete event", exact: true }).click();
  await page
    .getByRole("button", { name: "Confirm delete", exact: true })
    .click();
  await expect(
    page.getByRole("button", { name: /Outlook: Cloud updated/ }),
  ).toHaveCount(0);
  expect(writes.map((w) => w.method)).toEqual(["POST", "PATCH", "DELETE"]);
  deny();
  await page.getByRole("button", { name: "New event", exact: true }).click();
  await page.getByLabel("Title", { exact: true }).fill("Keep this draft");
  await page.getByRole("radio", { name: "Outlook Sync with Outlook" }).check();
  await page.getByRole("button", { name: "Save event", exact: true }).click();
  await expect(
    page.getByRole("dialog", { name: "New event" }).getByRole("alert"),
  ).toContainText("403");
  await expect(page.getByLabel("Title", { exact: true })).toHaveValue(
    "Keep this draft",
  );
  await page.getByRole("button", { name: "Cancel", exact: true }).click();
  await page.getByRole("button", { name: "Disconnect", exact: true }).click();
  await expect(
    page.getByRole("button", { name: /Outlook: Outlook appointment/ }),
  ).toHaveCount(0);
  await expect(
    page.getByRole("button", { name: /This calendar: Local-only plan/ }),
  ).toBeVisible();
});
test("desktop and mobile visual evidence with explicitly synthetic events", async ({
  page,
}) => {
  await mockMicrosoft(page);
  await page.addInitScript(() => {
    sessionStorage.setItem("test-connected", "yes");
    localStorage.setItem(
      "outlook-calendar-poc.local-events.v1",
      JSON.stringify([
        {
          id: "a",
          source: "local",
          title: "Design review",
          start: "2026-09-07T14:30:00",
          end: "2026-09-07T15:30:00",
          allDay: false,
        },
        {
          id: "b",
          source: "local",
          title: "Ship calendar POC",
          start: "2026-09-15",
          end: "2026-09-16",
          allDay: true,
        },
        {
          id: "c",
          source: "local",
          title: "Weekend away",
          start: "2026-09-19",
          end: "2026-09-21",
          allDay: true,
        },
      ]),
    );
  });
  await page.setViewportSize({ width: 1536, height: 1024 });
  await page.emulateMedia({ reducedMotion: "reduce" });
  await page.goto("/calendar/");
  await expect(
    page.getByRole("button", { name: /Outlook: Outlook appointment/ }),
  ).toBeVisible();
  await page.screenshot({
    path: "/tmp/calendar-poc-desktop.png",
    fullPage: false,
  });
  await page.getByRole("button", { name: "New event", exact: true }).click();
  await page
    .getByLabel("Title", { exact: true })
    .fill("Make time for the next idea");
  await page.screenshot({
    path: "/tmp/calendar-poc-editor.png",
    fullPage: false,
  });
  await page.getByRole("button", { name: "Cancel", exact: true }).click();
  await page.setViewportSize({ width: 390, height: 844 });
  await expect(page.locator("#account-name")).toBeVisible();
  expect(
    await page.evaluate(
      () => document.documentElement.scrollWidth <= window.innerWidth,
    ),
  ).toBe(true);
  await page.screenshot({
    path: "/tmp/calendar-poc-mobile.png",
    fullPage: true,
  });
});
