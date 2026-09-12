import { test, expect } from "@playwright/test";
test("local event CRUD, reload persistence and source filtering", async ({
  page,
}) => {
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto("/calendar/");
  await expect(page).toHaveTitle("My calendar");
  await expect(
    page.getByRole("heading", { name: "My calendar", exact: true }),
  ).toBeVisible();
  await page.getByRole("button", { name: "New event", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "New event" });
  await dialog.getByLabel("Title", { exact: true }).fill("POC local test");
  await dialog.getByRole("button", { name: "Save event", exact: true }).click();
  await expect(dialog).not.toBeVisible();
  await expect(
    page.getByRole("button", { name: /This calendar: POC local test/ }),
  ).toBeVisible();
  await page.reload();
  await expect(
    page.getByRole("button", { name: /This calendar: POC local test/ }),
  ).toBeVisible();
  await page.getByLabel("This calendar", { exact: true }).uncheck();
  await expect(
    page.getByRole("button", { name: /This calendar: POC local test/ }),
  ).toHaveCount(0);
  await page.getByLabel("This calendar", { exact: true }).check();
  await page
    .getByRole("button", { name: /This calendar: POC local test/ })
    .click();
  await page
    .getByRole("dialog")
    .getByLabel("Title", { exact: true })
    .fill("Updated local event");
  await page.getByRole("button", { name: "Save event", exact: true }).click();
  await page
    .getByRole("button", { name: /This calendar: Updated local event/ })
    .click();
  await page.getByRole("button", { name: "Delete event", exact: true }).click();
  await page
    .getByRole("button", { name: "Confirm delete", exact: true })
    .click();
  await expect(
    page.getByRole("button", { name: /This calendar: Updated local event/ }),
  ).toHaveCount(0);
  expect(errors).toEqual([]);
});
test("all-day editor uses inclusive end date while persistence uses exclusive end", async ({
  page,
}) => {
  await page.goto("/calendar/");
  await page.getByRole("button", { name: "New event", exact: true }).click();
  await page.getByLabel("Title", { exact: true }).fill("Day off");
  await page.getByLabel("All day", { exact: true }).check();
  const start = await page
    .getByLabel("Start date", { exact: true })
    .inputValue();
  expect(await page.getByLabel("End date", { exact: true }).inputValue()).toBe(
    start,
  );
  await page.getByRole("button", { name: "Save event", exact: true }).click();
  await expect(
    page.getByRole("button", { name: /This calendar: Day off/ }),
  ).toBeVisible();
  const stored = await page.evaluate(() =>
    JSON.parse(localStorage.getItem("outlook-calendar-poc.local-events.v1")),
  );
  expect(stored[0].end > stored[0].start).toBe(true);
});
test("Microsoft setup explains missing Client ID and mobile has no horizontal overflow", async ({
  page,
}) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto("/calendar/");
  await page
    .getByRole("button", { name: "Connect Outlook", exact: true })
    .click();
  await expect(
    page.getByRole("dialog", { name: "Connect your Outlook calendar" }),
  ).toBeVisible();
  await expect(page.getByLabel("Application (client) ID")).toBeVisible();
  await expect(page.locator("#redirect-uri")).toHaveText(
    "http://localhost:5173/calendar/redirect.html",
  );
  await expect(page.locator("#logout-uri")).toHaveText(
    "http://localhost:5173/calendar/logout.html",
  );
  await page.getByRole("button", { name: "Close setup", exact: true }).click();
  expect(
    await page.evaluate(
      () => document.documentElement.scrollWidth <= window.innerWidth,
    ),
  ).toBe(true);
  await page.getByRole("button", { name: "New event", exact: true }).click();
  await expect(
    page.getByRole("button", { name: "Save event", exact: true }),
  ).toBeVisible();
});
