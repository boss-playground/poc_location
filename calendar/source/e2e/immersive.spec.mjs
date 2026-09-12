import { test, expect } from "@playwright/test";

test("calendar uses the desktop viewport without a large empty footer and navigation still opens the editor", async ({
  page,
}) => {
  await page.emulateMedia({ reducedMotion: "reduce" });
  await page.setViewportSize({ width: 1536, height: 1024 });
  await page.goto("/calendar/");
  const main = await page.locator(".calendar-main").boundingBox();
  const height = await page.evaluate(
    () => document.documentElement.scrollHeight,
  );
  expect(height - (main.y + main.height)).toBeLessThanOrEqual(32);
  await expect(page.getByRole("button", { name: /^View Monday,/ })).toHaveCount(
    0,
  );
  const month = await page.locator("#month-title").textContent();
  await page.getByRole("button", { name: "Next month", exact: true }).click();
  await expect(page.locator("#month-title")).not.toHaveText(month);
  await page.getByRole("button", { name: "New event", exact: true }).click();
  await expect(
    page.getByRole("dialog", { name: "New event", exact: true }),
  ).toBeVisible();
});

test("ambient animation can be paused without blocking calendar navigation", async ({
  page,
}) => {
  await page.goto("/calendar/");
  const scene = page.locator(".ambient-scene");
  await expect(scene).toHaveCSS("animation-play-state", "running");
  await expect
    .poll(() =>
      page
        .locator(".ambient-landscape")
        .evaluate((image) => image.complete && image.naturalWidth > 0),
    )
    .toBe(true);
  await page
    .getByRole("button", { name: "Pause ambient animation", exact: true })
    .click();
  await expect(scene).toHaveCSS("animation-play-state", "paused");
  await expect(
    page.getByRole("button", { name: "Resume ambient animation", exact: true }),
  ).not.toHaveAttribute("aria-pressed");
  const month = await page.locator("#month-title").textContent();
  await page.getByRole("button", { name: "Next month", exact: true }).click();
  await expect(page.locator("#month-title")).not.toHaveText(month);
  await page
    .getByRole("button", { name: "Resume ambient animation", exact: true })
    .click();
  await expect(scene).toHaveCSS("animation-play-state", "running");
});

test("reduced motion stops ambient animation and leaves the event editor usable on mobile", async ({
  page,
}) => {
  await page.emulateMedia({ reducedMotion: "reduce" });
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto("/calendar/");
  await expect(page.locator(".ambient-scene")).toHaveCSS(
    "animation-name",
    "none",
  );
  await page.getByRole("button", { name: "New event", exact: true }).click();
  await page.getByLabel("Title", { exact: true }).fill("Night workspace test");
  await page.getByRole("button", { name: "Save event", exact: true }).click();
  await expect(
    page.getByRole("button", { name: /This calendar: Night workspace test/ }),
  ).toBeVisible();
  expect(
    await page.evaluate(
      () => document.documentElement.scrollWidth <= innerWidth,
    ),
  ).toBe(true);
});
