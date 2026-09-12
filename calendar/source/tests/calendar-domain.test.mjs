import test from "node:test";
import assert from "node:assert/strict";
import {
  validateDraft,
  monthCells,
  monthRange,
  addDays,
  intersects,
  eventsOnDay,
  displayDate,
} from "../src/calendar-domain.mjs";

const draft = {
  source: "local",
  title: "  Planning  ",
  start: "2026-09-12T09:00",
  end: "2026-09-12T10:00",
  allDay: false,
};
test("draft validation normalizes valid timed input without shifting Bangkok wall time", () => {
  assert.deepEqual(validateDraft(draft), {
    ...draft,
    title: "Planning",
    start: "2026-09-12T09:00:00",
    end: "2026-09-12T10:00:00",
  });
});
test("invalid calendar dates, reversed times, blanks and invalid source cannot be saved", () => {
  for (const change of [
    { title: " " },
    { source: "bad" },
    { start: "2026-02-30T09:00" },
    { end: "2026-09-12T08:00" },
    { start: "" },
    { allDay: "true" },
    { start: "2026-09-12T25:00" },
    { end: "2026-09-12T09:00" },
  ]) {
    assert.throws(() => validateDraft({ ...draft, ...change }));
  }
});
test("all-day dates use exclusive end and midnight is not required in storage", () => {
  assert.equal(
    validateDraft({
      ...draft,
      allDay: true,
      start: "2026-09-12",
      end: "2026-09-13",
    }).end,
    "2026-09-13",
  );
  assert.throws(() =>
    validateDraft({
      ...draft,
      allDay: true,
      start: "2026-09-12",
      end: "2026-09-12",
    }),
  );
});
test("month grid starts Monday, covers adjacent dates, and handles leap years", () => {
  const cells = monthCells("2026-09-01");
  assert.equal(cells[0], "2026-08-31");
  assert.equal(cells.at(-1), "2026-10-04");
  assert.equal(cells.length, 35);
  assert.deepEqual(monthRange("2024-02-12"), {
    start: "2024-02-01T00:00:00",
    end: "2024-03-01T00:00:00",
  });
  assert.equal(addDays("2024-02-28", 1), "2024-02-29");
  assert.equal(displayDate(new Date("2026-09-11T18:00:00Z")), "2026-09-12");
});
test("day overlap excludes end boundary and includes multi-day timed events", () => {
  const allDay = {
    ...draft,
    allDay: true,
    start: "2026-09-12",
    end: "2026-09-14",
  };
  assert.equal(eventsOnDay([allDay], "2026-09-13").length, 1);
  assert.equal(eventsOnDay([allDay], "2026-09-14").length, 0);
  assert.equal(
    intersects(
      { ...draft, start: "2026-09-11T23:00:00", end: "2026-09-12T02:00:00" },
      { start: "2026-09-12T00:00:00", end: "2026-09-13T00:00:00" },
    ),
    true,
  );
});
