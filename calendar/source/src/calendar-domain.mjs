export const EVENT_SOURCES = Object.freeze({
  LOCAL: "local",
  OUTLOOK: "outlook",
});
export const TIME_ZONE = "Asia/Bangkok";
export function displayDate(now = new Date()) {
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: TIME_ZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(now);
}
const pad = (n) => String(n).padStart(2, "0");
function isDate(value) {
  if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(value))
    return false;
  const parsed = new Date(`${value}T00:00:00Z`);
  return (
    Number.isFinite(+parsed) &&
    parsed.toISOString().slice(0, 10) === value &&
    value >= "1900-01-01" &&
    value < "2200-01-01"
  );
}
function normalizeTime(value, allDay) {
  if (allDay) {
    if (!isDate(value))
      throw new Error("Choose a valid date between 1900 and 2199.");
    return value;
  }
  if (
    typeof value !== "string" ||
    !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2})?$/.test(value) ||
    !isDate(value.slice(0, 10))
  )
    throw new Error("Choose a valid date and time.");
  const [hour, minute, second = 0] = value.slice(11).split(":").map(Number);
  if (hour > 23 || minute > 59 || second > 59)
    throw new Error("Choose a valid time.");
  return value.length === 16 ? `${value}:00` : value;
}
export function validateDraft(draft) {
  if (!draft || !Object.values(EVENT_SOURCES).includes(draft.source))
    throw new Error("Choose an event destination.");
  if (typeof draft.title !== "string" || !draft.title.trim())
    throw new Error("Enter an event title.");
  if (draft.title.trim().length > 255)
    throw new Error("Use a title of 255 characters or fewer.");
  if (typeof draft.allDay !== "boolean")
    throw new Error("Choose an all-day or timed event.");
  const start = normalizeTime(draft.start, draft.allDay),
    end = normalizeTime(draft.end, draft.allDay);
  if (end <= start) throw new Error("End must be after start.");
  return {
    source: draft.source,
    title: draft.title.trim(),
    start,
    end,
    allDay: draft.allDay,
  };
}
export function addDays(date, days) {
  const result = new Date(`${date.slice(0, 10)}T12:00:00Z`);
  result.setUTCDate(result.getUTCDate() + days);
  return result.toISOString().slice(0, 10);
}
export function shiftMonth(date, count) {
  const result = new Date(`${date.slice(0, 7)}-01T12:00:00Z`);
  result.setUTCMonth(result.getUTCMonth() + count);
  return result.toISOString().slice(0, 10);
}
export function monthRange(date) {
  return {
    start: `${date.slice(0, 7)}-01T00:00:00`,
    end: `${shiftMonth(date, 1)}T00:00:00`,
  };
}
export function monthCells(date) {
  const first = `${date.slice(0, 7)}-01`;
  const offset = (new Date(`${first}T12:00:00Z`).getUTCDay() + 6) % 7;
  const last = addDays(shiftMonth(date, 1), -1);
  const count = Number(last.slice(8)) + offset;
  return Array.from({ length: Math.ceil(count / 7) * 7 }, (_, i) =>
    addDays(first, i - offset),
  );
}
export function gridRange(date) {
  const cells = monthCells(date);
  return {
    start: `${cells[0]}T00:00:00`,
    end: `${addDays(cells.at(-1), 1)}T00:00:00`,
  };
}
const midnight = (value) => (value.length === 10 ? `${value}T00:00:00` : value);
export function intersects(event, range) {
  return (
    midnight(event.start) < midnight(range.end) &&
    midnight(event.end) > midnight(range.start)
  );
}
export function eventsOnDay(events, date) {
  return events
    .filter((event) =>
      intersects(event, { start: date, end: addDays(date, 1) }),
    )
    .sort(
      (a, b) =>
        Number(b.allDay) - Number(a.allDay) ||
        a.start.localeCompare(b.start) ||
        a.title.localeCompare(b.title),
    );
}
export function monthLabel(date) {
  return new Intl.DateTimeFormat("en-US", {
    month: "long",
    year: "numeric",
    timeZone: "UTC",
  }).format(new Date(`${date.slice(0, 7)}-01T12:00:00Z`));
}
export function dayLabel(date) {
  return new Intl.DateTimeFormat("en-US", {
    weekday: "long",
    day: "numeric",
    month: "long",
    year: "numeric",
    timeZone: "UTC",
  }).format(new Date(`${date}T12:00:00Z`));
}
export function timeLabel(event) {
  return event.allDay
    ? "All day"
    : `${event.start.slice(11, 16)} – ${event.end.slice(11, 16)}`;
}
export function defaultDraft(date = displayDate()) {
  return {
    source: "local",
    title: "",
    start: `${date}T09:00`,
    end: `${date}T10:00`,
    allDay: false,
  };
}
