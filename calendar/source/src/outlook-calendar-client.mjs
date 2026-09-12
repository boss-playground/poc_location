import { validateDraft } from "./calendar-domain.mjs";
const BASE = "https://graph.microsoft.com/v1.0";
const ZONE = "SE Asia Standard Time";
function graphWallTime(value) {
  if (!value || typeof value.dateTime !== "string")
    throw new Error("Outlook returned an invalid date/time response.");
  const wall = value.dateTime.replace(/\.\d+$/, "").replace(/Z$/, "");
  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/.test(wall))
    throw new Error("Outlook returned an invalid date/time response.");
  const parsed = new Date(`${wall}Z`);
  if (!Number.isFinite(+parsed) || parsed.toISOString().slice(0, 19) !== wall)
    throw new Error("Outlook returned an invalid date/time response.");
  if (value.timeZone === "UTC" || value.timeZone === "Etc/UTC")
    return new Date(+parsed + 7 * 3600000).toISOString().slice(0, 19);
  if (value.timeZone === ZONE || value.timeZone === "Asia/Bangkok") return wall;
  throw new Error(
    "Outlook returned a different timezone. Refresh the calendar to request Bangkok time.",
  );
}
function dateOnly(value) {
  const date = value?.dateTime?.slice(0, 10);
  if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(date))
    throw new Error("Outlook returned an invalid all-day date.");
  return date;
}
export function mapGraphEvent(raw) {
  if (
    !raw ||
    typeof raw.id !== "string" ||
    !raw.id ||
    typeof raw.isAllDay !== "boolean"
  )
    throw new Error("Outlook returned an invalid event response.");
  const event = validateDraft({
    source: "outlook",
    title: raw.subject?.trim() || "(Untitled event)",
    start: raw.isAllDay ? dateOnly(raw.start) : graphWallTime(raw.start),
    end: raw.isAllDay ? dateOnly(raw.end) : graphWallTime(raw.end),
    allDay: raw.isAllDay,
  });
  return {
    ...event,
    id: raw.id,
    lastModified: raw.lastModifiedDateTime,
    readOnly: raw.type !== "singleInstance" && !!raw.type,
    isMeeting: Array.isArray(raw.attendees) && raw.attendees.length > 0,
    graphTimeZone: raw.isAllDay
      ? raw.originalStartTimeZone || raw.start.timeZone
      : ZONE,
  };
}
function safeUrl(value) {
  const url = new URL(value);
  if (
    url.protocol !== "https:" ||
    url.host !== "graph.microsoft.com" ||
    !url.pathname.startsWith("/v1.0/") ||
    url.username ||
    url.password
  )
    throw new Error("Outlook returned an unsafe pagination URL.");
  return url.href;
}
function zoneHeader(zone) {
  if (
    typeof zone !== "string" ||
    !zone ||
    zone.length > 100 ||
    /["\r\n]/.test(zone)
  )
    throw new Error("Outlook returned an invalid time zone.");
  return `outlook.timezone="${zone}"`;
}
function responseError(status) {
  if (status === 401)
    return new Error(
      "Outlook session expired (401). Disconnect and connect Outlook again.",
    );
  if (status === 403)
    return new Error(
      "Outlook permission denied (403). Grant Calendars.ReadWrite when connecting.",
    );
  if (status === 404)
    return new Error(
      "This Outlook event no longer exists (404). Refresh the calendar.",
    );
  if (status === 429)
    return new Error(
      "Outlook is limiting requests (429). Wait a moment before syncing again.",
    );
  return new Error(
    `Outlook request failed (HTTP ${status}). Check your connection and try again.`,
  );
}
export function createOutlookCalendarClient({
  getAccessToken,
  fetchImpl = globalThis.fetch,
  timeZone = ZONE,
}) {
  async function request(
    url,
    { method = "GET", body, zone = timeZone, expected = 200 } = {},
  ) {
    const target = safeUrl(url),
      prefer = zoneHeader(zone);
    const token = await getAccessToken();
    let response;
    try {
      response = await fetchImpl(target, {
        method,
        redirect: "error",
        signal: AbortSignal.timeout(25000),
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json",
          Prefer: prefer,
          ...(body ? { "Content-Type": "application/json" } : {}),
        },
        ...(body ? { body: JSON.stringify(body) } : {}),
      });
    } catch {
      throw new Error(
        "Could not reach Outlook. Check your connection, then sync again. If this happened while saving, check Outlook before retrying.",
      );
    }
    if (response.status !== expected) throw responseError(response.status);
    if (expected === 204) return null;
    try {
      return await response.json();
    } catch {
      throw new Error(
        "Outlook returned an unreadable response. Refresh before retrying a change.",
      );
    }
  }
  async function normalize(raw) {
    // Request the source zone for all-day items so timezone projection never shifts calendar dates.
    if (
      raw?.isAllDay &&
      raw.originalStartTimeZone &&
      raw.originalStartTimeZone !== timeZone
    ) {
      raw = await request(`${BASE}/me/events/${encodeURIComponent(raw.id)}`, {
        zone: raw.originalStartTimeZone,
      });
    }
    return mapGraphEvent(raw);
  }
  function payload(draft) {
    const valid = validateDraft(draft);
    if (valid.source !== "outlook")
      throw new Error("Choose Outlook as the event destination.");
    const zone =
      valid.allDay && draft.graphTimeZone ? draft.graphTimeZone : timeZone;
    zoneHeader(zone);
    return {
      subject: valid.title,
      isAllDay: valid.allDay,
      start: {
        dateTime: valid.allDay ? `${valid.start}T00:00:00` : valid.start,
        timeZone: zone,
      },
      end: {
        dateTime: valid.allDay ? `${valid.end}T00:00:00` : valid.end,
        timeZone: zone,
      },
    };
  }
  return {
    async list(range) {
      const url = new URL(`${BASE}/me/calendar/calendarView`);
      url.searchParams.set("startDateTime", `${range.start}+07:00`);
      url.searchParams.set("endDateTime", `${range.end}+07:00`);
      url.searchParams.set("$top", "250");
      let next = url.href;
      const events = [],
        visited = new Set();
      while (next) {
        const safe = safeUrl(next);
        if (visited.has(safe))
          throw new Error("Outlook returned a repeated pagination URL.");
        visited.add(safe);
        const page = await request(safe);
        if (!Array.isArray(page.value))
          throw new Error("Outlook returned an invalid calendar response.");
        for (const raw of page.value) {
          if (!raw.isCancelled) events.push(await normalize(raw));
        }
        next = page["@odata.nextLink"];
      }
      return events;
    },
    async create(draft) {
      const body = payload(draft);
      body.transactionId = crypto.randomUUID();
      return mapGraphEvent(
        await request(`${BASE}/me/events`, {
          method: "POST",
          body,
          expected: 201,
          zone: body.start.timeZone,
        }),
      );
    },
    async update(id, draft) {
      const body = payload(draft);
      return mapGraphEvent(
        await request(`${BASE}/me/events/${encodeURIComponent(id)}`, {
          method: "PATCH",
          body,
          zone: body.start.timeZone,
        }),
      );
    },
    async remove(id) {
      await request(`${BASE}/me/events/${encodeURIComponent(id)}`, {
        method: "DELETE",
        expected: 204,
      });
    },
  };
}
