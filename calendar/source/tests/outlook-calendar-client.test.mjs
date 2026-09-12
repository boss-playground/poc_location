import test from "node:test";
test("successful all-day write in a source timezone does not depend on a second fetch", async () => {
  const raw = {
    id: "a",
    subject: "Holiday",
    isAllDay: true,
    type: "singleInstance",
    originalStartTimeZone: "Line Islands Standard Time",
    start: {
      dateTime: "2026-09-12T00:00:00",
      timeZone: "Line Islands Standard Time",
    },
    end: {
      dateTime: "2026-09-13T00:00:00",
      timeZone: "Line Islands Standard Time",
    },
  };
  const { client, calls } = fixture([json(raw)]);
  const saved = await client.update("a", {
    source: "outlook",
    title: "Holiday",
    allDay: true,
    start: "2026-09-12",
    end: "2026-09-13",
    graphTimeZone: "Line Islands Standard Time",
  });
  assert.equal(saved.start, "2026-09-12");
  assert.equal(calls.length, 1);
  assert.equal(
    calls[0].init.headers.Prefer,
    'outlook.timezone="Line Islands Standard Time"',
  );
});
import assert from "node:assert/strict";
import {
  createOutlookCalendarClient,
  mapGraphEvent,
} from "../src/outlook-calendar-client.mjs";
const graph = (id = "a", overrides = {}) => ({
  id,
  subject: "Dentist",
  start: {
    dateTime: "2026-09-12T09:00:00.0000000",
    timeZone: "SE Asia Standard Time",
  },
  end: {
    dateTime: "2026-09-12T10:00:00.0000000",
    timeZone: "SE Asia Standard Time",
  },
  isAllDay: false,
  type: "singleInstance",
  attendees: [],
  isOrganizer: true,
  ...overrides,
});
function fixture(responses) {
  const calls = [];
  const client = createOutlookCalendarClient({
    getAccessToken: async () => "test-token",
    fetchImpl: async (url, init) => {
      calls.push({ url, init });
      const response = responses.shift();
      if (response instanceof Error) throw response;
      assert.ok(response, "Unexpected request");
      return response;
    },
  });
  return { client, calls };
}
const json = (data, status = 200) =>
  new Response(JSON.stringify(data), {
    status,
    headers: { "content-type": "application/json" },
  });
const draft = {
  source: "outlook",
  title: "Plan",
  start: "2026-09-12T09:00:00",
  end: "2026-09-12T10:00:00",
  allDay: false,
};
test("calendarView is offset-bounded and follows pagination with authorization", async () => {
  const { client, calls } = fixture([
    json({
      value: [graph("a")],
      "@odata.nextLink":
        "https://graph.microsoft.com/v1.0/me/calendar/calendarView?$skiptoken=x",
    }),
    json({ value: [graph("b")] }),
  ]);
  const events = await client.list({
    start: "2026-09-01T00:00:00",
    end: "2026-10-01T00:00:00",
  });
  assert.deepEqual(
    events.map((e) => e.id),
    ["a", "b"],
  );
  const url = new URL(calls[0].url);
  assert.equal(
    url.searchParams.get("startDateTime"),
    "2026-09-01T00:00:00+07:00",
  );
  assert.equal(calls[0].init.headers.Authorization, "Bearer test-token");
  assert.equal(calls[0].init.redirect, "error");
});
test("unsafe pagination never forwards bearer tokens", async () => {
  for (const next of [
    "https://evil.example/steal",
    "http://graph.microsoft.com/v1.0/me",
    "https://graph.microsoft.com.evil.example/v1.0/me",
  ]) {
    const { client, calls } = fixture([
      json({ value: [], "@odata.nextLink": next }),
    ]);
    await assert.rejects(
      client.list({ start: "2026-09-01T00:00:00", end: "2026-10-01T00:00:00" }),
      /pagination|URL/i,
    );
    assert.equal(calls.length, 1);
  }
});
test("Graph mapping preserves time semantics, recurrence protection and meeting metadata", () => {
  const event = mapGraphEvent(
    graph("a", {
      type: "occurrence",
      attendees: [{ emailAddress: { address: "test@example.com" } }],
    }),
  );
  assert.equal(event.start, "2026-09-12T09:00:00");
  assert.equal(event.readOnly, true);
  assert.equal(event.isMeeting, true);
  assert.equal(
    mapGraphEvent(
      graph("u", {
        start: { dateTime: "2026-09-11T23:00:00", timeZone: "UTC" },
        end: { dateTime: "2026-09-12T00:00:00", timeZone: "UTC" },
      }),
    ).start,
    "2026-09-12T06:00:00",
  );
  assert.throws(
    () =>
      mapGraphEvent(
        graph("x", { start: { dateTime: "bogus", timeZone: "UTC" } }),
      ),
    /date|time|response/i,
  );
});
test("all-day creates midnight Graph payload with exclusive end and encoded update/delete IDs", async () => {
  const raw = graph("a/b", {
    isAllDay: true,
    start: {
      dateTime: "2026-09-12T00:00:00",
      timeZone: "SE Asia Standard Time",
    },
    end: { dateTime: "2026-09-13T00:00:00", timeZone: "SE Asia Standard Time" },
  });
  const { client, calls } = fixture([
    json(raw, 201),
    json(raw),
    new Response(null, { status: 204 }),
  ]);
  const allDay = {
    ...draft,
    allDay: true,
    start: "2026-09-12",
    end: "2026-09-13",
  };
  assert.equal((await client.create(allDay)).end, "2026-09-13");
  await client.update("a/b", allDay);
  await client.remove("a/b");
  const body = JSON.parse(calls[0].init.body);
  assert.equal(body.start.dateTime, "2026-09-12T00:00:00");
  assert.equal(body.end.dateTime, "2026-09-13T00:00:00");
  assert.equal(Object.hasOwn(body, "attendees"), false);
  assert.equal(
    calls[1].url,
    "https://graph.microsoft.com/v1.0/me/events/a%2Fb",
  );
  assert.equal(calls[1].init.method, "PATCH");
  assert.equal(calls[2].init.method, "DELETE");
});
test("all-day events in another source zone use original calendar dates", async () => {
  const projected = graph("all", {
    isAllDay: true,
    originalStartTimeZone: "Line Islands Standard Time",
    start: {
      dateTime: "2026-09-11T17:00:00",
      timeZone: "SE Asia Standard Time",
    },
    end: { dateTime: "2026-09-12T17:00:00", timeZone: "SE Asia Standard Time" },
  });
  const original = {
    ...projected,
    start: {
      dateTime: "2026-09-12T00:00:00",
      timeZone: "Line Islands Standard Time",
    },
    end: {
      dateTime: "2026-09-13T00:00:00",
      timeZone: "Line Islands Standard Time",
    },
  };
  const { client, calls } = fixture([
    json({ value: [projected] }),
    json(original),
  ]);
  const [event] = await client.list({
    start: "2026-09-01T00:00:00",
    end: "2026-10-01T00:00:00",
  });
  assert.equal(event.start, "2026-09-12");
  assert.equal(event.graphTimeZone, "Line Islands Standard Time");
  assert.match(calls[1].init.headers.Prefer, /Line Islands/);
});
test("failed writes surface an actionable error without retry and reject local sources before networking", async () => {
  const { client, calls } = fixture([
    json({ error: { message: "private details" } }, 403),
  ]);
  await assert.rejects(client.create(draft), /permission|403/i);
  assert.equal(calls.length, 1);
  await assert.rejects(
    client.create({ ...draft, source: "local" }),
    /Outlook|destination/i,
  );
  assert.equal(calls.length, 1);
});
