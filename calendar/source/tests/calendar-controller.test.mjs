import test from "node:test";
import assert from "node:assert/strict";
import { createCalendarController } from "../src/calendar-controller.mjs";
import { createLocalEventStore } from "../src/local-event-store.mjs";
const draft = {
  source: "local",
  title: "Browser plan",
  start: "2026-09-12T09:00:00",
  end: "2026-09-12T10:00:00",
  allDay: false,
};
const account = { homeAccountId: "me", username: "me@outlook.com" };
function setup(overrides = {}) {
  const db = new Map();
  const writes = [];
  let cloud = [];
  const localStore = createLocalEventStore(
    { getItem: (k) => db.get(k) ?? null, setItem: (k, v) => db.set(k, v) },
    () => "same-id",
  );
  const auth = {
    restore: async () => account,
    signIn: async () => account,
    signOut: async () => {},
    ...overrides.auth,
  };
  const outlookClient = {
    list: async () => cloud,
    create: async (d) => {
      writes.push("create");
      const e = { ...d, id: "same-id" };
      cloud = [e];
      return e;
    },
    update: async (id, d) => {
      writes.push("update");
      cloud = [{ ...d, id }];
      return cloud[0];
    },
    remove: async () => {
      writes.push("delete");
      cloud = [];
    },
    ...overrides.outlookClient,
  };
  const controller = createCalendarController({
    localStore,
    auth,
    outlookClient,
    today: () => "2026-09-12",
  });
  return { controller, localStore, writes, db };
}
test("CRUD dispatch isolates sources even when local and Outlook event IDs collide", async () => {
  const { controller: c, localStore, writes } = setup();
  await c.start();
  await c.save(draft);
  const remote = await c.save({
    ...draft,
    source: "outlook",
    title: "Cloud plan",
  });
  assert.equal(c.getState().events.length, 2);
  await c.save({ ...remote, title: "Cloud updated" }, remote);
  assert.equal(localStore.get("same-id").title, "Browser plan");
  await c.remove(remote);
  assert.equal(localStore.list().length, 1);
  assert.deepEqual(writes, ["create", "update", "delete"]);
  assert.equal(c.getState().events.length, 1);
});
test("local CRUD works while Outlook is unavailable", async () => {
  const { controller: c, localStore } = setup({
    outlookClient: {
      list: async () => {
        throw new Error("Offline");
      },
    },
  });
  await c.start();
  await c.save(draft);
  assert.equal(localStore.list().length, 1);
  assert.equal(c.getState().events[0].source, "local");
  assert.match(c.getState().syncError, /Offline/);
});
test("source changes, read-only recurrence edits and signed-out Outlook writes are blocked", async () => {
  const { controller: c, writes } = setup();
  await c.start();
  const local = await c.save(draft);
  await assert.rejects(
    c.save({ ...local, source: "outlook" }, local),
    /destination|source/i,
  );
  const rec = { ...draft, id: "rec", source: "outlook", readOnly: true };
  await assert.rejects(c.save(rec, rec), /recurr|read.only/i);
  await c.disconnectOutlook();
  await assert.rejects(c.save({ ...draft, source: "outlook" }), /connect/i);
  assert.deepEqual(writes, []);
});
test("old month results cannot replace a newer calendar view", async () => {
  let resolveOld;
  let first = true;
  const { controller: c } = setup({
    outlookClient: {
      list: async (range) => {
        if (first) {
          first = false;
          return [];
        }
        if (range.start.startsWith("2026-08"))
          return new Promise((r) => (resolveOld = r));
        return [
          {
            ...draft,
            id: "oct",
            source: "outlook",
            start: "2026-10-12T09:00:00",
            end: "2026-10-12T10:00:00",
          },
        ];
      },
    },
  });
  await c.start();
  const old = c.refresh();
  await c.setCursorDate("2026-10-01");
  resolveOld([{ ...draft, id: "sep", source: "outlook" }]);
  await old;
  assert.equal(c.getState().events[0].id, "oct");
});
test("logout invalidates in-flight responses and retains browser events", async () => {
  let resolve;
  let pending = false;
  const { controller: c } = setup({
    outlookClient: {
      list: async () => (pending ? new Promise((r) => (resolve = r)) : []),
    },
  });
  await c.start();
  await c.save(draft);
  pending = true;
  const request = c.refresh();
  await c.disconnectOutlook();
  resolve([{ ...draft, source: "outlook", id: "cloud" }]);
  await request;
  assert.equal(c.getState().account, null);
  assert.equal(c.getState().events.length, 1);
  assert.equal(c.getState().events[0].source, "local");
});
test("successful cloud create is not reported as failed when following refresh fails", async () => {
  let fail = false;
  const { controller: c, writes } = setup({
    outlookClient: {
      list: async () => {
        if (fail) throw new Error("Refresh unavailable");
        return [];
      },
    },
  });
  await c.start();
  fail = true;
  const saved = await c.save({ ...draft, source: "outlook" });
  assert.equal(saved.id, "same-id");
  assert.deepEqual(writes, ["create"]);
  assert.match(c.getState().syncError, /Refresh unavailable/);
});
