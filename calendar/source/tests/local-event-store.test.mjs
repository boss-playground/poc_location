import test from "node:test";
import assert from "node:assert/strict";
import {
  createLocalEventStore,
  STORAGE_KEY,
} from "../src/local-event-store.mjs";
function memoryStorage(initial) {
  const map = new Map(initial);
  return {
    getItem: (k) => map.get(k) ?? null,
    setItem: (k, v) => map.set(k, v),
  };
}
const draft = {
  source: "local",
  title: "My plan",
  start: "2026-09-12T09:00",
  end: "2026-09-12T10:00",
  allDay: false,
};
test("a denied localStorage getter does not prevent construction and is reported by reads", () => {
  const previous = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
  Object.defineProperty(globalThis, "localStorage", {
    configurable: true,
    get() {
      throw new DOMException("Blocked by browser policy", "SecurityError");
    },
  });
  try {
    let store;
    assert.doesNotThrow(() => {
      store = createLocalEventStore();
    });
    assert.throws(() => store.list(), /Could not read stored events/);
    assert.throws(() => store.create(draft), /Could not read stored events/);
  } finally {
    if (previous) Object.defineProperty(globalThis, "localStorage", previous);
    else delete globalThis.localStorage;
  }
});
test("storage denied between read and write reports save failure without losing records", () => {
  const previous = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
  const storage = memoryStorage([[STORAGE_KEY, "[]"]]);
  let accesses = 0;
  Object.defineProperty(globalThis, "localStorage", {
    configurable: true,
    get() {
      if (++accesses > 1)
        throw new DOMException("Blocked by browser policy", "SecurityError");
      return storage;
    },
  });
  try {
    const store = createLocalEventStore(undefined, () => "new-id");
    assert.throws(
      () => store.create(draft),
      /Could not save to browser storage/,
    );
    assert.equal(storage.getItem(STORAGE_KEY), "[]");
  } finally {
    if (previous) Object.defineProperty(globalThis, "localStorage", previous);
    else delete globalThis.localStorage;
  }
});
test("local create/update/delete survives reconstructed store and never touches unrelated keys", () => {
  const storage = memoryStorage([["unrelated", "keep"]]);
  const store = createLocalEventStore(storage, () => "local-1");
  const saved = store.create(draft);
  assert.equal(saved.id, "local-1");
  const reloaded = createLocalEventStore(storage);
  assert.equal(reloaded.list().length, 1);
  assert.equal(
    reloaded.update(saved.id, { ...draft, title: "Updated" }).title,
    "Updated",
  );
  assert.equal(store.get(saved.id).title, "Updated");
  assert.equal(store.delete(saved.id), true);
  assert.equal(reloaded.list().length, 0);
  assert.equal(storage.getItem("unrelated"), "keep");
});
test("local store blocks source changes and missing-record update", () => {
  const store = createLocalEventStore(memoryStorage(), () => "a");
  store.create(draft);
  assert.throws(() => store.create({ ...draft, source: "outlook" }));
  assert.throws(() => store.update("a", { ...draft, source: "outlook" }));
  assert.throws(() => store.update("missing", draft));
  assert.equal(store.get("a").source, "local");
});
test("corrupt storage is reported and preserved rather than replaced by a new event", () => {
  for (const raw of [
    "{broken",
    "{}",
    JSON.stringify([{ ...draft, id: "x", source: "outlook" }]),
  ]) {
    const storage = memoryStorage([[STORAGE_KEY, raw]]);
    const store = createLocalEventStore(storage);
    assert.throws(() => store.list(), /stored|storage/i);
    assert.throws(() => store.create(draft));
    assert.equal(storage.getItem(STORAGE_KEY), raw);
  }
});
test("storage quota failure is surfaced and never reports a successful create", () => {
  const store = createLocalEventStore({
    getItem: () => null,
    setItem: () => {
      throw new Error("QuotaExceededError");
    },
  });
  assert.throws(() => store.create(draft), /storage|save/i);
});
