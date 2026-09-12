import { validateDraft, intersects } from "./calendar-domain.mjs";
export const STORAGE_KEY = "outlook-calendar-poc.local-events.v1";
export function createLocalEventStore(
  storage,
  createId = () => crypto.randomUUID(),
) {
  function read() {
    try {
      const raw = (storage ?? globalThis.localStorage).getItem(STORAGE_KEY);
      if (raw === null) return [];
      const events = JSON.parse(raw);
      if (!Array.isArray(events)) throw new Error();
      const ids = new Set();
      return events.map((event) => {
        if (
          event.source !== "local" ||
          typeof event.id !== "string" ||
          !event.id ||
          ids.has(event.id)
        )
          throw new Error();
        ids.add(event.id);
        return {
          ...validateDraft(event),
          id: event.id,
          lastModified: event.lastModified,
        };
      });
    } catch {
      throw new Error(
        "Could not read stored events. Browser storage may be unavailable or damaged; existing data has been preserved.",
      );
    }
  }
  function write(events) {
    try {
      (storage ?? globalThis.localStorage).setItem(
        STORAGE_KEY,
        JSON.stringify(events),
      );
    } catch {
      throw new Error(
        "Could not save to browser storage. Check available space and browser privacy settings.",
      );
    }
  }
  function localDraft(draft) {
    const valid = validateDraft(draft);
    if (valid.source !== "local")
      throw new Error("This store only accepts local events.");
    return valid;
  }
  return {
    list(range) {
      return read().filter((event) => !range || intersects(event, range));
    },
    get(id) {
      return read().find((event) => event.id === id) ?? null;
    },
    create(draft) {
      const events = read();
      const event = {
        ...localDraft(draft),
        id: createId(),
        lastModified: new Date().toISOString(),
      };
      events.push(event);
      write(events);
      return event;
    },
    update(id, draft) {
      const events = read();
      const index = events.findIndex((event) => event.id === id);
      if (index < 0)
        throw new Error("This event no longer exists. Refresh the calendar.");
      const event = {
        ...localDraft({ ...events[index], ...draft }),
        id,
        lastModified: new Date().toISOString(),
      };
      events[index] = event;
      write(events);
      return event;
    },
    delete(id) {
      const events = read();
      const remaining = events.filter((event) => event.id !== id);
      if (remaining.length === events.length) return false;
      write(remaining);
      return true;
    },
  };
}
