import {
  displayDate,
  gridRange,
  validateDraft,
  intersects,
} from "./calendar-domain.mjs";

export function createCalendarController({
  localStore,
  outlookClient,
  auth,
  today = displayDate,
}) {
  let state = {
    cursorDate: today(),
    account: null,
    events: [],
    showLocal: true,
    showOutlook: true,
    syncing: false,
    busy: false,
    status: "Ready when you are.",
    localError: "",
    syncError: "",
    lastSynced: null,
  };
  let local = [],
    remote = [],
    generation = 0;
  const listeners = new Set();
  const getState = () => ({
    ...state,
    events: state.events.map((e) => ({ ...e })),
  });
  function emit() {
    state.events = [
      ...(state.showLocal ? local : []),
      ...(state.showOutlook ? remote : []),
    ].sort(
      (a, b) =>
        a.start.localeCompare(b.start) || a.title.localeCompare(b.title),
    );
    for (const fn of listeners) fn(getState());
  }
  function readLocal() {
    try {
      local = localStore.list(gridRange(state.cursorDate));
      state.localError = "";
    } catch (error) {
      state.localError = error.message;
    }
  }
  async function refresh() {
    const request = ++generation,
      accountId = state.account?.homeAccountId;
    readLocal();
    emit();
    if (!state.account || !outlookClient) {
      state.syncing = false;
      remote = [];
      emit();
      return;
    }
    state.syncing = true;
    emit();
    try {
      const events = await outlookClient.list(gridRange(state.cursorDate));
      if (request !== generation || accountId !== state.account?.homeAccountId)
        return;
      remote = events;
      state.syncError = "";
      state.lastSynced = new Date().toISOString();
    } catch (error) {
      if (request === generation) state.syncError = error.message;
    } finally {
      if (request === generation) {
        state.syncing = false;
        emit();
      }
    }
  }
  function checkWritable(draft, original) {
    if (original && original.source !== draft.source)
      throw new Error("An existing event cannot change its destination.");
    if (original?.readOnly)
      throw new Error(
        "Recurring events are read-only here. Edit this event in Outlook.",
      );
    if (draft.source === "outlook" && (!state.account || !outlookClient))
      throw new Error("Connect Outlook before saving an Outlook event.");
  }
  async function save(draft, original = null) {
    if (state.busy) throw new Error("An operation is already in progress.");
    const valid = validateDraft(draft);
    checkWritable(valid, original);
    const identity = state.account?.homeAccountId;
    state.busy = true;
    emit();
    try {
      let saved;
      if (valid.source === "local")
        saved = original
          ? localStore.update(original.id, valid)
          : localStore.create(valid);
      else
        saved = original
          ? await outlookClient.update(original.id, {
              ...valid,
              graphTimeZone: original.graphTimeZone,
            })
          : await outlookClient.create(valid);
      state.status =
        valid.source === "local"
          ? "Saved in this browser."
          : "Saved to Outlook.";
      if (valid.source === "local") {
        readLocal();
        emit();
      } else if (identity === state.account?.homeAccountId) {
        remote = remote.filter((e) => e.id !== saved.id);
        if (intersects(saved, gridRange(state.cursorDate))) remote.push(saved);
        emit();
        await refresh();
      }
      return saved;
    } finally {
      state.busy = false;
      emit();
    }
  }
  async function remove(event) {
    if (state.busy) throw new Error("An operation is already in progress.");
    checkWritable(event, event);
    state.busy = true;
    emit();
    try {
      if (event.source === "local") {
        localStore.delete(event.id);
        readLocal();
      } else if (event.source === "outlook") {
        await outlookClient.remove(event.id);
        remote = remote.filter((e) => e.id !== event.id);
      } else throw new Error("Unknown event source.");
      state.status =
        event.source === "local"
          ? "Event deleted from this browser."
          : "Event deleted from Outlook.";
      emit();
      if (event.source === "outlook") await refresh();
    } finally {
      state.busy = false;
      emit();
    }
  }
  return {
    getState,
    subscribe(fn) {
      listeners.add(fn);
      fn(getState());
      return () => listeners.delete(fn);
    },
    async start() {
      readLocal();
      emit();
      try {
        state.account = auth ? await auth.restore() : null;
      } catch (error) {
        state.syncError = error.message;
      }
      await refresh();
    },
    refresh,
    save,
    remove,
    async setCursorDate(date) {
      ++generation;
      state.cursorDate = date;
      remote = [];
      local = [];
      state.lastSynced = null;
      state.syncError = "";
      return refresh();
    },
    setVisible(source, value) {
      if (source === "local") state.showLocal = value;
      if (source === "outlook") state.showOutlook = value;
      emit();
    },
    async connectOutlook() {
      if (!auth)
        throw new Error(
          "Set up your Microsoft application before connecting Outlook.",
        );
      if (state.busy) return;
      state.busy = true;
      emit();
      try {
        const account = await auth.signIn();
        ++generation;
        remote = [];
        state.account = account;
        state.syncError = "";
        state.status = "Outlook connected.";
        await refresh();
      } finally {
        state.busy = false;
        emit();
      }
    },
    async disconnectOutlook() {
      if (state.busy)
        throw new Error("Wait for the current operation before disconnecting.");
      ++generation;
      state.account = null;
      remote = [];
      state.syncing = false;
      state.lastSynced = null;
      state.syncError = "";
      state.status = "Outlook disconnected. Browser events are kept.";
      emit();
      if (auth) await auth.signOut();
    },
    setStatus(status) {
      state.status = status;
      emit();
    },
  };
}
