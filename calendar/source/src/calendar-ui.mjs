import {
  displayDate,
  defaultDraft,
  monthCells,
  monthLabel,
  dayLabel,
  shiftMonth,
  addDays,
  eventsOnDay,
  timeLabel,
} from "./calendar-domain.mjs";

const sourceLabel = (source) =>
  source === "outlook" ? "Outlook" : "This calendar";
function element(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}
export function mountCalendarApp({
  controller,
  onConnect,
  onConfigure,
  clientId = "",
  redirectUri,
}) {
  const $ = (selector) => document.querySelector(selector);
  const dialog = $("#event-dialog"),
    form = $("#event-form"),
    deleteDialog = $("#delete-dialog"),
    setup = $("#setup-dialog"),
    dayDialog = $("#day-dialog");
  let current = null,
    selectedDay = displayDate(),
    snapshot = controller.getState();
  const report = (target, error) => {
    target.textContent = error?.message ?? String(error);
    target.hidden = false;
  };
  function displayFailure(error) {
    controller.setStatus(error.message ?? String(error));
  }
  function chip(event, large = false) {
    const node = element(
      "button",
      `event-chip ${event.source}${large ? " day-list-event" : ""}`,
    );
    node.type = "button";
    node.setAttribute(
      "aria-label",
      `${sourceLabel(event.source)}: ${event.title}, ${timeLabel(event)}`,
    );
    const text = element("span", "event-text");
    text.append(
      element("span", "event-title", event.title),
      element(
        "span",
        "event-meta",
        `${event.allDay ? "All day" : event.start.slice(11, 16)} · ${sourceLabel(event.source)}${event.readOnly ? " · Recurring" : ""}`,
      ),
    );
    node.append(text);
    node.addEventListener("click", (e) => {
      e.stopPropagation();
      if (dayDialog.open) dayDialog.close();
      openEvent(event);
    });
    return node;
  }
  function render(state) {
    snapshot = state;
    $("#month-title").textContent = monthLabel(state.cursorDate);
    $("#mini-title").textContent = monthLabel(state.cursorDate);
    const cells = monthCells(state.cursorDate),
      grid = $("#month-grid"),
      mini = $("#mini-calendar");
    grid.replaceChildren();
    mini.replaceChildren();
    grid.setAttribute(
      "aria-label",
      `${monthLabel(state.cursorDate)} month calendar`,
    );
    for (const weekday of ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"])
      grid.append(element("div", "weekday", weekday));
    const today = displayDate();
    for (const [i, date] of cells.entries()) {
      const outside = date.slice(0, 7) !== state.cursorDate.slice(0, 7),
        isToday = date === today;
      const cell = element(
        "div",
        `day-cell${outside ? " outside" : ""}${isToday ? " is-today" : ""}${i % 7 >= 5 ? " weekend" : ""}`,
      );
      const heading = element("div", "day-heading");
      const dateButton = element(
        "button",
        "date-button",
        String(Number(date.slice(8))),
      );
      dateButton.type = "button";
      dateButton.setAttribute("aria-label", `Add event on ${dayLabel(date)}`);
      if (isToday) dateButton.setAttribute("aria-current", "date");
      dateButton.addEventListener("click", (e) => {
        e.stopPropagation();
        openEvent(null, date);
      });
      heading.append(dateButton);
      cell.append(heading);
      const events = eventsOnDay(state.events, date);
      for (const event of events.slice(0, 2)) cell.append(chip(event));
      if (events.length > 2) {
        const more = element(
          "button",
          "more-events",
          `+${events.length - 2} more`,
        );
        more.type = "button";
        more.setAttribute(
          "aria-label",
          `Show all ${events.length} events on ${dayLabel(date)}`,
        );
        more.addEventListener("click", (e) => {
          e.stopPropagation();
          openDay(date);
        });
        cell.append(more);
      }
      cell.addEventListener("click", (e) => {
        if (e.target === cell) openEvent(null, date);
      });
      grid.append(cell);
      const miniButton = element(
        "button",
        `mini-date${outside ? " outside" : ""}${isToday ? " current" : ""}${date === selectedDay ? " selected" : ""}`,
        String(Number(date.slice(8))),
      );
      miniButton.type = "button";
      miniButton.setAttribute("aria-label", `View ${dayLabel(date)}`);
      miniButton.addEventListener("click", () => openDay(date));
      mini.append(miniButton);
    }
    $("#show-local").checked = state.showLocal;
    $("#show-outlook").checked = state.showOutlook;
    $("#show-outlook").disabled = !state.account;
    $("#local-count").textContent = state.events.filter(
      (e) => e.source === "local",
    ).length;
    $("#outlook-count").textContent = state.events.filter(
      (e) => e.source === "outlook",
    ).length;
    $("#month-summary").textContent = state.events.length
      ? `${state.events.length} event${state.events.length === 1 ? "" : "s"} in view · Bangkok time`
      : "A clear month ahead. Make room for what matters.";
    $("#connect-card").hidden = !!state.account;
    $("#account-card").hidden = !state.account;
    $("#account-name").textContent = state.account?.username ?? "";
    $("#header-connection").textContent = state.account
      ? "Outlook connected"
      : "Personal workspace";
    $("#connection-dot").classList.toggle("connected", !!state.account);
    $("#status").textContent = state.syncing
      ? "Refreshing Outlook calendar…"
      : state.status;
    $("#sync-time").textContent = state.lastSynced
      ? `Last synced ${new Intl.DateTimeFormat("en-GB", { timeZone: "Asia/Bangkok", hour: "2-digit", minute: "2-digit" }).format(new Date(state.lastSynced))} · Manual sync`
      : "Click a day to add an event";
    const errors = [state.localError, state.syncError].filter(Boolean);
    $("#error-banner").hidden = !errors.length;
    $("#error-banner").textContent = errors.join(" ");
    $("#refresh").disabled = !state.account || state.syncing || state.busy;
    $("#connect-outlook").disabled = state.busy;
    $("#disconnect-outlook").disabled = state.busy;
    for (const id of [
      "save-event",
      "delete-event",
      "confirm-delete",
      "new-event",
      "prev-month",
      "next-month",
      "today",
    ])
      $(`#${id}`).disabled = state.busy;
    if (dialog.open) updateEditor();
  }
  function updateEditor() {
    const allDay = $("#event-all-day").checked,
      readOnly = !!current?.readOnly;
    $(".date-fields").classList.toggle("all-day", allDay);
    $("#all-day-hint").hidden = !allDay;
    document
      .querySelectorAll(".time-field")
      .forEach((node) => (node.hidden = allDay));
    for (const input of [$("#event-start-time"), $("#event-end-time")])
      input.required = !allDay;
    for (const input of form.querySelectorAll("input"))
      input.disabled = readOnly || snapshot.busy;
    for (const input of form.querySelectorAll('[name="source"]'))
      input.disabled =
        !!current ||
        snapshot.busy ||
        (input.value === "outlook" && !snapshot.account);
    $("#save-event").hidden = readOnly;
    $("#delete-event").hidden = !current || readOnly;
    const source = form.elements.source.value;
    $("#destination-hint").textContent =
      source === "outlook"
        ? "Changes will be saved to your Outlook calendar."
        : snapshot.account
          ? "Saved only in this browser. Outlook will not be changed."
          : "Saved only in this browser. Connect Outlook to choose it as a destination.";
    const notice = $("#meeting-notice");
    notice.hidden = !readOnly && !current?.isMeeting;
    notice.textContent = readOnly
      ? "This is a recurring Outlook event. Manage this occurrence or series in Outlook."
      : current?.isMeeting
        ? "This Outlook event includes attendees. Changes or deletion may send meeting updates or cancellations."
        : "";
  }
  function openEvent(event = null, date = displayDate()) {
    current = event;
    selectedDay = date;
    const values = event ?? defaultDraft(date);
    form.reset();
    $("#form-error").hidden = true;
    $("#event-dialog-title").textContent = event
      ? event.readOnly
        ? "Event details"
        : "Edit event"
      : "New event";
    $("#event-title").value = values.title;
    form.elements.source.value = values.source;
    $("#event-all-day").checked = values.allDay;
    $("#event-start-date").value = values.start.slice(0, 10);
    $("#event-end-date").value = values.allDay
      ? addDays(values.end, -1)
      : values.end.slice(0, 10);
    $("#event-start-time").value = values.allDay
      ? "09:00"
      : values.start.slice(11, 19);
    $("#event-end-time").value = values.allDay
      ? "10:00"
      : values.end.slice(11, 19);
    updateEditor();
    dialog.showModal();
    if (!event?.readOnly) $("#event-title").focus();
  }
  function openDay(date) {
    selectedDay = date;
    $("#day-dialog-title").textContent = dayLabel(date);
    const list = $("#day-events");
    list.replaceChildren();
    const events = eventsOnDay(snapshot.events, date);
    if (!events.length)
      list.append(element("p", "empty-day", "Nothing planned yet."));
    else for (const event of events) list.append(chip(event, true));
    dayDialog.showModal();
  }
  function closeEvent() {
    if (!snapshot.busy) dialog.close();
  }
  $("#new-event").addEventListener("click", () =>
    openEvent(
      null,
      snapshot.cursorDate.slice(0, 7) === displayDate().slice(0, 7)
        ? displayDate()
        : `${snapshot.cursorDate.slice(0, 7)}-01`,
    ),
  );
  $("#close-event").addEventListener("click", closeEvent);
  $("#cancel-event").addEventListener("click", closeEvent);
  dialog.addEventListener("cancel", (event) => {
    if (snapshot.busy) event.preventDefault();
  });
  form.addEventListener("change", updateEditor);
  $("#event-all-day").addEventListener("change", () => {
    if (
      $("#event-all-day").checked &&
      $("#event-end-date").value < $("#event-start-date").value
    )
      $("#event-end-date").value = $("#event-start-date").value;
    updateEditor();
  });
  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    $("#form-error").hidden = true;
    const allDay = $("#event-all-day").checked;
    const draft = {
      source: form.elements.source.value,
      title: $("#event-title").value,
      allDay,
      start: allDay
        ? $("#event-start-date").value
        : `${$("#event-start-date").value}T${$("#event-start-time").value}`,
      end: allDay
        ? addDays($("#event-end-date").value, 1)
        : `${$("#event-end-date").value}T${$("#event-end-time").value}`,
    };
    try {
      await controller.save(draft, current);
      dialog.close();
    } catch (error) {
      report($("#form-error"), error);
    }
  });
  $("#delete-event").addEventListener("click", () => {
    if (!current) return;
    $("#delete-error").hidden = true;
    $("#delete-description").textContent =
      `Delete “${current.title}” from ${current.source === "local" ? "this browser" : "Outlook"}?${current.isMeeting ? " Meeting attendees may receive a cancellation." : ""}`;
    deleteDialog.showModal();
  });
  $("#cancel-delete").addEventListener("click", () => {
    if (!snapshot.busy) deleteDialog.close();
  });
  deleteDialog.addEventListener("cancel", (e) => {
    if (snapshot.busy) e.preventDefault();
  });
  $("#confirm-delete").addEventListener("click", async () => {
    try {
      await controller.remove(current);
      deleteDialog.close();
      dialog.close();
    } catch (error) {
      report($("#delete-error"), error);
    }
  });
  $("#prev-month").addEventListener("click", () =>
    controller.setCursorDate(shiftMonth(snapshot.cursorDate, -1)),
  );
  $("#next-month").addEventListener("click", () =>
    controller.setCursorDate(shiftMonth(snapshot.cursorDate, 1)),
  );
  $("#today").addEventListener("click", () =>
    controller.setCursorDate(displayDate()),
  );
  $("#show-local").addEventListener("change", (e) =>
    controller.setVisible("local", e.target.checked),
  );
  $("#show-outlook").addEventListener("change", (e) =>
    controller.setVisible("outlook", e.target.checked),
  );
  $("#refresh").addEventListener("click", () => controller.refresh());
  $("#connect-outlook").addEventListener("click", async () => {
    try {
      if (!clientId) {
        openSetup();
        return;
      }
      await onConnect();
    } catch (error) {
      displayFailure(error);
    }
  });
  $("#disconnect-outlook").addEventListener("click", async () => {
    try {
      await controller.disconnectOutlook();
    } catch (error) {
      displayFailure(error);
    }
  });
  function openSetup() {
    $("#client-id").value = clientId;
    $("#redirect-uri").textContent = redirectUri;
    $("#logout-uri").textContent = new URL("logout.html", redirectUri).href;
    $("#setup-error").hidden = true;
    setup.showModal();
  }
  $("#setup-trigger").addEventListener("click", openSetup);
  $("#close-setup").addEventListener("click", () => setup.close());
  $("#setup-form").addEventListener("submit", async (event) => {
    event.preventDefault();
    $("#apply-config").disabled = true;
    try {
      await onConfigure($("#client-id").value.trim());
    } catch (error) {
      report($("#setup-error"), error);
      $("#apply-config").disabled = false;
    }
  });
  $("#close-day").addEventListener("click", () => dayDialog.close());
  $("#add-day-event").addEventListener("click", () => {
    dayDialog.close();
    openEvent(null, selectedDay);
  });
  window.addEventListener("storage", (event) => {
    if (event.key === "outlook-calendar-poc.local-events.v1")
      controller.refresh();
  });
  controller.subscribe(render);
  return { openSetup };
}
