const BATCH_SIZE  = 5;    // calendar API calls per batch
const BATCH_SLEEP = 1000; // ms to sleep between batches

// Returns an array of planned actions without touching the calendar.
// Each item: { action, rowNum, label, detail }
function previewSync() {
  const sheetName  = PROPS.getProperty("sheetName")  || "Sheet1";
  const calendarId = PROPS.getProperty("calendarId") || "primary";

  const sheet    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const calendar = CalendarApp.getCalendarById(calendarId);

  if (!sheet)    throw new Error(`Sheet "${sheetName}" not found`);
  if (!calendar) throw new Error(`Calendar "${calendarId}" not found or not accessible`);

  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const allValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers   = allValues[0].map(h => String(h).trim());

  const idx = buildIdx(headers);
  if (idx.date  < 0) throw new Error("Date column is not mapped.");
  if (idx.title < 0) throw new Error("Event Title column is not mapped.");

  const actions = [];

  allValues.slice(1).forEach((row, i) => {
    const rowNum  = i + 2;
    const dateVal = idx.date  >= 0 ? row[idx.date]  : null;
    const title   = idx.title >= 0 ? String(row[idx.title]).trim() : "";
    const eventId = idx.eventId >= 0 ? String(row[idx.eventId]).trim() : "";

    if (!dateVal && !title) return;

    if (!dateVal || !title) {
      const missing = [!dateVal && "date", !title && "title"].filter(Boolean).join(" & ");
      actions.push({ action: "skip", rowNum, label: `Row ${rowNum}`, detail: `Missing ${missing}` });
      return;
    }

    const date       = new Date(dateVal);
    const dateStr    = date.toLocaleDateString();
    const isAllDay   = !(idx.startTime >= 0 && row[idx.startTime]);
    const timeDetail = isAllDay ? "all day" : formatTimeRange(row, idx);

    if (!eventId) {
      actions.push({ action: "create", rowNum, label: title, detail: `${dateStr} · ${timeDetail}` });
      return;
    }

    let existing;
    try { existing = calendar.getEventById(eventId); } catch (e) { existing = null; }

    if (!existing) {
      actions.push({ action: "create", rowNum, label: title, detail: `${dateStr} · ${timeDetail} (re-create — event was deleted)` });
      return;
    }

    const changes = diffEvent(existing, title, date, isAllDay, row, idx);
    if (changes.length > 0) {
      actions.push({ action: "update", rowNum, label: title, detail: `${dateStr} · ${changes.join(", ")}` });
    } else {
      actions.push({ action: "none", rowNum, label: title, detail: `${dateStr} · no changes` });
    }
  });

  return actions;
}

// Executes only the rows in the saved plan, in batches.
// Returns result objects: { action, label, detail }
function executePlan() {
  const planJson = PROPS.getProperty("syncPlan");
  if (!planJson) throw new Error("No sync plan found. Please run Preview first.");

  const plan = JSON.parse(planJson);
  const work  = plan.filter(a => a.action === "create" || a.action === "update");
  if (work.length === 0) return [];

  const sheetName  = PROPS.getProperty("sheetName")  || "Sheet1";
  const calendarId = PROPS.getProperty("calendarId") || "primary";

  const sheet    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const calendar = CalendarApp.getCalendarById(calendarId);

  if (!sheet)    throw new Error(`Sheet "${sheetName}" not found`);
  if (!calendar) throw new Error(`Calendar "${calendarId}" not found or not accessible`);

  const lastCol   = sheet.getLastColumn();
  const allValues = sheet.getRange(1, sheet.getLastRow(), 1, lastCol).getValues(); // header only first
  const headers   = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const idx       = buildIdx(headers);

  const results = [];

  for (let b = 0; b < work.length; b += BATCH_SIZE) {
    if (b > 0) Utilities.sleep(BATCH_SLEEP);

    const batch = work.slice(b, b + BATCH_SIZE);

    batch.forEach(planned => {
      const row      = sheet.getRange(planned.rowNum, 1, 1, lastCol).getValues()[0];
      const dateVal  = idx.date  >= 0 ? row[idx.date]  : null;
      const title    = idx.title >= 0 ? String(row[idx.title]).trim() : "";
      const eventId  = idx.eventId >= 0 ? String(row[idx.eventId]).trim() : "";
      const date     = new Date(dateVal);
      const dateStr  = date.toLocaleDateString();
      const isAllDay = !(idx.startTime >= 0 && row[idx.startTime]);
      const location = idx.location >= 0 ? String(row[idx.location] || "").trim() : "";
      const notes    = idx.notes    >= 0 ? String(row[idx.notes]    || "").trim() : "";
      const options  = {};
      if (location) options.location    = location;
      if (notes)    options.description = notes;

      try {
        if (planned.action === "create" || !eventId) {
          if (idx.eventId >= 0) sheet.getRange(planned.rowNum, idx.eventId + 1).setValue("");
          createEvent(calendar, sheet, planned.rowNum, date, title, isAllDay, row, idx, options);
          results.push({ action: "created", label: planned.label, detail: dateStr });
        } else {
          const outcome = updateEvent(calendar, sheet, planned.rowNum, eventId, date, title, isAllDay, row, idx, options);
          results.push({ action: outcome, label: planned.label, detail: dateStr });
        }
      } catch (e) {
        results.push({ action: "failed", label: planned.label, detail: `${dateStr} — ${e.message}` });
      }
    });
  }

  PROPS.deleteProperty("syncPlan");
  return results;
}

// ── Diff helpers ──────────────────────────────────────────────────────────────

function diffEvent(event, newTitle, newDate, isAllDay, row, idx) {
  const changes = [];
  if (event.getTitle() !== newTitle) changes.push("title");

  const newLoc  = idx.location >= 0 ? String(row[idx.location] || "").trim() : "";
  const newDesc = idx.notes    >= 0 ? String(row[idx.notes]    || "").trim() : "";
  if (event.getLocation()    !== newLoc)  changes.push("location");
  if (event.getDescription() !== newDesc) changes.push("notes");

  if (isAllDay) {
    const existingDate = event.getAllDayStartDate();
    if (!existingDate || existingDate.toDateString() !== newDate.toDateString()) changes.push("date");
  } else {
    const { start, end } = buildDateTimes(newDate, row, idx);
    if (Math.abs(event.getStartTime() - start) > 60000) changes.push("start time");
    if (Math.abs(event.getEndTime()   - end)   > 60000) changes.push("end time");
  }
  return changes;
}

function formatTimeRange(row, idx) {
  const startVal = idx.startTime >= 0 ? row[idx.startTime] : null;
  const endVal   = idx.endTime   >= 0 ? row[idx.endTime]   : null;
  const fmt = t => t instanceof Date
    ? t.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })
    : String(t);
  return startVal ? (endVal ? `${fmt(startVal)}–${fmt(endVal)}` : fmt(startVal)) : "";
}

// ── Cards ─────────────────────────────────────────────────────────────────────

function buildPreviewCard(actions) {
  // Save the plan so Confirm can execute it without re-diffing
  const plan = actions.filter(a => a.action === "create" || a.action === "update")
                      .map(({ action, rowNum, label }) => ({ action, rowNum, label }));
  PROPS.setProperty("syncPlan", JSON.stringify(plan));

  const counts = { create: 0, update: 0, none: 0, skip: 0 };
  actions.forEach(a => counts[a.action] = (counts[a.action] || 0) + 1);

  const summary = [
    counts.create && `${counts.create} to create`,
    counts.update && `${counts.update} to update`,
    counts.skip   && `${counts.skip} skipped`,
    counts.none   && `${counts.none} unchanged`,
  ].filter(Boolean).join("  ·  ") || "Nothing to sync";

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle("Sync preview")
      .setSubtitle(summary));

  const SECTIONS = [
    { key: "create", heading: "Will create",  icon: CardService.Icon.INVITE      },
    { key: "update", heading: "Will update",  icon: CardService.Icon.DESCRIPTION },
    { key: "skip",   heading: "Skipped rows", icon: CardService.Icon.MEMBERSHIP  },
    { key: "none",   heading: "No changes",   icon: CardService.Icon.BOOKMARK    },
  ];

  SECTIONS.forEach(({ key, heading, icon }) => {
    const items = actions.filter(a => a.action === key);
    if (items.length === 0) return;

    const section = CardService.newCardSection()
      .setHeader(heading)
      .setCollapsible(key === "none" || key === "skip");

    items.forEach(({ label, detail }) => {
      section.addWidget(
        CardService.newDecoratedText()
          .setTopLabel(label)
          .setText(detail)
          .setStartIcon(CardService.newIconImage().setIcon(icon))
      );
    });
    card.addSection(section);
  });

  if (actions.length === 0) {
    card.addSection(CardService.newCardSection().addWidget(
      CardService.newTextParagraph().setText("No rows found to sync.")
    ));
  }

  const nothingToDo = counts.create === 0 && counts.update === 0;
  card.addSection(CardService.newCardSection().addWidget(
    CardService.newTextButton()
      .setText(`Confirm & sync (${plan.length} events, ~${Math.ceil(plan.length / BATCH_SIZE)} batch${Math.ceil(plan.length / BATCH_SIZE) === 1 ? "" : "es"})`)
      .setOnClickAction(CardService.newAction().setFunctionName("confirmSyncFromPreview"))
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setDisabled(nothingToDo)
  ));

  return card.build();
}

function confirmSyncFromPreview() {
  let results;
  try {
    results = executePlan();
  } catch (ex) {
    return notify(`Sync error: ${ex.message}`);
  }

  const ts = new Date().toLocaleString();
  PROPS.setProperty("lastSync", ts);

  const nav = CardService.newNavigation().updateCard(buildResultsCard(results, ts));
  return CardService.newActionResponseBuilder().setNavigation(nav).build();
}

function buildResultsCard(results, ts) {
  const counts = {};
  results.forEach(r => { counts[r.action] = (counts[r.action] || 0) + 1; });

  const hasFailed = (counts.failed || 0) > 0;
  const summary = [
    counts.created && `${counts.created} created`,
    counts.updated && `${counts.updated} updated`,
    counts.failed  && `${counts.failed} failed`,
  ].filter(Boolean).join("  ·  ") || "Nothing synced";

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle(hasFailed ? "Sync completed with errors" : "Sync complete")
      .setSubtitle(summary));

  const SECTIONS = [
    { key: "failed",  heading: "Failed",  icon: CardService.Icon.STAR        },
    { key: "created", heading: "Created", icon: CardService.Icon.INVITE      },
    { key: "updated", heading: "Updated", icon: CardService.Icon.DESCRIPTION },
  ];

  SECTIONS.forEach(({ key, heading, icon }) => {
    const items = results.filter(r => r.action === key);
    if (items.length === 0) return;

    const section = CardService.newCardSection()
      .setHeader(`${heading} (${items.length})`)
      .setCollapsible(key !== "failed");

    items.forEach(({ label, detail }) => {
      section.addWidget(
        CardService.newDecoratedText()
          .setTopLabel(label)
          .setText(detail)
          .setStartIcon(CardService.newIconImage().setIcon(icon))
      );
    });
    card.addSection(section);
  });

  card.addSection(CardService.newCardSection()
    .addWidget(CardService.newDecoratedText().setTopLabel("Completed at").setText(ts))
    .addWidget(CardService.newTextButton()
      .setText("Done")
      .setOnClickAction(CardService.newAction().setFunctionName("goHome"))
      .setTextButtonStyle(CardService.TextButtonStyle.OUTLINED)
    )
  );

  return card.build();
}

function goHome() {
  const nav = CardService.newNavigation().popToRoot().updateCard(buildMainCard());
  return CardService.newActionResponseBuilder().setNavigation(nav).build();
}
